"""
pdf_export.py — Excel Invoice 시트를 PDF로 변환.

구현 방식 — OS 분기:
  · Windows: Excel COM (pywin32). Streamlit 워커 스레드의 COM apartment 문제를
    피하려고 별도 Python subprocess 에서 Dispatch.
  · Linux  : LibreOffice headless (`soffice --convert-to pdf`). Streamlit Cloud
    같은 Linux 환경 대응. 변환 전에 openpyxl 로 sheet_name 외 시트는 hidden
    처리하고 A4·여백·"한 페이지에 맞춤" 설정을 주입해 Excel COM 과 가능한
    가까운 레이아웃을 얻는다.

요구 사항:
  · Windows: Microsoft Excel + pywin32
  · Linux  : libreoffice (apt: `libreoffice-calc`, `libreoffice-core`)
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path


# subprocess 안에서 실행될 Excel COM 스크립트.
# xlsx → PDF 변환만 수행; 결과/오류는 stdout/stderr 로 보고.
_CHILD_SCRIPT = r'''
import sys, os, traceback
try:
    import pythoncom
    import win32com.client as win32
except Exception as e:
    sys.stderr.write(f"pywin32 import failed: {e}\n")
    sys.exit(2)

XLSX_PATH  = sys.argv[1]
PDF_PATH   = sys.argv[2]
SHEET_NAME = sys.argv[3]

XL_TYPE_PDF  = 0
XL_PORTRAIT  = 1
XL_LANDSCAPE = 2
XL_STANDARD  = 0
XL_PAPER_A4  = 9

excel = None
wb    = None
try:
    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx("Excel.Application")
    except Exception as e:
        sys.stderr.write(f"Excel dispatch failed: {e}\n")
        sys.exit(3)

    excel.Visible       = False
    excel.DisplayAlerts = False
    try:
        excel.AskToUpdateLinks   = False
        excel.AutomationSecurity = 3
    except Exception:
        pass

    wb = excel.Workbooks.Open(os.path.abspath(XLSX_PATH),
                              ReadOnly=True, UpdateLinks=0)
    try:
        excel.CalculateFull()
    except Exception:
        pass

    try:
        ws = wb.Worksheets(SHEET_NAME)
    except Exception as e:
        sys.stderr.write(f"Worksheet '{SHEET_NAME}' not found: {e}\n")
        sys.exit(4)

    try:
        # 세로 A4, 가로/세로 모두 1페이지에 맞춤 (첨부 샘플 PDF 동일 레이아웃)
        ws.PageSetup.PaperSize          = XL_PAPER_A4
        ws.PageSetup.Orientation        = XL_PORTRAIT
        ws.PageSetup.Zoom               = False
        ws.PageSetup.FitToPagesWide     = 1
        ws.PageSetup.FitToPagesTall     = 1     # 한 페이지 안에 높이도 맞춤
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.CenterVertically   = False
        # 여백 최소화 — 빈 공간 감소
        ws.PageSetup.LeftMargin   = excel.InchesToPoints(0.25)
        ws.PageSetup.RightMargin  = excel.InchesToPoints(0.25)
        ws.PageSetup.TopMargin    = excel.InchesToPoints(0.3)
        ws.PageSetup.BottomMargin = excel.InchesToPoints(0.3)
        ws.PageSetup.HeaderMargin = excel.InchesToPoints(0.1)
        ws.PageSetup.FooterMargin = excel.InchesToPoints(0.1)
        # 프린트 영역을 실제 콘텐츠 범위로 좁히기 (빈 행으로 인한 추가 페이지 방지)
        try:
            used = ws.UsedRange
            ws.PageSetup.PrintArea = used.Address
        except Exception:
            pass
    except Exception:
        pass

    ws.Select()

    ws.ExportAsFixedFormat(
        Type=XL_TYPE_PDF,
        Filename=os.path.abspath(PDF_PATH),
        Quality=XL_STANDARD,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )

    if not os.path.exists(PDF_PATH) or os.path.getsize(PDF_PATH) == 0:
        sys.stderr.write("Excel returned but PDF was not written.\n")
        sys.exit(5)

    sys.stdout.write("OK\n")
    sys.exit(0)

except SystemExit:
    raise
except Exception:
    sys.stderr.write(traceback.format_exc())
    sys.exit(1)
finally:
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        if excel is not None:
            excel.Quit()
    except Exception:
        pass
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass
'''


def _find_soffice() -> str | None:
    """Linux/Mac 의 LibreOffice 바이너리 경로. PATH 우선, 없으면 흔한 위치 탐색."""
    for cand in ("soffice", "libreoffice"):
        p = shutil.which(cand)
        if p:
            return p
    for p in (
        "/usr/bin/soffice", "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice",
        "/snap/bin/libreoffice",
    ):
        if Path(p).exists():
            return p
    return None


def _xlsx_to_pdf_libreoffice(
    xlsx_bytes: bytes,
    sheet_name: str,
    timeout_sec: int,
) -> tuple[bytes | None, str | None]:
    """LibreOffice headless 로 xlsx → PDF. sheet_name 외 시트는 숨겨서 첫 페이지가
    해당 시트가 되도록 한다. PageSetup(A4·여백·1페이지 맞춤)도 openpyxl 로 주입.
    """
    soffice = _find_soffice()
    if not soffice:
        return None, (
            "LibreOffice(soffice) 를 찾을 수 없습니다. "
            "packages.txt 에 `libreoffice-calc` 가 포함되어 있는지 확인하세요."
        )

    base = Path(tempfile.gettempdir()) / f"sph_pdf_{uuid.uuid4().hex[:10]}"
    base.mkdir(parents=True, exist_ok=True)
    xlsx_path = base / "invoice.xlsx"
    # 출력 PDF 는 soffice 가 입력 파일명 기준 같은 이름으로 만든다(invoice.pdf)
    pdf_path = base / "invoice.pdf"

    try:
        xlsx_path.write_bytes(xlsx_bytes)

        # ── sheet_name 외 시트 숨김 + PageSetup 주입 ──────────────────
        try:
            import openpyxl
            wb = openpyxl.load_workbook(xlsx_path)
            if sheet_name not in wb.sheetnames:
                return None, f"엑셀 파일에 '{sheet_name}' 시트가 없습니다."
            # 대상 시트를 active 로 설정 + 다른 시트는 hidden
            for nm in wb.sheetnames:
                ws = wb[nm]
                if nm == sheet_name:
                    wb.active = wb.sheetnames.index(nm)
                else:
                    ws.sheet_state = "hidden"
            ws = wb[sheet_name]
            # 페이지 설정 — Excel COM 분기와 동일한 의도(A4, 1페이지 fit, 좁은 여백)
            ws.page_setup.orientation     = ws.ORIENTATION_PORTRAIT
            ws.page_setup.paperSize       = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth      = 1
            ws.page_setup.fitToHeight     = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.print_options.horizontalCentered = True
            ws.page_margins.left   = 0.25
            ws.page_margins.right  = 0.25
            ws.page_margins.top    = 0.3
            ws.page_margins.bottom = 0.3
            ws.page_margins.header = 0.1
            ws.page_margins.footer = 0.1
            wb.save(xlsx_path)
        except Exception as e:
            # openpyxl 처리 실패해도 그대로 변환 시도 (품질만 낮아질 뿐).
            sys.stderr.write(f"pre-pdf openpyxl tweak failed: {e}\n")

        # ── LibreOffice headless 변환 ───────────────────────────────
        # HOME 환경변수 없는 컨테이너 환경 대응 — --env:UserInstallation
        user_profile = base / "lo_profile"
        completed = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore", "--nofirststartwizard", "--nologo",
                f"-env:UserInstallation=file://{user_profile}",
                "--convert-to", "pdf",
                "--outdir", str(base),
                str(xlsx_path),
            ],
            capture_output=True,
            timeout=timeout_sec,
        )

        if pdf_path.exists() and pdf_path.stat().st_size > 0:
            return pdf_path.read_bytes(), None

        stderr = (completed.stderr or b"").decode(errors="ignore").strip()
        stdout = (completed.stdout or b"").decode(errors="ignore").strip()
        return None, (
            f"LibreOffice 변환 실패 [rc={completed.returncode}]: "
            f"{stderr or stdout or '(추가 정보 없음)'}"
        )
    except subprocess.TimeoutExpired:
        return None, f"PDF 변환이 {timeout_sec} 초 내에 완료되지 않았습니다."
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"
    finally:
        for p in (xlsx_path, pdf_path):
            try:
                if p.exists(): p.unlink()
            except Exception:
                pass
        try:
            shutil.rmtree(base, ignore_errors=True)
        except Exception:
            pass


def xlsx_sheet_to_pdf(
    xlsx_bytes: bytes,
    sheet_name: str = "Invoice",
    timeout_sec: int = 90,
) -> tuple[bytes | None, str | None]:
    """xlsx 바이트에서 지정 시트를 PDF 바이트로 추출.

    반환:
      (pdf_bytes, None)  — 성공
      (None, err_msg)    — 실패 (err_msg 는 사용자에게 표시할 원인)
    """
    # Linux/macOS → LibreOffice 분기 (Streamlit Cloud 등 Excel 없는 환경 대응).
    if platform.system() != "Windows":
        # LibreOffice 가 무거운 변환이라 기본 timeout 을 좀 더 넉넉히.
        return _xlsx_to_pdf_libreoffice(xlsx_bytes, sheet_name, max(timeout_sec, 120))

    # 임시 작업 디렉터리 — Excel COM 은 한글/공백 경로에 약하므로 %TEMP% 하위 영문 경로 사용
    base = Path(tempfile.gettempdir()) / f"sph_pdf_{uuid.uuid4().hex[:10]}"
    base.mkdir(parents=True, exist_ok=True)
    xlsx_path = base / "invoice.xlsx"
    pdf_path  = base / "invoice.pdf"
    script_path = base / "_conv.py"

    pdf_bytes: bytes | None = None
    err_msg:   str   | None = None

    try:
        xlsx_path.write_bytes(xlsx_bytes)
        script_path.write_text(_CHILD_SCRIPT, encoding="utf-8")

        # subprocess 로 격리 실행 (Streamlit 의 COM apartment 와 완전 분리)
        completed = subprocess.run(
            [sys.executable, str(script_path),
             str(xlsx_path), str(pdf_path), sheet_name],
            capture_output=True,
            timeout=timeout_sec,
            # Windows 에서 콘솔 창이 뜨지 않도록
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

        if completed.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0:
            pdf_bytes = pdf_path.read_bytes()
        else:
            stderr = (completed.stderr or b"").decode(errors="ignore").strip()
            stdout = (completed.stdout or b"").decode(errors="ignore").strip()
            rc = completed.returncode
            hint = {
                2: "pywin32 가 설치되지 않았거나 COM 등록이 필요합니다 "
                   "(`python -m pip install --upgrade pywin32` 후 "
                   "`python Scripts/pywin32_postinstall.py -install`).",
                3: "Microsoft Excel 이 설치되지 않았거나 현재 계정에서 실행 불가합니다.",
                4: f"엑셀 파일에 '{sheet_name}' 시트가 존재하지 않습니다.",
                5: "Excel 이 PDF 를 생성하지 못했습니다 (프린터 드라이버 확인).",
            }.get(rc, "원인 불명")
            err_msg = f"[rc={rc}] {hint}\n{stderr or stdout or '(추가 정보 없음)'}"

    except subprocess.TimeoutExpired:
        err_msg = f"PDF 변환이 {timeout_sec} 초 내에 완료되지 않았습니다."
    except FileNotFoundError as e:
        err_msg = f"Python 실행 파일을 찾지 못했습니다: {e}"
    except Exception as e:
        err_msg = f"{type(e).__name__}: {e}"
    finally:
        # 임시 파일 정리
        for p in (xlsx_path, pdf_path, script_path):
            try:
                if p.exists(): p.unlink()
            except Exception:
                pass
        try:
            os.rmdir(base)
        except Exception:
            pass

    return pdf_bytes, err_msg


def is_available() -> bool:
    """PDF 변환 가능 여부.
       Windows: subprocess 로 Excel dispatch 시도.
       Linux/macOS: LibreOffice(soffice) 바이너리 존재 확인.
    """
    if platform.system() != "Windows":
        return _find_soffice() is not None

    script = (
        "import sys\n"
        "try:\n"
        "    import pythoncom, win32com.client as w\n"
        "    pythoncom.CoInitialize()\n"
        "    e = w.DispatchEx('Excel.Application')\n"
        "    e.Quit()\n"
        "    print('OK')\n"
        "except Exception as x:\n"
        "    sys.stderr.write(str(x))\n"
        "    sys.exit(1)\n"
    )
    try:
        r = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, timeout=15,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return r.returncode == 0
    except Exception:
        return False
