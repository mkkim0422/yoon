"""
pdf_export.py — Excel Invoice 시트를 PDF로 변환.

구현 방식: 별도 Python subprocess 를 띄워 그 안에서 Excel COM 자동화를
수행한다. Streamlit 은 요청을 워커 스레드에서 처리하므로 같은 프로세스에서
COM 을 초기화하면 apartment threading / GIL / 전역 Excel 상태 공유 등으로
비결정적 실패가 잦다. subprocess 로 격리하면 COM 상태가 매번 새로 초기화
되고, 실패 시 stderr 를 그대로 캡처해 사용자에게 보여줄 수 있다.

Windows + Microsoft Excel + pywin32 가 필요.
"""
from __future__ import annotations

import os
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
    """PDF 변환 가능 여부 — Windows 에서 subprocess 로 Excel dispatch 를 시도해 판정."""
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
