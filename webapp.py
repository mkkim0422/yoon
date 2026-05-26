"""
webapp.py — SPH GMP 정산 자동화 시스템 v2

streamlit run webapp.py
"""
from __future__ import annotations

import hashlib
import json
import tempfile
from datetime import date
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st
from streamlit_sortables import sort_items

from billing.engine import calculate_billing, calculate_billing_by_project
from billing.loader import (
    build_sku_master_from_usage,
    detect_missing_skus,
    detect_price_list_currency,
    get_billable_sku_names,
    get_free_caps_from_price_list,
    load_usage_rows,
)
from billing.preprocessor import extract_company_names, preprocess_usage_file
from invoice_generator import generate_formatted_invoice, validate_invoice_excel
import github_storage

# ── 경로 상수 ─────────────────────────────────────────────────────────────────
MASTER_CSV             = Path(__file__).parent / "billing" / "master_data.csv"
PRICE_LIST_SAVED       = Path(__file__).parent / "billing" / "saved_price_list.xlsx"       # 레거시
PRICE_LIST_SAVED_USD   = Path(__file__).parent / "billing" / "saved_price_list_usd.xlsx"
PRICE_LIST_SAVED_KRW   = Path(__file__).parent / "billing" / "saved_price_list_krw.xlsx"
SAVED_ORDERS_FILE      = Path(__file__).parent / "billing" / "saved_orders.json"
SAVED_BILLING_MODE_FILE = Path(__file__).parent / "billing" / "saved_billing_mode.json"
SAVED_MIN_CHARGE_FILE   = Path(__file__).parent / "billing" / "saved_min_charge.json"
SAVED_RATE_LABEL_FILE   = Path(__file__).parent / "billing" / "saved_rate_label.json"
SAVED_INCLUDE_PROJECT_FILE = Path(__file__).parent / "billing" / "saved_include_project.json"
SAVED_SUBTOTAL_ROUND_FILE  = Path(__file__).parent / "billing" / "saved_subtotal_round.json"
SAVED_HIDDEN_SKUS_FILE     = Path(__file__).parent / "billing" / "saved_hidden_skus.json"
SAVED_MANUAL_SKUS_FILE     = Path(__file__).parent / "billing" / "saved_manual_skus.json"
SAVED_BATCH_SELECTION_FILE = Path(__file__).parent / "billing" / "saved_batch_selection.json"
SAVED_COMPANY_NOTES_FILE   = Path(__file__).parent / "billing" / "saved_company_notes.json"

# 일괄 정산 "⭐ 즐겨찾기" 버튼 클릭 시 체크되는 회사 명단 (임시 하드코딩).
# 검색 무관 전체 적용. 1인 사용 환경 가정으로 별도 JSON 저장 없이 코드에 박음.
# 명단 변경 필요 시 이 리스트만 수정. _norm_account_key 정규화로 매칭하므로
# 대소문자/공백/하이픈 차이는 자동 흡수.
_BATCH_FAVORITE_COMPANIES: list[str] = [
    "amorepacific", "atomy", "Bespinglobal-Myrealtrip", "Bespinglobal-Socar",
    "BespinGlobal-TeamO2", "Bespinglobal-Vanpl", "Beyless - Monitoring System",
    "coupang", "dogtra", "Ground K", "han-pass", "hanatour", "hankookn",
    "Hecto Innovation", "hyundai-autolink", "hyundaicard-universe", "jch",
    "kakao", "kakao-mobility", "kopri", "koreanair", "lg-dmst", "LG-MCS",
    "lg-uplus", "lotte-card", "megazonesoft-yanolja", "mofa-callcenter",
    "newbalance", "pantos", "pittasoft", "rememberapp", "s1",
    "Samsung Wallet", "samsung-find", "Samsung-Logitech", "samsung-m-gspn",
    "Samsung-NowBrief", "samsung-store", "samsung-visitin",
    "smartthings-find", "SoftEN Corp.", "StudioG", "Timing Golf",
    "triphos", "verygoodtour", "webtour",
]

# GitHub 원격 경로 (Streamlit 휴면 후에도 단가표가 유지되도록 repo 에 백업/복원).
# `.streamlit/secrets.toml` 의 [github] 설정이 없으면 조용히 no-op.
_PRICE_LIST_REMOTES = {
    PRICE_LIST_SAVED_USD: "billing/saved_price_list_usd.xlsx",
    PRICE_LIST_SAVED_KRW: "billing/saved_price_list_krw.xlsx",
    PRICE_LIST_SAVED:     "billing/saved_price_list.xlsx",
}

# 계정별 과금 모드 — "account"(회사 통합 waterfall, 기본) / "per_project"(프로젝트 독립)
BILLING_MODE_ACCOUNT     = "account"
BILLING_MODE_PER_PROJECT = "per_project"

# 월 최소사용비용 기본값 (Google Maps Platform 기본 정책 — ₩500,000).
# 회사마다 계약으로 달라질 수 있어, 계정별로 UI 에서 재설정 가능.
DEFAULT_MIN_CHARGE_AMOUNT   = 0
DEFAULT_MIN_CHARGE_CURRENCY = "KRW"

# 인보이스 환율 표기 — "환율(하나은행 2026.02.27 최종 송금환율 기준)" 조합용.
# 사용자가 UI 에서 자유롭게 바꿀 수 있도록 선택지 + 직접입력 제공.
# 외환·송금환율 고시를 공식적으로 제공하는 국내 주요 은행만 기본 항목으로
# 노출 (인터넷 전문 은행 등 외환 취급 제한 은행은 제외 — 필요 시 직접입력).
MAJOR_BANKS = [
    "하나은행", "국민은행", "신한은행", "우리은행", "농협은행",
    "기업은행", "SC제일은행", "씨티은행",
]
RATE_PHRASES = [
    "매매기준",
    "최종 매매기준율",
    "최종 송금환율 기준",
    "최초 매매기준율",
    "최초고시 매매율 기준",
]
DEFAULT_RATE_PHRASE = "최종 송금환율 기준"
DEFAULT_BANK_NAME   = "하나은행"

# ─── 테스트 모드 (기간 한정 — 아래 False 로 바꾸면 전부 해제됨) ──────────────
# 활성화 시:
#   * 환율 입력값을 1480.80 으로 자동 프리필
# 종료 시: _TEST_DEFAULTS = False 한 줄만 바꾸면 됨.
_TEST_DEFAULTS          = True
_TEST_DEFAULT_RATE      = "1480.80"


# ── tax/VAT SKU 판별 (인보이스 본문에서 제외되는 항목) ─────────────────────
import re as _re
# 세금/VAT 판별 패턴.
#   한글 "세금": substring 매칭 (CSV 상 단독 단어로만 등장)
#   영문 tax / vat: **단어 경계** 기준 — "Elevation" 에 'vat' 이 포함돼 오탐되던
#   버그(Beyless Elevation 16,179 누락) 방지.
_TAX_RE = _re.compile(r"세금|\btax\b|\bvat\b", _re.IGNORECASE)
def _is_tax_sku(name: str) -> bool:
    return bool(_TAX_RE.search(name or ""))


# ── SKU 순서 저장/로드 (계정별) ─────────────────────────────────────────────
# 하드코딩 화이트리스트 없이 저장/로드. 회사마다 실제로 쓰는 SKU 가
# 다르므로, 저장된 순서 중 CSV 에 없는 항목은 나중에 `_existing` 교집합
# 단계에서 자연스럽게 걸러진다(사용 이력이 있는 SKU 만 UI 에 노출됨).
def _load_saved_orders() -> dict[str, list[str]]:
    if not SAVED_ORDERS_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_ORDERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return {
        acc: [n for n in order if isinstance(n, str) and n]
        for acc, order in data.items()
        if isinstance(order, list)
    }


def _load_company_notes() -> dict[str, str]:
    """회사별 메모(비고) 저장값 로드. 일괄 정산 화면 표시 전용 (엑셀 출력 무관)."""
    if not SAVED_COMPANY_NOTES_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_COMPANY_NOTES_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return {
        str(k): str(v) for k, v in (data or {}).items()
        if isinstance(k, str)
    }


def _save_company_note_for_account(account: str, note: str) -> None:
    data = _load_company_notes()
    _v = (note or "").strip()
    if _v:
        data[account] = _v
    else:
        # 빈 입력이면 키 제거 → 다음 로드 시 깨끗.
        data.pop(account, None)
    SAVED_COMPANY_NOTES_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_COMPANY_NOTES_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _save_order_for_account(account: str, order: list[str]) -> None:
    data = _load_saved_orders()
    data[account] = [str(n) for n in order if n]
    SAVED_ORDERS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_ORDERS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _norm_account_key(s) -> str:
    """계정 키 정규화 — 영숫자만 남기고 소문자화.

    GMP.xlsx 표기명을 키로 미리 저장해 둔 회사의 경우, 실제 CSV 결제계정명이
    대소문자/공백/하이픈 등 구분자만 다를 수 있다. 저장된 순서·직접등록
    SKU 를 그런 회사에도 매칭시키기 위한 폴백 비교용.
    """
    return _re.sub(r"[^a-z0-9]+", "", str(s).lower())


def _lookup_account(mapping: dict, account: str):
    """exact 키 우선, 없으면 정규화 일치 키로 폴백. 둘 다 없으면 None."""
    if account in mapping:
        return mapping[account]
    _na = _norm_account_key(account)
    for k, v in mapping.items():
        if _norm_account_key(k) == _na:
            return v
    return None


# ── 계정별 과금 모드 저장/로드 ─────────────────────────────────────────────
# 회사마다 Google 청구 정책이 달라 "회사 통합 waterfall" vs "프로젝트별
# 독립 waterfall" 선호가 다르다. UI 에서 선택한 값을 계정별로 저장해
# 다음 정산 때 자동 로드되도록 한다.
def _load_billing_modes() -> dict[str, str]:
    if not SAVED_BILLING_MODE_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_BILLING_MODE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    valid = {BILLING_MODE_ACCOUNT, BILLING_MODE_PER_PROJECT}
    return {
        acc: v for acc, v in data.items()
        if isinstance(acc, str) and v in valid
    }


def _save_billing_mode_for_account(account: str, mode: str) -> None:
    if mode not in (BILLING_MODE_ACCOUNT, BILLING_MODE_PER_PROJECT):
        return
    data = _load_billing_modes()
    data[account] = mode
    SAVED_BILLING_MODE_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_BILLING_MODE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── 계정별 Project 시트 포함 여부 저장/로드 ─────────────────────────────────
# 일부 회사는 Project(요약) 시트를 사용하지 않는다. 계정별로 선택값을 저장해
# 다음 정산 때 자동 로드되도록 한다. 기본값은 True(포함).
def _load_include_project_flags() -> dict[str, bool]:
    if not SAVED_INCLUDE_PROJECT_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_INCLUDE_PROJECT_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return {
        acc: bool(v) for acc, v in data.items() if isinstance(acc, str)
    }


def _save_include_project_for_account(account: str, value: bool) -> None:
    data = _load_include_project_flags()
    data[account] = bool(value)
    SAVED_INCLUDE_PROJECT_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_INCLUDE_PROJECT_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── 계정별 소계 반올림 자리수 저장/로드 ────────────────────────────────────
# Invoice 시트의 SKU 소계 수식(=ROUND(SUM(I..:I..),N)) 의 N 값.
# 0 (정수) / 2 (소수 두 자리) 만 유효. 미저장이면 통화별 기본값 적용.
# 주의: tier 단가/금액 포맷에는 영향 없음 — ROUND 수식의 자리수만 변경.
def _load_subtotal_round_map() -> dict[str, int]:
    if not SAVED_SUBTOTAL_ROUND_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_SUBTOTAL_ROUND_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    out: dict[str, int] = {}
    for acc, v in data.items():
        if not isinstance(acc, str):
            continue
        try:
            iv = int(v)
        except (TypeError, ValueError):
            continue
        if iv in (0, 2):
            out[acc] = iv
    return out


def _save_subtotal_round_for_account(account: str, value: int) -> None:
    if value not in (0, 2):
        return
    data = _load_subtotal_round_map()
    data[account] = int(value)
    SAVED_SUBTOTAL_ROUND_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_SUBTOTAL_ROUND_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── 계정별 엑셀 미노출 SKU 목록 저장/로드 ─────────────────────────────────
# 사용자가 UI 에서 수동 지정한 "엑셀 출력에서 제외할 SKU" 목록. 엔진 계산
# (waterfall / sku_master / line_items 산출) 은 원본 그대로 수행하고,
# generate_formatted_invoice 호출 직전에 line_items / proj_results /
# per_project_invoices 의 **sku_name 매칭 항목만** 제거한다 → 엔진 결과에는
# 영향 없이 출력물에서만 빠진다.
def _load_hidden_skus_map() -> dict[str, list[str]]:
    if not SAVED_HIDDEN_SKUS_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_HIDDEN_SKUS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    out: dict[str, list[str]] = {}
    for acc, lst in data.items():
        if not isinstance(acc, str):
            continue
        if isinstance(lst, list):
            out[acc] = [str(x) for x in lst if isinstance(x, str) and x.strip()]
    return out


def _save_hidden_skus_for_account(account: str, skus: list[str]) -> None:
    data = _load_hidden_skus_map()
    data[account] = [str(s) for s in skus if isinstance(s, str) and s.strip()]
    SAVED_HIDDEN_SKUS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_HIDDEN_SKUS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# 사용자가 UI 에서 수동으로 "마스터 SKU 목록에서 가져와 강제로 노출"한 항목.
# 현재 CSV 에 사용량이 없는 SKU 도 인보이스에 빈 라인으로 노출하고 싶을 때 사용.
# sku_order 의 가장 마지막에 [직접등록] prefix 로 들어가며, 사용자가 X 버튼으로
# 개별 제거 가능.
def _load_manual_skus_map() -> dict[str, list[str]]:
    if not SAVED_MANUAL_SKUS_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_MANUAL_SKUS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    out: dict[str, list[str]] = {}
    for acc, lst in data.items():
        if not isinstance(acc, str):
            continue
        if isinstance(lst, list):
            out[acc] = [str(x) for x in lst if isinstance(x, str) and x.strip()]
    return out


def _save_manual_skus_for_account(account: str, skus: list[str]) -> None:
    data = _load_manual_skus_map()
    data[account] = [str(s) for s in skus if isinstance(s, str) and s.strip()]
    SAVED_MANUAL_SKUS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_MANUAL_SKUS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── 일괄 정산 선택 상태 저장/로드 ──────────────────────────────────────────
# 일괄 정산 UI 에서 사용자가 체크한 회사 목록을 다음 방문 시에도 복원하고,
# 이전에 한 번이라도 화면에 노출되었던 회사 집합(known)과 비교해 "🆕 신규"
# 마커를 표시한다. 다운로드 옵션(xlsx/pdf)·정책 라디오 값도 함께 저장.
def _load_batch_selection() -> dict:
    default = {
        "selected": [], "known": [],
        "dl_xlsx": True, "dl_pdf": False,
        "policy": "as_is",  # as_is | move | skip
    }
    if not SAVED_BATCH_SELECTION_FILE.exists():
        return default
    try:
        data = json.loads(SAVED_BATCH_SELECTION_FILE.read_text(encoding="utf-8"))
    except Exception:
        return default
    return {
        "selected": [str(x) for x in data.get("selected", []) if isinstance(x, str)],
        "known":    [str(x) for x in data.get("known", [])    if isinstance(x, str)],
        "dl_xlsx":  bool(data.get("dl_xlsx", True)),
        "dl_pdf":   bool(data.get("dl_pdf", False)),
        "policy":   str(data.get("policy", "as_is")),
    }


def _format_billing_timings_line(company: str, timings: dict) -> str:
    """배치 정산 progress bar 아래에 표시할 회사별 단계 소요시간 한 줄.
    예: "han-pass — 데이터 전처리: 0.12초 / 정산 계산: 0.45초 / 엑셀 생성: 1.23초 / 합계: 1.80초"
    PDF 변환 시간은 dl_pdf 가 켜져 있어 > 0 일 때만 표시.
    """
    _t = timings or {}
    _parts = [
        f"데이터 전처리: {float(_t.get('preprocess', 0) or 0):.2f}초",
        f"정산 계산: {float(_t.get('calc', 0) or 0):.2f}초",
        f"엑셀 생성: {float(_t.get('excel', 0) or 0):.2f}초",
    ]
    _pdf = float(_t.get("pdf", 0) or 0)
    if _pdf > 0:
        _parts.append(f"PDF 변환: {_pdf:.2f}초")
    _parts.append(f"합계: {float(_t.get('total', 0) or 0):.2f}초")
    return f"`{company}` — " + " / ".join(_parts)


def _update_ema(prev: float | None, new_value: float, alpha: float = 0.3) -> float:
    """지수가중이동평균 갱신. prev=None 이면 첫 샘플을 그대로 반환."""
    if prev is None:
        return new_value
    return alpha * new_value + (1 - alpha) * prev


def _format_eta_seconds(seconds: float) -> str:
    """남은 시간 사람용 포맷.
    시각적 안정성을 위해 큰 단위에선 버킷(round-to-nearest) 적용 — EMA 와 함께
    8분→10분 같은 1~2분 미세 변동이 같은 버킷으로 흡수되어 표시가 흔들리지 않음.

      < 60초:   '약 N초'                          (정밀)
      1~5분:    '약 N분 N초'                       (1초 단위)
      5~10분:   '약 N분'   (2분 단위 round)        (2,4,6,8,10)
      10~30분:  '약 N분'   (5분 단위 round)        (10,15,20,25,30)
      30분+:    '약 N분'   (10분 단위 round)       (30,40,50,…)
    """
    if seconds is None or seconds <= 0:
        return "곧 완료"
    _s = int(round(float(seconds)))
    if _s < 60:
        return f"약 {_s}초"
    _m, _r = divmod(_s, 60)
    # 1~5분 구간: 초 단위까지
    if _m < 5:
        if _r == 0:
            return f"약 {_m}분"
        return f"약 {_m}분 {_r}초"
    # 5분+ 구간: 버킷으로 라운드
    if _m < 10:
        _bucket = max(2, round(_m / 2) * 2)
    elif _m < 30:
        _bucket = max(10, round(_m / 5) * 5)
    else:
        _bucket = max(30, round(_m / 10) * 10)
    return f"약 {_bucket}분"


def _render_batch_overlay(
    placeholder,
    *,
    idx: int,
    total: int,
    company: str,
    remaining_seconds: float | None = None,
    done: bool = False,
) -> None:
    """일괄 정산 진행 중 전체 화면 dim + 가운데 진행 카드 렌더.

    - position: fixed 로 페이지 전체를 덮어 위젯 오해/오클릭 방지.
    - placeholder.html(...) 로 호출마다 내부 텍스트 갱신.
    - remaining_seconds: 호출부가 EMA 등으로 미리 계산해 넘긴 남은 초.
      None 이면 ETA 미표시 (첫 회사 처리 중 등 추정 불가 상태).
    - 완료 시 done=True 로 호출하면 자동으로 placeholder.empty() 호출하지 않고
      메시지만 바꿔 두므로, 호출 직후 placeholder.empty() 로 오버레이를 제거할 것.
    """
    if total <= 0:
        return
    _pct = min(100, int(round(idx / total * 100)))
    if remaining_seconds is not None and not done:
        _meta = f"남은 시간 {_format_eta_seconds(remaining_seconds)}"
    else:
        _meta = "&nbsp;"
    # company 안전 escape — < > & 만 처리해도 충분 (단순 텍스트)
    import html as _html
    _safe_company = _html.escape(str(company or ""))
    _title_text = "정산 완료" if done else "정산 진행 중"
    _sub = "잠시 후 결과가 표시됩니다…" if done else f"현재: {_safe_company}"
    # 진행 중일 때만 스피너 + 진행바 시머 애니메이션. 완료 시 정적 체크.
    _spinner_html = (
        '<div class="sph-overlay-check">✓</div>' if done
        else '<div class="sph-overlay-spinner"></div>'
    )
    _bar_class = "sph-overlay-bar-fill" + ("" if done else " sph-overlay-bar-shimmer")
    placeholder.html(f"""
<style>
@keyframes sph-spin {{
  0%   {{ transform: rotate(0deg); }}
  100% {{ transform: rotate(360deg); }}
}}
@keyframes sph-shimmer {{
  0%   {{ background-position: -200px 0; }}
  100% {{ background-position: 200px 0; }}
}}
@keyframes sph-pulse {{
  0%, 100% {{ opacity: 1; }}
  50%      {{ opacity: 0.55; }}
}}
.sph-overlay-wrap {{
  position: fixed; inset: 0; z-index: 9999;
  background: rgba(15,18,22,0.72);
  display: flex; align-items: center; justify-content: center;
  backdrop-filter: blur(2px);
  -webkit-backdrop-filter: blur(2px);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}}
.sph-overlay-card {{
  background: #ffffff;
  border-radius: 16px;
  padding: 28px 36px;
  min-width: 380px; max-width: 92vw;
  box-shadow: 0 20px 60px rgba(0,0,0,0.45);
  text-align: center;
}}
.sph-overlay-title-row {{
  display: flex; align-items: center; justify-content: center;
  gap: 12px; margin-bottom: 16px;
}}
.sph-overlay-spinner {{
  width: 22px; height: 22px;
  border: 3px solid #e6ebf0;
  border-top-color: #0b6fda;
  border-radius: 50%;
  animation: sph-spin 0.9s linear infinite;
  flex-shrink: 0;
}}
.sph-overlay-check {{
  width: 22px; height: 22px;
  background: #21b96d; color: #ffffff;
  border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-weight: 700; font-size: 0.9rem;
  flex-shrink: 0;
}}
.sph-overlay-title {{
  font-size: 1.15rem; font-weight: 700; color: #1f2933;
  letter-spacing: 0.2px;
}}
.sph-overlay-count {{
  font-size: 2.4rem; font-weight: 800; color: #0b6fda;
  line-height: 1.0; margin: 6px 0 10px 0;
  font-variant-numeric: tabular-nums;
}}
.sph-overlay-bar {{
  width: 100%; height: 10px; background: #e6ebf0;
  border-radius: 999px; overflow: hidden; margin: 6px 0 14px 0;
}}
.sph-overlay-bar-fill {{
  height: 100%;
  background: linear-gradient(90deg, #0b6fda, #21b6f6);
  width: {_pct}%;
  transition: width 0.25s ease-out;
  border-radius: 999px;
}}
.sph-overlay-bar-shimmer {{
  background: linear-gradient(
    90deg,
    #0b6fda 0%,
    #21b6f6 40%,
    #5fd0fa 50%,
    #21b6f6 60%,
    #0b6fda 100%
  );
  background-size: 200px 100%;
  animation: sph-shimmer 1.4s linear infinite;
}}
.sph-overlay-sub {{
  font-size: 0.95rem; color: #475568; margin-bottom: 4px;
  word-break: break-all;
  animation: sph-pulse 1.8s ease-in-out infinite;
}}
.sph-overlay-meta {{
  font-size: 0.85rem; color: #7a8a90;
  font-variant-numeric: tabular-nums;
}}
/* 본문 위젯 클릭 차단 (시각/물리 둘 다) */
section[data-testid="stMain"], section[data-testid="stSidebar"] {{
  pointer-events: none !important;
}}
.sph-overlay-wrap, .sph-overlay-card {{
  pointer-events: auto !important;
}}
</style>
<div class="sph-overlay-wrap">
  <div class="sph-overlay-card">
    <div class="sph-overlay-title-row">
      {_spinner_html}
      <div class="sph-overlay-title">{_title_text}</div>
    </div>
    <div class="sph-overlay-count">{idx} / {total}</div>
    <div class="sph-overlay-bar"><div class="{_bar_class}"></div></div>
    <div class="sph-overlay-sub">{_sub}</div>
    <div class="sph-overlay-meta">{_meta}</div>
  </div>
</div>
""")


def _format_progress_text(
    idx: int, total: int, company: str,
    elapsed: float, finished: int,
) -> str:
    """progress bar 텍스트 — 진행률 + 평균/사 + 남은 추정 시간.
    finished == 0 이면 평균/남은시간 미표시 (아직 추정 불가).
    """
    _base = f"({idx}/{total}) {company} 정산 중..."
    if finished <= 0 or total <= finished:
        return _base
    _avg = elapsed / finished
    _remain_n = total - finished
    _eta = _format_eta_seconds(_avg * _remain_n)
    return f"{_base} · 평균 {_avg:.1f}초/사 · 남은 {_eta} ({_remain_n}개)"


def _save_batch_selection(data: dict) -> None:
    SAVED_BATCH_SELECTION_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_BATCH_SELECTION_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── 일괄 정산: 단일 회사 정산 함수 ─────────────────────────────────────────
# 단일 모드(webapp 의 정산 실행 블록) 와 **동일한 함수 체인** 으로 정산하므로
# 같은 입력(회사/CSV/단가표/회사별 저장값) 에 대해 같은 결과를 반환한다.
# UI 호출(st.warning/loading 등) 제외, 예외는 잡아 result dict 에 담는다.
def _run_batch_single_billing(
    *,
    selected_company: str,
    billing_month: str,
    tmp_input_path: Path,
    price_list_file,
    currency: str,
    exchange_rate: float,
    margin_rate: float,
    rate_date_str: str,
    billing_mode: str,
    include_project_sheet: bool,
    subtotal_round: int,
    bank_name: str,
    rate_phrase_text: str,
    rate_extra_text: str,
    min_charge_amount: float,
    min_charge_currency: str,
    sku_order: list[str],
    manual_skus: list[str],
    hidden_skus: list[str],
    billable_skus: set | None,
    dl_xlsx: bool,
    dl_pdf: bool,
    pdf_converter=None,
) -> dict:
    """한 회사 정산 → result dict.
    반환 키: ok, error, excel_bytes, pdf_bytes, pdf_error, paid_in_hidden,
            traceback (실패 시).
    """
    import time as _t_perf
    _t_start = _t_perf.time()
    _tag = f"[정산] {selected_company}"
    # 단계별 소요시간 누적 — 호출자가 UI 에 표시할 수 있도록 result dict 에 동봉.
    _timings: dict = {
        "preprocess": 0.0, "calc": 0.0, "excel": 0.0, "pdf": 0.0, "total": 0.0,
    }
    from decimal import Decimal as _Decimal
    from billing.models import BillingLineItem as _BLI
    try:
        # ── 1) 전처리 + sku_master ─────────────────────────────────
        _t0 = _t_perf.time()
        raw_rows  = preprocess_usage_file(
            str(tmp_input_path), billing_month, company_filter=selected_company,
        )
        _t_preprocess = _t_perf.time() - _t0
        print(f"{_tag} 1a) preprocess_usage_file: {_t_preprocess:.3f}초 (rows={len(raw_rows)})")

        _t0 = _t_perf.time()
        usage_rows = load_usage_rows(raw_rows)
        _t_load = _t_perf.time() - _t0
        print(f"{_tag} 1b) load_usage_rows: {_t_load:.3f}초")

        _t0 = _t_perf.time()
        sku_master = build_sku_master_from_usage(usage_rows, price_list_file)
        _t_master = _t_perf.time() - _t0
        print(f"{_tag} 1c) build_sku_master: {_t_master:.3f}초 (skus={len(sku_master)})")

        _t0 = _t_perf.time()
        _missing_skus = detect_missing_skus(usage_rows, sku_master)
        _t_missing = _t_perf.time() - _t0
        print(f"{_tag} 1d) detect_missing_skus: {_t_missing:.3f}초")
        _timings["preprocess"] = _t_preprocess + _t_load + _t_master + _t_missing

        _ex = _Decimal(str(exchange_rate))
        _mr = _Decimal(str(margin_rate))

        # ── 2) per_project 모드 free cap rollover ─────────────────
        _per_proj_invoices = None
        _proj_sku_free_cap_map: dict = {}
        if billing_mode == "per_project":
            from collections import defaultdict as _dd
            _proj_rows_map: dict = _dd(list)
            _proj_name_map: dict[str, str] = {}
            _proj_sku_usage: dict[str, dict[str, int]] = _dd(lambda: _dd(int))
            _sid_to_name: dict[str, str] = {}
            for _r in usage_rows:
                _proj_rows_map[_r.project_id].append(_r)
                _proj_name_map.setdefault(
                    _r.project_id,
                    getattr(_r, "project_name", None) or _r.project_id,
                )
                _proj_sku_usage[_r.project_id][_r.sku_id] += int(_r.usage_amount or 0)
                _nm = getattr(_r, "sku_name", None)
                if _nm and _r.sku_id not in _sid_to_name:
                    _sid_to_name[_r.sku_id] = str(_nm).strip()

            _price_caps: dict[str, int] = {}
            if price_list_file is not None:
                try:
                    _price_caps = get_free_caps_from_price_list(price_list_file)
                except Exception:
                    _price_caps = {}

            def _full_cap_for_sid(_sid: str) -> int:
                _nm = _sid_to_name.get(_sid, "")
                if _nm and _nm in _price_caps:
                    return int(_price_caps[_nm])
                _sku = sku_master.get(_sid)
                return int(getattr(_sku, "free_usage_cap", 0) or 0)

            _all_sku_ids: set[str] = set()
            for _sku_map in _proj_sku_usage.values():
                _all_sku_ids.update(_sku_map.keys())

            _proj_sku_free_cap: dict = _dd(dict)
            for _sid in _all_sku_ids:
                _full_cap = _full_cap_for_sid(_sid)
                if _full_cap <= 0:
                    continue
                _rank = sorted(
                    [(pid, _proj_sku_usage[pid].get(_sid, 0))
                     for pid in _proj_rows_map.keys()
                     if _proj_sku_usage[pid].get(_sid, 0) > 0],
                    key=lambda x: (-x[1], x[0]),
                )
                _remaining = _full_cap
                for _pid, _pu in _rank:
                    _take = min(_remaining, _pu)
                    _proj_sku_free_cap[_pid][_sid] = int(_take)
                    _remaining -= _take
            _proj_sku_free_cap_map = dict(_proj_sku_free_cap)

            _per_proj_invoices = []
            for _pid in sorted(_proj_rows_map.keys()):
                _items = calculate_billing(
                    _proj_rows_map[_pid], sku_master, _ex, _mr,
                    mode="account",
                    free_cap_override=_proj_sku_free_cap.get(_pid),
                )
                _per_proj_invoices.append({
                    "proj_name":  _proj_name_map[_pid],
                    "line_items": _items,
                })

        # ── 3) calculate_billing / by_project ────────────────────
        _t0 = _t_perf.time()
        line_items   = calculate_billing(
            usage_rows, sku_master, _ex, _mr, mode=billing_mode,
        )
        _t_calc = _t_perf.time() - _t0
        print(f"{_tag} 3a) calculate_billing: {_t_calc:.3f}초 (items={len(line_items)})")

        _t0 = _t_perf.time()
        proj_results = calculate_billing_by_project(
            usage_rows, sku_master, _ex, _mr, mode=billing_mode,
            proj_sku_free_cap=(
                _proj_sku_free_cap_map if billing_mode == "per_project" else None
            ),
        )
        _t_calc_proj = _t_perf.time() - _t0
        print(f"{_tag} 3b) calculate_billing_by_project: {_t_calc_proj:.3f}초")
        _timings["calc"] = _t_calc + _t_calc_proj

        # ── 4) manual_skus stub 주입 ─────────────────────────────
        _manual_keep_set: set[str] = set()
        if manual_skus:
            def _make_stub(_nm: str):
                return _BLI(
                    billing_month   = billing_month or "",
                    project_id      = "",
                    project_name    = "",
                    sku_id          = "",
                    sku_name        = _nm,
                    total_usage     = 0,
                    free_usage_cap  = 0,
                    free_cap_applied= 0,
                    billable_usage  = 0,
                    tier_breakdown  = [],
                    subtotal_usd    = _Decimal("0"),
                    exchange_rate   = _ex,
                    margin_rate     = _mr,
                    final_krw       = _Decimal("0"),
                )
            _existing_names = {getattr(_it, "sku_name", "") for _it in line_items}
            _missing_manual = [m for m in manual_skus if m and m not in _existing_names]
            for _nm in _missing_manual:
                line_items.append(_make_stub(_nm))
            _manual_keep_set |= set(manual_skus)
            if _per_proj_invoices:
                for _entry in _per_proj_invoices:
                    _proj_items = _entry.get("line_items") or []
                    _proj_names = {getattr(_it, "sku_name", "") for _it in _proj_items}
                    for _nm in manual_skus:
                        if _nm and _nm not in _proj_names:
                            _proj_items.append(_make_stub(_nm))
                    _entry["line_items"] = _proj_items

        # ── 5) hidden 안 비용발생 SKU 검출 ───────────────────────
        _hidden_set = set(hidden_skus or [])
        _paid_in_hidden = [
            (getattr(_it, "sku_name", ""), int(getattr(_it, "final_krw", 0) or 0))
            for _it in line_items
            if getattr(_it, "sku_name", "") in _hidden_set
            and int(getattr(_it, "final_krw", 0) or 0) > 0
        ]

        # ── 6) hidden 필터 (출력용 사본) ─────────────────────────
        if _hidden_set:
            _line_items_out = [
                _it for _it in line_items
                if getattr(_it, "sku_name", "") not in _hidden_set
            ]
            _proj_results_out = []
            for _pr in (proj_results or []):
                _skus_f = {
                    _nm: _v for _nm, _v in (_pr.get("skus") or {}).items()
                    if _nm not in _hidden_set
                }
                if not _skus_f:
                    continue
                _new_pr = dict(_pr)
                _new_pr["skus"] = _skus_f
                _new_pr["total_usd"] = sum(
                    (_v.get("subtotal_usd") or 0) for _v in _skus_f.values()
                )
                _new_pr["total_krw"] = sum(
                    (_v.get("final_krw") or 0) for _v in _skus_f.values()
                )
                _proj_results_out.append(_new_pr)
            _per_proj_invoices_out = None
            if _per_proj_invoices is not None:
                _per_proj_invoices_out = []
                for _entry in _per_proj_invoices:
                    _items_f = [
                        _it for _it in (_entry.get("line_items") or [])
                        if getattr(_it, "sku_name", "") not in _hidden_set
                    ]
                    _per_proj_invoices_out.append({
                        "proj_name":  _entry.get("proj_name"),
                        "line_items": _items_f,
                    })
        else:
            _line_items_out        = line_items
            _proj_results_out      = proj_results
            _per_proj_invoices_out = _per_proj_invoices

        _sku_order_out = [
            _n for _n in (sku_order or []) if _n not in _hidden_set
        ]

        # ── 7) Excel 생성 ────────────────────────────────────────
        _excel_bytes = None
        if dl_xlsx or dl_pdf:
            _t0 = _t_perf.time()
            _excel_bytes = generate_formatted_invoice(
                line_items           = _line_items_out,
                company_name         = selected_company or "전체",
                billing_month        = billing_month,
                exchange_rate        = _ex,
                margin_rate          = _mr,
                bank_name            = bank_name,
                proj_results         = _proj_results_out,
                price_list_file      = price_list_file,
                sku_order            = _sku_order_out or None,
                currency             = currency,
                billable_skus        = billable_skus,
                billing_mode         = billing_mode,
                per_project_invoices = _per_proj_invoices_out,
                min_charge_amount    = float(min_charge_amount),
                min_charge_currency  = min_charge_currency,
                rate_date_str        = rate_date_str,
                rate_phrase          = rate_phrase_text,
                rate_extra           = (rate_extra_text or "").strip(),
                include_project_sheet= include_project_sheet,
                subtotal_round       = subtotal_round,
                force_keep_skus      = _manual_keep_set or None,
            )
            _t_excel = _t_perf.time() - _t0
            print(f"{_tag} 7) generate_formatted_invoice: {_t_excel:.3f}초 (bytes={len(_excel_bytes) if _excel_bytes else 0})")
            _timings["excel"] = _t_excel

        # ── 7-1) 엑셀 자체 정합성 검사 ─────────────────────────────
        # *0 수식 박힘 같은 사고를 외부 발송 전에 차단하기 위한 사후 검사.
        # 결과 dict 의 validation_warnings 로 호출자가 UI 표시/차단 결정.
        _val_warns: list[str] = []
        if _excel_bytes:
            try:
                _val_warns = validate_invoice_excel(
                    _excel_bytes,
                    line_items=_line_items_out,
                    company_name=selected_company or "전체",
                )
            except Exception as _ve:
                _val_warns = [f"검사 함수 자체 오류: {type(_ve).__name__}: {_ve}"]
            if _val_warns:
                print(f"{_tag} ⚠ 정합성 경고 {len(_val_warns)}건:")
                for _w in _val_warns[:3]:
                    print(f"   - {_w}")

        # ── 8) PDF 변환 ─────────────────────────────────────────
        _pdf_bytes = None
        _pdf_error = None
        if dl_pdf and _excel_bytes:
            _t0 = _t_perf.time()
            if billing_mode == "per_project" and _per_proj_invoices_out:
                from invoice_generator import _safe_sheet_title
                _pdf_sheet = _safe_sheet_title(
                    _per_proj_invoices_out[0]["proj_name"], used=[]
                )
            else:
                _pdf_sheet = "Invoice"
            # pdf_converter 가 주어지면 (일괄 정산) 살아있는 Excel 인스턴스 재사용,
            # 아니면 단일 호출 (회사당 Excel 새로 띄움).
            if pdf_converter is not None:
                _pdf_bytes, _pdf_error = pdf_converter.convert(_excel_bytes, _pdf_sheet)
                _pdf_via = "BatchExcelPdf"
            else:
                from pdf_export import xlsx_sheet_to_pdf
                _pdf_bytes, _pdf_error = xlsx_sheet_to_pdf(_excel_bytes, _pdf_sheet)
                _pdf_via = "single"
            _t_pdf = _t_perf.time() - _t0
            print(f"{_tag} 8) PDF 변환 [{_pdf_via}]: {_t_pdf:.3f}초 (bytes={len(_pdf_bytes) if _pdf_bytes else 0})")
            _timings["pdf"] = _t_pdf

        _t_total = _t_perf.time() - _t_start
        print(f"{_tag} === 합계: {_t_total:.3f}초 ===")
        _timings["total"] = _t_total
        return {
            "ok": True, "error": None,
            "excel_bytes": _excel_bytes, "pdf_bytes": _pdf_bytes,
            "pdf_error": _pdf_error, "paid_in_hidden": _paid_in_hidden,
            "missing_skus": _missing_skus,
            "timings": _timings,
            "validation_warnings": _val_warns,
        }
    except Exception as e:
        import traceback as _tb
        _timings["total"] = _t_perf.time() - _t_start
        return {
            "ok": False, "error": f"{type(e).__name__}: {e}",
            "traceback": _tb.format_exc(),
            "excel_bytes": None, "pdf_bytes": None, "pdf_error": None,
            "paid_in_hidden": [], "missing_skus": [],
            "timings": _timings,
            "validation_warnings": [],
        }


# ══════════════════════════════════════════════════════════════════════════════
# 일괄 정산 UI — 회사 리스트 체크 + 일괄 환율/날짜 입력 + 전체 정산 → zip
# ══════════════════════════════════════════════════════════════════════════════
def render_batch_billing_ui(
    *,
    tmp_input_path: Path,
    companies: list[str],
    billing_month: str,
    price_list_file,
    currency: str,
    billable_skus,
):
    """일괄 정산 모드 진입 시 호출되는 메인 UI.
       정상 단일 모드와 동일한 입력(saved 값) 을 사용하므로 결과가 일치한다.
    """
    import datetime as _dt
    import zipfile as _zip
    import io as _io
    import re as _re

    st.markdown("#### 전체 일괄 정산")
    st.caption(
        "체크한 회사들을 한 번에 정산해 회사별 폴더로 zip 다운로드합니다. "
        "각 회사의 **저장된 설정**(과금방식·소수점·최소사용비용·환율 표기·"
        "미노출 SKU·직접등록·SKU 순서) 을 그대로 사용하므로 개별 정산과 결과가 일치합니다."
    )

    # ── 저장된 선택 상태 로드 + 신규 회사 표시용 known 계산 ───────────
    _saved_batch = _load_batch_selection()
    _known_set   = set(_saved_batch.get("known", []))
    _norm_known  = {_norm_account_key(k) for k in _known_set}
    def _is_known(c: str) -> bool:
        return _norm_account_key(c) in _norm_known

    _new_companies = [c for c in companies if not _is_known(c)]
    _updated_known = sorted(_known_set | set(companies))

    # 일괄 로드 (회사 100+개 환경에서 매 회사마다 디스크 read 하던 병목 제거).
    _orders_all      = _load_saved_orders()
    _manual_all      = _load_manual_skus_map()
    _hidden_all      = _load_hidden_skus_map()
    _mode_all        = _load_billing_modes()
    _round_all       = _load_subtotal_round_map()
    _proj_flag_all   = _load_include_project_flags()
    _rate_all        = _load_rate_labels()
    _min_charges_all = _load_min_charges()
    _notes_all       = _load_company_notes()

    # 정규화 키 dict precompute — _lookup_account 의 O(N) 키 순회를 O(1) 로 변경.
    def _norm_dict(d: dict) -> dict:
        return {_norm_account_key(k): v for k, v in d.items()}
    _orders_norm    = _norm_dict(_orders_all)
    _manual_norm    = _norm_dict(_manual_all)
    _hidden_norm    = _norm_dict(_hidden_all)
    _mode_norm      = _norm_dict(_mode_all)
    _proj_flag_norm = _norm_dict(_proj_flag_all)
    _rate_norm      = _norm_dict(_rate_all)
    _min_norm       = _norm_dict(_min_charges_all)
    _notes_norm     = _norm_dict(_notes_all)

    def _fast_lookup(exact_d, norm_d, c, norm_c):
        if c in exact_d:
            return exact_d[c]
        return norm_d.get(norm_c)

    # 회사명 → 정규화 키 캐시 (정규식 호출 최소화)
    _norm_of = {c: _norm_account_key(c) for c in companies}

    def _summary(c: str) -> str:
        _nc = _norm_of.get(c) or _norm_account_key(c)
        _mode  = _fast_lookup(_mode_all,  _mode_norm,  c, _nc) or BILLING_MODE_ACCOUNT
        _round = _round_all.get(c, 0 if currency == "KRW" else 2)
        _proj  = _fast_lookup(_proj_flag_all, _proj_flag_norm, c, _nc)
        _proj  = True if _proj is None else bool(_proj)
        _mc    = _fast_lookup(_min_charges_all, _min_norm, c, _nc) or {}
        _amt   = float(_mc.get("amount", DEFAULT_MIN_CHARGE_AMOUNT) or 0)
        _cur   = _mc.get("currency", DEFAULT_MIN_CHARGE_CURRENCY)
        _hidden_n = len(_fast_lookup(_hidden_all, _hidden_norm, c, _nc) or [])
        _manual_n = len(_fast_lookup(_manual_all, _manual_norm, c, _nc) or [])
        _mode_tag = "회사통합" if _mode == BILLING_MODE_ACCOUNT else "프로젝트별"
        _round_tag = ",0" if _round == 0 else ",2"
        _proj_tag  = "Proj✓" if _proj else "Proj✗"
        _min_tag = (
            f"최소 {_cur} {int(_amt):,}" if (_amt and _amt > 0) else "최소-"
        )
        return (
            f"{_mode_tag} · {_round_tag} · {_proj_tag} · {_min_tag} · "
            f"hidden {_hidden_n} · 직접등록 {_manual_n}"
        )

    _today = _dt.date.today()
    _default_prev_bd = _last_business_day_of_prev_month(_today)
    with st.container(border=True):
        st.markdown("#### 💱 일괄 입력 (USD 회사에만 적용)")
        c1, c2 = st.columns([1, 1])
        with c1:
            # 단일 모드와 동일 UX — 비어있으면 빨간 테두리(CSS placeholder-shown).
            batch_rate = st.number_input(
                "환율 (₩/$)",
                min_value=0.0, value=None, step=0.01, format="%.2f",
                placeholder="예: 1427.87",
                key="_batch_rate_input",
                help="USD 단가표 회사들에만 적용. 변경 시 회사선택 영역의 "
                     "환율도 일괄 갱신됩니다.",
            )
            # 정산 시작 시 환율 미입력 에러 메시지가 채워질 자리.
            _batch_rate_error_ph = st.empty()
            # 사용자가 입력하면 invalid flash 자동 해제.
            if batch_rate is not None and st.session_state.get("_batch_rate_invalid_flash"):
                st.session_state.pop("_batch_rate_invalid_flash", None)
            # flash 가 켜진 채 rerun 됐으면 메시지 + 인풋 포커스/스크롤.
            if st.session_state.get("_batch_rate_invalid_flash"):
                _batch_rate_error_ph.markdown(
                    '<div style="color:#ef4444; font-size:0.85rem; '
                    'margin-top:-10px; padding-left:4px;">'
                    '환율을 입력해 주세요</div>',
                    unsafe_allow_html=True,
                )
                # JS 실행 보장: st.html 은 환경에 따라 script 차단되므로
                # components.v1.html (iframe 렌더) 사용. iframe 안에서 부모
                # 문서 접근으로 number_input 찾고 retry 로 DOM 안정 대기.
                from streamlit.components.v1 import html as _focus_html
                _focus_html(
                    """
                    <script>
                    (function(){
                      function tryFocus(attempt){
                        try {
                          var doc = window.parent.document;
                          var container = doc.querySelector('.st-key-_batch_rate_input');
                          if (!container) {
                            if (attempt < 30) return setTimeout(
                              function(){ tryFocus(attempt+1); }, 80
                            );
                            return;
                          }
                          var inp = container.querySelector(
                            'input:not([type="hidden"])'
                          );
                          if (inp && !inp.disabled) {
                            inp.scrollIntoView({behavior:'smooth', block:'center'});
                            setTimeout(function(){
                              try {
                                inp.focus();
                                var len = (inp.value || '').length;
                                inp.setSelectionRange(len, len);
                              } catch(e) {}
                            }, 250);
                          }
                        } catch(e) {}
                      }
                      tryFocus(0);
                    })();
                    </script>
                    """,
                    height=0,
                )
        with c2:
            batch_rate_date = st.date_input(
                "환율 날짜",
                value=_default_prev_bd,
                key="_batch_rate_date",
                format="YYYY-MM-DD",
                help=(
                    "기본값: 오늘 기준 전월 마지막 은행 영업일"
                    f" ({_default_prev_bd.strftime('%Y.%m.%d')}). "
                    "한국 공휴일·주말 자동 제외. 변경 시 회사선택 영역의 "
                    "날짜도 일괄 갱신됩니다."
                ),
            )
        _batch_rate_date_str = batch_rate_date.strftime("%Y.%m.%d")

    # ── 일괄 입력 → 회사 widget state 동기화 ──────────────────────────
    # batch_rate / batch_rate_date 와 회사 session_state 가 일치하도록
    # *첫 진입 + 일괄값 변경 시* 모두 일괄값으로 덮어쓴다. 동기화가 일어나면
    # data_editor 의 버전 카운터를 증가시켜 강제 재마운트 → 편집 델타 무시.
    _prev_br = st.session_state.get("_prev_batch_rate")
    _prev_bd_state = st.session_state.get("_prev_batch_date")
    # batch_rate=None(빈칸) 이면 sync 생략 — 0/None 이 회사별로 퍼지지 않도록.
    _do_sync_rate = (
        batch_rate is not None
        and ((_prev_br is None) or (batch_rate != _prev_br))
    )
    _do_sync_date = (_prev_bd_state is None) or (batch_rate_date != _prev_bd_state)
    if _do_sync_rate or _do_sync_date:
        for _cc in companies:
            _ncc = _norm_account_key(_cc)
            if _do_sync_rate:
                st.session_state[f"_batch_rate_{_ncc}"] = float(batch_rate)
            if _do_sync_date:
                st.session_state[f"_batch_date_{_ncc}"] = batch_rate_date
        # data_editor 재마운트 — 편집 델타가 덮어쓴 셀이 일괄값으로 강제 반영.
        st.session_state["_batch_de_version"] = (
            st.session_state.get("_batch_de_version", 0) + 1
        )
    st.session_state["_prev_batch_rate"] = batch_rate
    st.session_state["_prev_batch_date"] = batch_rate_date

    with st.container(border=True):
        st.markdown("#### ⚙️ 비용발생 SKU 정책")
        _policy_options = {
            "그대로 진행 (기본)": "as_is",
            "노출로 옮기고 재정산": "move",
            "건너뛰기 (해당 회사 정산 안 함)": "skip",
        }
        _policy_labels = list(_policy_options.keys())
        _saved_policy = _saved_batch.get("policy", "as_is")
        _policy_idx = next(
            (i for i, lb in enumerate(_policy_labels)
             if _policy_options[lb] == _saved_policy), 0,
        )
        _policy_label = st.radio(
            "hidden 안에 무료 한도 초과 SKU 가 발견되었을 때:",
            options=_policy_labels, index=_policy_idx, horizontal=False,
            key="_batch_policy",
            help=(
                "• 그대로 진행: 엑셀 총액이 실제 청구액보다 적게 표시될 수 있음.\n"
                "• 노출로 옮김: 해당 SKU 를 hidden 에서 빼고 saved_orders 끝에 추가 후 재정산. saved 데이터가 영구 변경됩니다.\n"
                "• 건너뛰기: 그 회사는 정산 결과에서 제외됨."
            ),
        )
        batch_policy = _policy_options[_policy_label]

    with st.container(border=True):
        st.markdown("#### 📥 다운로드 옵션")
        _pdf_ok = _pdf_export_available()
        dl_xlsx = st.checkbox(
            "📗 엑셀 (.xlsx)",
            value=bool(_saved_batch.get("dl_xlsx", True)),
            key="_batch_dl_xlsx",
        )
        dl_pdf = st.checkbox(
            ("📄 PDF" if _pdf_ok else "📄 PDF (현재 환경에서 변환 불가)"),
            value=bool(_saved_batch.get("dl_pdf", False)) and _pdf_ok,
            key="_batch_dl_pdf",
            disabled=not _pdf_ok,
        )

    # ── 회사 선택 영역 전체를 @st.fragment 로 격리 ─────────────────────
    # 핵심 깜빡임 원인: data_editor 셀 편집 → page 전체 rerun → 사이드바·
    # 일괄 입력·정책·다운로드·정산 시작 모든 위젯 재실행. fragment 안에 두면
    # 그 안 위젯 편집은 fragment-only rerun → page rerun 안 일어남.
    # 헤더 카운트도 같은 fragment 안에서 갱신해 실시간 반영.
    # 주의: fragment 안에서 st.rerun() 절대 호출 X. fragment 밖 위젯 직접 수정 X.
    _saved_selected_norm = {
        _norm_account_key(k) for k in _saved_batch.get("selected", [])
    }
    _PHRASE_CUSTOM = "✏️ 직접 입력"
    _phrase_options = list(RATE_PHRASES) + [_PHRASE_CUSTOM]

    # ── fragment 3개 분리 — 카운트/테이블/액션이 서로 영향 안 주도록 격리.
    # 핵심: data_editor (fragment B) 가 카운트 변경 (fragment A) 으로 인해
    # 재렌더되지 않도록 분리. 시각적으로는 같은 container 안에 들어가 하나의
    # 박스로 보임.

    # === Fragment A: 헤더 (타이틀만) ========================================
    # 카운트는 fragment B 의 placeholder 가 실시간 표시.
    @st.fragment
    def _render_header():
        st.markdown("#### 📋 회사 선택")
        st.caption(
            "🆕 = 이전에 본 적 없는 회사. 체크/해제 상태는 다음 방문 시 자동 복원."
        )

    # === Fragment B: 검색 + data_editor + 편집 동기화 =====================
    @st.fragment
    def _render_table():
        # 키워드 검색.
        _search_q = st.text_input(
            "🔍 회사명 검색",
            value="",
            key="_batch_search",
            placeholder="회사명 일부를 입력하면 해당 회사만 표시됩니다",
            label_visibility="collapsed",
        )
        _q_norm = _norm_account_key(_search_q) if _search_q else ""
        if _q_norm:
            _visible_companies = [
                c for c in companies
                if _q_norm in _norm_account_key(c)
                or _search_q.lower() in c.lower()
            ]
            if not _visible_companies:
                st.caption(f"🔍 '{_search_q}' 와 일치하는 회사가 없습니다.")
                # 빈 DataFrame 으로 진행하면 _df.drop(columns=["_nc"]) 에서
                # KeyError. 검색 결과 0개는 fragment 를 그냥 빠져나간다.
                return
        else:
            _visible_companies = list(companies)

        # DataFrame 캐싱 — 매 rerun 새 객체 회피.
        # batch_rate 가 None(빈칸) 일 수 있어 float 강제 변환 불필요.
        # saved_*.json 변경(단일 정산에서 옵션 수정 등) 도 감지하도록 각 dict
        # 의 짧은 md5 해시를 cache_version 에 포함. 변경 시 자동 무효화.
        import hashlib as _hashlib
        def _dict_hash(d) -> str:
            try:
                return _hashlib.md5(
                    json.dumps(d, sort_keys=True, ensure_ascii=False, default=str)
                    .encode("utf-8")
                ).hexdigest()[:10]
            except Exception:
                return ""
        _cache_version = (
            _search_q,
            batch_rate,
            batch_rate_date.isoformat() if hasattr(batch_rate_date, "isoformat") else str(batch_rate_date),
            len(companies),
            st.session_state.get("_batch_de_version", 0),
            _dict_hash(_mode_all),
            _dict_hash(_round_all),
            _dict_hash(_proj_flag_all),
            _dict_hash(_min_charges_all),
            _dict_hash(_hidden_all),
            _dict_hash(_manual_all),
            _dict_hash(_notes_all),
            _dict_hash(_rate_all),
            _dict_hash(_orders_all),
        )
        _DF_CACHE_KEY = "_batch_df_cache"
        _DF_ROWS_CACHE_KEY = "_batch_df_rows_cache"
        _DF_VER_KEY = "_batch_df_cache_ver"
        if (st.session_state.get(_DF_VER_KEY) != _cache_version
                or _DF_CACHE_KEY not in st.session_state):
            _df_rows = []
            for c in _visible_companies:
                _is_new = not _is_known(c)
                _nc = _norm_of.get(c) or _norm_account_key(c)
                _rl_c = _fast_lookup(_rate_all, _rate_norm, c, _nc) or {}
                _saved_bank   = _rl_c.get("bank")   or DEFAULT_BANK_NAME
                _saved_phrase_raw = _rl_c.get("phrase") or DEFAULT_RATE_PHRASE
                _saved_phrase = (
                    _saved_phrase_raw if _saved_phrase_raw in RATE_PHRASES
                    else DEFAULT_RATE_PHRASE
                )
                _chk_key = f"_batch_chk_{_nc}"
                if _chk_key not in st.session_state:
                    st.session_state[_chk_key] = _nc in _saved_selected_norm
                _cur_chk  = bool(st.session_state.get(_chk_key, False))
                _cur_bank = str(st.session_state.get(f"_batch_bank_{_nc}", _saved_bank))
                _cur_date = st.session_state.get(f"_batch_date_{_nc}", batch_rate_date)
                _cur_phr  = str(st.session_state.get(f"_batch_phrase_{_nc}", _saved_phrase))
                if _cur_phr not in RATE_PHRASES:
                    _cur_phr = DEFAULT_RATE_PHRASE
                # batch_rate=None(빈칸) 일 때 회사별 widget 키도 sync 안 됐을 수
                # 있어 둘 다 None 대비. 0 으로 fallback (data_editor 표시는 0.00).
                _rate_raw = st.session_state.get(f"_batch_rate_{_nc}", batch_rate)
                _cur_rate = float(_rate_raw) if _rate_raw is not None else 0.0
                _cur_note = str(
                    st.session_state.get(
                        f"_batch_note_{_nc}",
                        _fast_lookup(_notes_all, _notes_norm, c, _nc) or "",
                    ) or ""
                )
                _df_rows.append({
                    "_nc":      _nc,
                    "_company": c,        # 표기 prefix 제거 전 원본 회사명 (영구저장 키)
                    "선택":      _cur_chk,
                    "회사명":     (f"🆕 {c}" if _is_new else c),
                    "메타":      _summary(c),
                    "은행":      _cur_bank,
                    "기준일":     _cur_date,
                    "환율종류":   _cur_phr,
                    "환율값":     _cur_rate,
                    "비고":      _cur_note,
                })
            _df = pd.DataFrame(_df_rows)
            # 비고: pandas 가 빈 문자열을 NaN 으로 추론하면 data_editor 가
            # "None" 으로 렌더할 수 있어 명시적으로 빈 문자열 + str 형변환.
            if "비고" in _df.columns:
                _df["비고"] = _df["비고"].fillna("").astype(str).replace("None", "")
            st.session_state[_DF_CACHE_KEY] = _df
            st.session_state[_DF_ROWS_CACHE_KEY] = _df_rows
            st.session_state[_DF_VER_KEY] = _cache_version
        else:
            _df = st.session_state[_DF_CACHE_KEY]
            _df_rows = st.session_state[_DF_ROWS_CACHE_KEY]

        # 카운트 placeholder + 일괄 토글 버튼 3개 — 같은 줄 배치.
        # 검색 무관 전체 적용 (사용자 의도: 1인 사용 + 사고 방지).
        _cnt_col, _bt_all, _bt_none, _bt_fav = st.columns([4, 1, 1, 1])
        with _cnt_col:
            _cnt_ph = st.empty()
        def _bulk_set_checks(check_value_fn):
            """모든 companies 의 체크 상태 일괄 변경 + 캐시 무효화 + rerun."""
            for _cc in companies:
                _ncc = _norm_of.get(_cc) or _norm_account_key(_cc)
                st.session_state[f"_batch_chk_{_ncc}"] = bool(check_value_fn(_cc, _ncc))
            st.session_state.pop(_DF_VER_KEY, None)
            st.session_state["_batch_de_version"] = (
                st.session_state.get("_batch_de_version", 0) + 1
            )
            st.rerun()
        with _bt_all:
            if st.button("✅ 전체선택", key="_batch_btn_all",
                         use_container_width=True,
                         help="검색과 무관하게 모든 회사를 체크합니다."):
                _bulk_set_checks(lambda _c, _nc: True)
        with _bt_none:
            if st.button("⬜ 전체해제", key="_batch_btn_none",
                         use_container_width=True,
                         help="검색과 무관하게 모든 회사를 해제합니다."):
                _bulk_set_checks(lambda _c, _nc: False)
        with _bt_fav:
            _fav_exact = set(_BATCH_FAVORITE_COMPANIES)
            _fav_norm  = {_norm_account_key(x) for x in _BATCH_FAVORITE_COMPANIES}
            if st.button(f"⭐ 즐겨찾기 ({len(_BATCH_FAVORITE_COMPANIES)})",
                         key="_batch_btn_fav",
                         use_container_width=True,
                         help="즐겨찾기 회사들만 체크. 나머지는 모두 해제."):
                _bulk_set_checks(
                    lambda _c, _nc: (_c in _fav_exact) or (_nc in _fav_norm)
                )

        # data_editor — 셀 편집 시 이 fragment 만 rerun.
        _de_key = f"_batch_de_v{st.session_state.get('_batch_de_version', 0)}"
        _edited_df = st.data_editor(
            _df.drop(columns=["_nc", "_company"]),
            key=_de_key,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "선택": st.column_config.CheckboxColumn(
                    "선택", width="small", default=False,
                ),
                "회사명": st.column_config.TextColumn(
                    "회사명", width="medium", disabled=True,
                ),
                "메타": st.column_config.TextColumn(
                    "메타정보", width="medium", disabled=True,
                    help="회사통합/프로젝트별 · 소수점 · Proj 시트 · 최소사용비용 · hidden · 직접등록",
                ),
                "은행": st.column_config.TextColumn("은행", width="small"),
                "기준일": st.column_config.DateColumn(
                    "기준일", width="small", format="YYYY-MM-DD",
                ),
                "환율종류": st.column_config.SelectboxColumn(
                    "환율종류", width="medium",
                    options=list(RATE_PHRASES), required=True,
                ),
                "환율값": st.column_config.NumberColumn(
                    "환율값", width="small",
                    min_value=0.0, step=0.01, format="%.2f",
                ),
                "비고": st.column_config.TextColumn(
                    "비고", width="large",
                    help="회사별 메모 (자유 입력). 표시 전용 — 엑셀 출력엔 영향 없음.",
                ),
            },
            # 전체 행 펼침 — height = 행수 × 35 + 헤더 40 + 여유 20.
            # 검색으로 행 수가 줄어들면 자연스럽게 더 작아짐. 브라우저 스크롤로 이동.
            height=max(120, len(_df_rows) * 35 + 60),
        )

        # 편집 델타 → session_state 동기화 (호환 키 유지).
        # 은행/환율종류 편집 시 saved_rate_label.json 으로 즉시 영구 저장 →
        # 다음 세션 진입 시 자동 복원. (체크 상태처럼 정산 클릭과 무관하게 보존)
        _de_changes = st.session_state.get(_de_key, {}) or {}
        _edited_rows = _de_changes.get("edited_rows", {}) or {}
        _rate_label_dirty: set[str] = set()   # 영구저장이 필요한 회사명 모음
        for _row_idx, _changes_dict in _edited_rows.items():
            try:
                _row_idx = int(_row_idx)
            except (TypeError, ValueError):
                continue
            if not (0 <= _row_idx < len(_df_rows)):
                continue
            _nc = _df_rows[_row_idx]["_nc"]
            _orig_company = _df_rows[_row_idx]["_company"]
            if "선택" in _changes_dict:
                st.session_state[f"_batch_chk_{_nc}"] = bool(_changes_dict["선택"])
            if "은행" in _changes_dict:
                _v = str(_changes_dict["은행"] or "").strip() or DEFAULT_BANK_NAME
                st.session_state[f"_batch_bank_{_nc}"] = _v
                _rate_label_dirty.add(_orig_company)
            if "기준일" in _changes_dict:
                st.session_state[f"_batch_date_{_nc}"] = _changes_dict["기준일"]
            if "환율종류" in _changes_dict:
                st.session_state[f"_batch_phrase_{_nc}"] = str(
                    _changes_dict["환율종류"] or DEFAULT_RATE_PHRASE
                )
                _rate_label_dirty.add(_orig_company)
            if "환율값" in _changes_dict:
                try:
                    st.session_state[f"_batch_rate_{_nc}"] = float(
                        _changes_dict["환율값"] or 0
                    )
                except (TypeError, ValueError):
                    st.session_state[f"_batch_rate_{_nc}"] = float(batch_rate or 0)
            if "비고" in _changes_dict:
                _note_v = str(_changes_dict["비고"] or "").strip()
                st.session_state[f"_batch_note_{_nc}"] = _note_v
                _save_company_note_for_account(_orig_company, _note_v)
                # DataFrame 캐시 무효화 → 다음 rerun 시 재빌드해서 "None" 잔재 제거.
                st.session_state.pop(_DF_VER_KEY, None)

        # 편집된 회사별 환율 표기(은행/문구) 영구 저장 — 다음 접속 시 복원.
        # 환율 값/기준일은 일괄 입력값을 우선시하므로 saved_rate_label 에는
        # 안 넣고 session_state 만 유지 (기존 동작 유지).
        if _rate_label_dirty:
            for _c in _rate_label_dirty:
                _nc_c = _norm_account_key(_c)
                _bank_c = st.session_state.get(
                    f"_batch_bank_{_nc_c}", DEFAULT_BANK_NAME,
                ) or DEFAULT_BANK_NAME
                _phr_c = st.session_state.get(
                    f"_batch_phrase_{_nc_c}", DEFAULT_RATE_PHRASE,
                ) or DEFAULT_RATE_PHRASE
                _existing = _load_rate_labels().get(_c) or {}
                _save_rate_label_for_account(
                    _c,
                    bank=_bank_c, phrase=_phr_c,
                    extra=_existing.get("extra", ""),
                    date_str=_existing.get("date"),
                    rate=None, preserve_rate_if_none=True,
                )

        # 카운트 갱신 — data_editor 반환 DataFrame 의 "선택" 컬럼 sum.
        # placeholder 만 갱신하므로 data_editor 재마운트 없음 (DF 캐싱 + key 고정).
        try:
            _checked_count = int(_edited_df["선택"].sum())
        except Exception:
            _checked_count = 0
        _cnt_ph.markdown(
            f'<div style="font-weight:600; font-size:0.95rem; color:#1a3540; '
            f'margin:6px 0 4px 0;">'
            f'전체 {len(_df_rows)}개 / 체크 {_checked_count}'
            + ('  <span style="color:#7a8a90; font-size:0.85rem;">(검색 결과)</span>'
               if len(_df_rows) < len(companies) else '')
            + '</div>',
            unsafe_allow_html=True,
        )

        # batch_rate=None(빈칸) 일 때 형식화 폭발 방지 — 안내 문구만 다르게 표시.
        _rate_disp = (
            f"₩{batch_rate:,.2f}" if batch_rate is not None else "환율 미입력"
        )
        st.caption(
            f"💡 환율 날짜 기본값: 전월 마지막 은행 영업일 "
            f"**({_default_prev_bd.strftime('%Y.%m.%d')})** "
            f"· 일괄 입력값({_rate_disp} · {_batch_rate_date_str}) 변경 시 "
            "모든 회사 환율/날짜가 자동 동기화됩니다."
        )

    # === Fragment C: 정산 버튼 (체크 0 이어도 항상 표시 — 클릭 시 검증) ====
    @st.fragment
    def _render_action():
        if st.button(
            "정산하기",
            type="primary", use_container_width=True,
            key="_batch_run_btn",
        ):
            # 클릭 시점에 최신 체크 수 검증 (data_editor 의 동기화 결과 사용).
            _cnt = sum(
                1 for _c in companies
                if st.session_state.get(
                    f"_batch_chk_{_norm_of.get(_c) or _norm_account_key(_c)}",
                    False,
                )
            )
            if _cnt == 0:
                st.warning("정산할 회사를 1개 이상 체크해 주세요.")
                return
            if not (dl_xlsx or dl_pdf):
                st.warning("다운로드 형식(엑셀/PDF) 을 1개 이상 선택해 주세요.")
                return
            # 환율 미입력 차단 — flash 켜고 rerun → 환율 영역에서 안내+포커스 처리.
            if st.session_state.get("_batch_rate_input") is None:
                st.session_state["_batch_rate_invalid_flash"] = True
                st.rerun()
                return
            st.session_state["_batch_start_trigger"] = True
            st.rerun()  # page rerun — fragment 밖 정산 실행 로직 트리거.

    # === 시각적 통합 — 한 container 안에 3 fragment 호출 ===================
    # 카운트는 _render_table fragment 안의 placeholder 가 실시간 갱신 (data_editor
    # 의 DataFrame 캐싱 + key 고정 덕에 깜빡임 없음).
    with st.container(border=True, key="_batch_main_container"):
        _render_header()
        _render_table()
        _render_action()

    # fragment 밖 — 정산 실행은 trigger 가 True 일 때만.
    _start = bool(st.session_state.pop("_batch_start_trigger", False))
    _checked = [
        c for c in companies
        if st.session_state.get(
            f"_batch_chk_{_norm_of.get(c) or _norm_account_key(c)}", False
        )
    ]

    # selected (체크된 회사 목록) 는 _start 클릭 시에만 갱신 — 사용자가
    # 의도적으로 "정산하기" 한 시점의 체크 상태가 다음 진입 시 복원됨.
    # known/dl_xlsx/dl_pdf/policy 옵션은 매 rerun 즉시 저장.
    _save_batch_selection({
        "selected": (_checked if _start else _saved_batch.get("selected", [])),
        "known":    _updated_known,
        "dl_xlsx":  dl_xlsx,
        "dl_pdf":   dl_pdf,
        "policy":   batch_policy,
    })

    if not _start:
        return
    if not _checked:
        st.info("정산할 회사를 1개 이상 체크해 주세요.")
        return
    if price_list_file is None:
        st.error("Price List(xlsx) 가 없습니다. 사이드바에서 업로드해 주세요.")
        return

    # 정산 중 화면 dim + 가운데 진행 카드. 오해/오클릭 차단.
    overlay_ph = st.empty()
    log_lines:   list[str] = []
    results:     list[dict] = []
    safe_re = _re.compile(r'[\\/*?:"<>|]')

    # 정산 시작 직전 — saved_rate_label.json 재로드만.
    # ⚠ 이전: 체크된 회사 전체에 대해 session_state 값으로 일괄 덮어쓰는
    # 루프가 있었으나, 사용자가 편집 안 한 회사도 session_state 비어 있어
    # default("하나은행"/"최종 송금환율 기준") 로 강제 덮어쓰는 사고 발생.
    # data_editor 편집 시점에 이미 _save_rate_label_for_account 가 호출
    # 되어 즉시 영구 저장되므로 이 루프는 중복이라 제거. 정산 루프 진입
    # 직전에 최신 saved 값만 재로드.
    _rate_all = _load_rate_labels()
    _rate_norm = _norm_dict(_rate_all)

    import time as _t_perf2
    _t_loop_start = _t_perf2.time()
    print(f"[정산루프] 시작 - 총 {len(_checked)}개사")

    # per-company timing 라인은 결과 expander 용으로만 수집 (UI 표시는 안 함 — 오버레이로 대체)
    _timings_lines: list[str] = []

    zip_buf = _io.BytesIO()
    _finished_count = 0  # 누적 완료 회사 수 (ETA 계산용)
    # EMA(지수가중이동평균) — 한두 회사의 큰 처리 시간이 평균을 흔들지 않도록 부드럽게.
    # alpha=0.3 → 새 데이터 30%, 기존 평균 70% 반영.
    _ema_per_company: float | None = None
    _EMA_ALPHA = 0.3
    # PDF 변환은 Excel COM subprocess 시작/종료가 회사당 11~40초로 가장 큰 병목.
    # dl_pdf 가 켜진 경우만 BatchExcelPdf 컨텍스트로 Excel 1개를 batch 내내 살려두고
    # 회사마다 stdin 으로 변환 명령만 전달 → 2번째 호출부터 회사당 2~4초.
    from pdf_export import BatchExcelPdf as _BatchExcelPdf
    print(f"[정산루프] dl_pdf={dl_pdf}, BatchExcelPdf 사용여부 결정")
    _pdf_ctx = _BatchExcelPdf() if dl_pdf else None
    try:
      # __enter__ 도 try 안에서 호출 — 서버 시작 자체에서 예외가 나도 finally
      # 에서 안전하게 __exit__(=정리) 가 불리도록.
      if _pdf_ctx is not None:
          _t_enter = _t_perf2.time()
          _pdf_ctx.__enter__()
          _t_enter_elapsed = _t_perf2.time() - _t_enter
          _server_ready = (_pdf_ctx._proc is not None and _pdf_ctx._proc.poll() is None)
          print(
              f"[정산루프] BatchExcelPdf.__enter__: {_t_enter_elapsed:.3f}초, "
              f"서버 가동={_server_ready} "
              f"({'재사용 모드 (빠름)' if _server_ready else 'fallback 모드 (단일 호출)'})"
          )
      with _zip.ZipFile(zip_buf, "w", _zip.ZIP_DEFLATED) as zf:
        total = len(_checked)
        for idx, c in enumerate(_checked, 1):
            _t_company_start = _t_perf2.time()
            _remaining = (
                _ema_per_company * (total - _finished_count)
                if _ema_per_company is not None and total > _finished_count
                else None
            )
            # ──[진단] section A: overlay render ─────────────────────
            _t_sec = _t_perf2.time()
            _render_batch_overlay(
                overlay_ph, idx=idx, total=total, company=c,
                remaining_seconds=_remaining,
            )
            _t_A_overlay = _t_perf2.time() - _t_sec

            # ──[진단] section B: 회사별 저장값 lookup ───────────────
            _t_sec = _t_perf2.time()
            _nc          = _norm_of.get(c) or _norm_account_key(c)
            _saved_for   = _fast_lookup(_orders_all,    _orders_norm,    c, _nc) or []
            _manual_for  = _fast_lookup(_manual_all,    _manual_norm,    c, _nc) or []
            _hidden_for  = _fast_lookup(_hidden_all,    _hidden_norm,    c, _nc) or []
            _mode        = _fast_lookup(_mode_all,      _mode_norm,      c, _nc) or BILLING_MODE_ACCOUNT
            _round_val   = _round_all.get(c, 0 if currency == "KRW" else 2)
            _proj_flag   = _fast_lookup(_proj_flag_all, _proj_flag_norm, c, _nc)
            _proj_flag   = True if _proj_flag is None else bool(_proj_flag)
            _rl          = _fast_lookup(_rate_all,      _rate_norm,      c, _nc) or {}
            # 은행/문구는 위에서 widget 값으로 영구 저장된 _rate_all 사용.
            _bank        = _rl.get("bank", "")  or DEFAULT_BANK_NAME
            _phrase      = _rl.get("phrase", "") or DEFAULT_RATE_PHRASE
            _extra       = _rl.get("extra", "")
            # 환율/날짜 — widget session_state 의 현재 값 사용. 사용자가 회사
            # 별로 따로 수정한 경우 그 값이, 아니면 일괄값(동기화됨) 이 사용됨.
            _rate_state  = st.session_state.get(f"_batch_rate_{_nc}", batch_rate)
            _rate_for_c  = float(_rate_state) if _rate_state is not None else 0.0
            _date_state  = st.session_state.get(f"_batch_date_{_nc}")
            if isinstance(_date_state, _dt.date):
                _rate_date_for_c = _date_state.strftime("%Y.%m.%d")
            else:
                _rate_date_for_c = _batch_rate_date_str
            _min_amt, _min_cur = _min_charge_for_account(c)
            _t_B_lookup = _t_perf2.time() - _t_sec

            # ──[진단] section C: 정산 함수 호출 (xlsx + pdf 내부 측정) ─
            _t_sec = _t_perf2.time()
            res = _run_batch_single_billing(
                selected_company=c,
                billing_month=billing_month,
                tmp_input_path=tmp_input_path,
                price_list_file=price_list_file,
                currency=currency,
                exchange_rate=_rate_for_c,
                margin_rate=1.0,
                rate_date_str=_rate_date_for_c,
                billing_mode=_mode,
                include_project_sheet=_proj_flag,
                subtotal_round=int(_round_val),
                bank_name=_bank,
                rate_phrase_text=_phrase,
                rate_extra_text=_extra,
                min_charge_amount=float(_min_amt or 0),
                min_charge_currency=_min_cur or "KRW",
                sku_order=_saved_for,
                manual_skus=_manual_for,
                hidden_skus=_hidden_for,
                billable_skus=billable_skus,
                dl_xlsx=dl_xlsx,
                dl_pdf=dl_pdf,
                pdf_converter=_pdf_ctx,
            )
            _t_C_billing = _t_perf2.time() - _t_sec

            _paid_names = [nm for nm, _ in (res.get("paid_in_hidden") or [])]
            policy_applied = None
            if _paid_names:
                if batch_policy == "skip":
                    results.append({
                        "company": c, "status": "skipped",
                        "paid_in_hidden": res.get("paid_in_hidden") or [],
                        "error": "정책=건너뛰기",
                    })
                    log_lines.append(f"⏭ **{c}** — 비용발생 SKU {len(_paid_names)}개 → 건너뜀")
                    _ema_per_company = _update_ema(
                        _ema_per_company, _t_perf2.time() - _t_company_start, _EMA_ALPHA,
                    )
                    _finished_count += 1
                    continue
                elif batch_policy == "move":
                    _new_hidden = [s for s in _hidden_for if s not in _paid_names]
                    _save_hidden_skus_for_account(c, _new_hidden)
                    _new_saved = list(_saved_for) + [
                        n for n in _paid_names if n not in _saved_for
                    ]
                    _save_order_for_account(c, _new_saved)
                    policy_applied = "moved"
                    _t_sec2 = _t_perf2.time()
                    res = _run_batch_single_billing(
                        selected_company=c,
                        billing_month=billing_month,
                        tmp_input_path=tmp_input_path,
                        price_list_file=price_list_file,
                        currency=currency,
                        exchange_rate=_rate_for_c,
                        margin_rate=1.0,
                        rate_date_str=_rate_date_for_c,
                        billing_mode=_mode,
                        include_project_sheet=_proj_flag,
                        subtotal_round=int(_round_val),
                        bank_name=_bank,
                        rate_phrase_text=_phrase,
                        rate_extra_text=_extra,
                        min_charge_amount=float(_min_amt or 0),
                        min_charge_currency=_min_cur or "KRW",
                        sku_order=_new_saved,
                        manual_skus=_manual_for,
                        hidden_skus=_new_hidden,
                        billable_skus=billable_skus,
                        dl_xlsx=dl_xlsx,
                        dl_pdf=dl_pdf,
                        pdf_converter=_pdf_ctx,
                    )
                    # policy=move 인 경우 정산이 1회 더 — 그 시간도 C 에 합산.
                    _t_C_billing += _t_perf2.time() - _t_sec2

            if not res["ok"]:
                results.append({
                    "company": c, "status": "error",
                    "error": res.get("error"),
                    "paid_in_hidden": [],
                })
                log_lines.append(f"❌ **{c}** — {res.get('error')}")
                _ema_per_company = _update_ema(
                    _ema_per_company, _t_perf2.time() - _t_company_start, _EMA_ALPHA,
                )
                _finished_count += 1
                continue

            # ──[진단] section D: zip writestr (메모리 → 압축) ───────
            _t_sec = _t_perf2.time()
            _safe = safe_re.sub("_", c).strip() or "전체"
            _stem = f"sGMP_Invoice_{_safe}"
            if dl_xlsx and res.get("excel_bytes"):
                zf.writestr(f"{_safe}/{_stem}.xlsx", res["excel_bytes"])
            if dl_pdf and res.get("pdf_bytes"):
                zf.writestr(f"{_safe}/{_stem}.pdf", res["pdf_bytes"])
            _t_D_zip = _t_perf2.time() - _t_sec

            # ──[진단] section E: 결과 bookkeeping ────────────────────
            _t_sec = _t_perf2.time()
            _val_w = res.get("validation_warnings") or []
            results.append({
                "company": c,
                "status": "ok" if not res.get("paid_in_hidden") else "ok_with_paid",
                "paid_in_hidden": res.get("paid_in_hidden") or [],
                "pdf_error": res.get("pdf_error"),
                "policy_applied": policy_applied,
                "validation_warnings": _val_w,
            })
            _hint = ""
            if res.get("paid_in_hidden") and not policy_applied:
                _hint = f" · ⚠ 비용발생 {len(res['paid_in_hidden'])}개"
            elif policy_applied == "moved":
                _hint = f" · 🔁 노출이동·재정산 ({len(_paid_names)}개)"
            if _val_w:
                _hint += f" · 🚨 정합성 경고 {len(_val_w)}건"
            log_lines.append(f"✅ **{c}**{_hint}")
            _t_E_bookkeep = _t_perf2.time() - _t_sec
            _t_company_elapsed = _t_perf2.time() - _t_company_start
            _t_other = _t_company_elapsed - (
                _t_A_overlay + _t_B_lookup + _t_C_billing + _t_D_zip + _t_E_bookkeep
            )
            print(
                f"[정산루프] ({idx}/{total}) {c} 합계 {_t_company_elapsed:.2f}s = "
                f"A.overlay {_t_A_overlay:.2f} + B.lookup {_t_B_lookup:.3f} + "
                f"C.billing {_t_C_billing:.2f} + D.zip {_t_D_zip:.3f} + "
                f"E.bookkeep {_t_E_bookkeep:.3f} + 기타 {_t_other:.3f}"
            )
            _ema_per_company = _update_ema(_ema_per_company, _t_company_elapsed, _EMA_ALPHA)
            # 결과 expander 표시용 timing 라인만 수집 (UI 실시간 표시는 오버레이가 담당)
            _timings_lines.append(
                _format_billing_timings_line(c, res.get("timings") or {})
            )
            if _val_w:
                _timings_lines.append(
                    f"  🚨 **{c}** 정합성 경고: " + " / ".join(_val_w[:2])
                    + (f" (외 {len(_val_w)-2}건)" if len(_val_w) > 2 else "")
                )
            _finished_count += 1

        _t_loop_total = _t_perf2.time() - _t_loop_start
        print(f"[정산루프] 전체 완료: {_t_loop_total:.3f}초 ({len(_checked)}개사)")
        _timings_lines.append(
            f"**전체 완료: {_t_loop_total:.2f}초 ({len(_checked)}개사)**"
        )
    finally:
        # BatchExcelPdf 정리 — Excel 인스턴스 종료. 정산 도중 예외나도 반드시 호출.
        if _pdf_ctx is not None:
            try:
                _pdf_ctx.__exit__(None, None, None)
            except Exception:
                pass

    # 루프 종료 — 오버레이 제거 (결과 영역이 자연스럽게 노출됨)
    overlay_ph.empty()

    n_ok     = sum(1 for r in results if r["status"] in ("ok", "ok_with_paid"))
    n_paid   = sum(1 for r in results if r["status"] == "ok_with_paid")
    n_skip   = sum(1 for r in results if r["status"] == "skipped")
    n_err    = sum(1 for r in results if r["status"] == "error")
    n_total  = len(results)

    st.markdown("---")
    st.markdown(f"### 결과 요약 ({n_ok}/{n_total} 성공)")
    if n_paid > 0:
        st.warning(
            f"⚠️ 비용발생 SKU 가 포함된 회사 **{n_paid}개** — 엑셀 총액이 실제 "
            "청구액보다 적게 표시되었을 수 있습니다."
        )
    if n_err > 0:
        st.error(f"❌ 정산 실패 {n_err}개사 — 아래 로그 확인")
    if n_skip > 0:
        st.info(f"⏭ 건너뛴 회사 {n_skip}개사 (정책=건너뛰기)")

    # 정합성 경고 회사 — 외부 발송 전 사용자 확인 필요. 결과 요약 직후 상단 노출.
    _val_alerts = [r for r in results if r.get("validation_warnings")]
    if _val_alerts:
        st.error(
            f"🚨 **엑셀 정합성 경고가 발생한 회사 {len(_val_alerts)}개** — "
            "외부 발송 전 반드시 확인하세요."
        )
        with st.expander("🚨 정합성 경고 상세", expanded=True):
            for r in _val_alerts:
                st.markdown(f"**{r['company']}**")
                for _w in r["validation_warnings"][:10]:
                    st.markdown(f"- {_w}")
                _rem = len(r["validation_warnings"]) - 10
                if _rem > 0:
                    st.markdown(f"- … 외 {_rem}건")

    with st.expander("📋 회사별 정산 로그", expanded=False):
        for ln in log_lines:
            st.markdown(ln)
        _has_paid = [r for r in results if r.get("paid_in_hidden")]
        if _has_paid:
            st.markdown("---")
            st.markdown("**[비용발생] 상세:**")
            for r in _has_paid:
                _items = ", ".join(
                    f"{nm}(₩{kw:,})" for nm, kw in r["paid_in_hidden"]
                )
                st.markdown(f"- {r['company']}: {_items}")

    # 단계별 소요시간은 디버그용 expander 에 별도 노출 (기본 닫힘)
    if _timings_lines:
        with st.expander("⏱ 단계별 소요시간 (디버그)", expanded=False):
            st.markdown("  \n".join(_timings_lines))

    if n_ok > 0:
        _ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        _zip_name = f"전체정산_{billing_month or 'all'}_{_ts}.zip"
        st.download_button(
            f"📦 zip 다운로드 ({_zip_name})",
            data=zip_buf.getvalue(),
            file_name=_zip_name,
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )
    else:
        st.info("다운로드할 결과가 없습니다.")


# ── 계정별 최소사용비용 저장/로드 ───────────────────────────────────────────
# Google Maps Platform 은 기본 월 ₩500,000 최소사용비용 규정이 있지만,
# 회사별 계약으로 금액·통화가 달라질 수 있다(예: USD 기준, 0 원 = 적용 안 함).
# 저장 구조: { account: {"amount": 500000, "currency": "KRW"} }
def _load_min_charges() -> dict[str, dict]:
    if not SAVED_MIN_CHARGE_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_MIN_CHARGE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    out: dict[str, dict] = {}
    for acc, v in (data or {}).items():
        if not isinstance(acc, str) or not isinstance(v, dict):
            continue
        try:
            amt = float(v.get("amount", 0) or 0)
        except (TypeError, ValueError):
            amt = 0.0
        cur = v.get("currency") if v.get("currency") in ("KRW", "USD") else "KRW"
        out[acc] = {"amount": amt, "currency": cur}
    return out


def _save_min_charge_for_account(account: str, amount: float, currency: str) -> None:
    if currency not in ("KRW", "USD"):
        currency = "KRW"
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        amount = 0.0
    data = _load_min_charges()
    data[account] = {"amount": amount, "currency": currency}
    SAVED_MIN_CHARGE_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_MIN_CHARGE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _min_charge_for_account(account: str) -> tuple[float, str]:
    """계정별 최소사용비용 (금액, 통화) 반환. 미저장 시 기본값."""
    d = _load_min_charges().get(account) or {}
    amt = d.get("amount", DEFAULT_MIN_CHARGE_AMOUNT)
    cur = d.get("currency", DEFAULT_MIN_CHARGE_CURRENCY)
    return float(amt), cur


# ── 계정별 환율 표기 저장/로드 ──────────────────────────────────────────────
# Invoice 의 "환율(하나은행 2026.02.27 최종 송금환율 기준)" 조합용.
# 구조: { account: {"bank": str, "phrase": str, "extra": str, "date": "YYYY-MM-DD"|None} }
# date 가 None 이면 billing_month 의 마지막 날로 자동 세팅.
def _load_rate_labels() -> dict[str, dict]:
    if not SAVED_RATE_LABEL_FILE.exists():
        return {}
    try:
        data = json.loads(SAVED_RATE_LABEL_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    out: dict[str, dict] = {}
    for acc, v in (data or {}).items():
        if not isinstance(acc, str) or not isinstance(v, dict):
            continue
        # rate: 회사별 환율(USD 모드 일괄 정산 개별 수정값). None=일괄값 사용.
        try:
            _rv = v.get("rate")
            _rate = float(_rv) if _rv is not None else None
        except (TypeError, ValueError):
            _rate = None
        out[acc] = {
            "bank":   str(v.get("bank")   or DEFAULT_BANK_NAME),
            "phrase": str(v.get("phrase") or DEFAULT_RATE_PHRASE),
            "extra":  str(v.get("extra")  or ""),
            "date":   v.get("date") if isinstance(v.get("date"), str) else None,
            "rate":   _rate,
        }
    return out


def _save_rate_label_for_account(account: str, bank: str, phrase: str,
                                  extra: str, date_str: str | None,
                                  rate: float | None = None,
                                  preserve_rate_if_none: bool = True) -> None:
    """계정별 환율 표기 저장.

    rate 파라미터:
      - 명시값 (float): 그대로 저장
      - None + preserve_rate_if_none=True: 기존 저장값 유지 (단일 정산 화면 호환)
      - None + preserve_rate_if_none=False: 명시적으로 None 저장 (초기화)
    """
    data = _load_rate_labels()
    if rate is None and preserve_rate_if_none:
        _existing = data.get(account) or {}
        rate = _existing.get("rate")
    data[account] = {
        "bank":   (bank or DEFAULT_BANK_NAME).strip(),
        "phrase": (phrase or DEFAULT_RATE_PHRASE).strip(),
        "extra":  (extra or "").strip(),
        "date":   (date_str or None),
        "rate":   float(rate) if rate is not None else None,
    }
    SAVED_RATE_LABEL_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_RATE_LABEL_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _rate_label_for_account(account: str) -> dict:
    """계정별 환율 표기 설정 반환. 미저장 시 기본값."""
    d = _load_rate_labels().get(account) or {}
    return {
        "bank":   d.get("bank")   or DEFAULT_BANK_NAME,
        "phrase": d.get("phrase") or DEFAULT_RATE_PHRASE,
        "extra":  d.get("extra")  or "",
        "date":   d.get("date"),
        "rate":   d.get("rate"),
    }


def _last_business_day_of_prev_month(today: date | None = None) -> date:
    """전월 마지막 은행 영업일(주말·한국 공휴일 제외) 반환.

    한국 공휴일은 `holidays` 라이브러리(설/추석/부처님오신날 등 음력 명절,
    임시공휴일·대체공휴일 자동 반영) 로 계산. import 실패 시 주말만 제외하는
    폴백 동작. 사용자가 회사별 UI 에서 직접 수정도 가능.
    """
    from datetime import timedelta as _td
    today = today or date.today()
    first_of_this = today.replace(day=1)
    d = first_of_this - _td(days=1)  # 전월 말일

    # 공휴일 셋 준비 — 라이브러리 없으면 빈 셋(주말만 제외).
    try:
        import holidays as _h
        # 전월/전전월을 모두 커버하도록 두 해를 포함 (12월 → 1월 공휴일 등).
        _kr = _h.KR(years=[d.year, d.year - 1, d.year + 1])
    except Exception:
        _kr = set()

    # 주말(5=토, 6=일) 이거나 한국 공휴일이면 하루씩 앞으로.
    while d.weekday() >= 5 or d in _kr:
        d -= _td(days=1)
    return d


def _match_bank_prefix(text: str) -> str | None:
    """MAJOR_BANKS 중 입력 문자열로 시작하는 첫 후보 반환. 한 글자만 쳐도
    예: '하' → '하나은행', '국' → '국민은행' 식으로 자동완성 제안용."""
    if not text:
        return None
    t = text.strip()
    for b in MAJOR_BANKS:
        if b.startswith(t):
            return b
    return None


def _unique_skus_for_account(tmp_path: str, billing_month: str,
                              account: str | None) -> list[str]:
    """선택된 계정의 CSV 에서 **실제 사용량이 있는** SKU 이름 리스트.

    하드코딩 화이트리스트 없음 — 회사별로 쓰는 Google Maps Platform
    제품이 모두 다르므로, CSV 의 실측 usage > 0 인 SKU 만 그대로 반환한다.
    세금(`세금`/`tax`/`vat`) 항목은 인보이스 합계 외 별도 처리이므로 제외.
    초기 순서는 sku_name 가나다순 — 사용자가 드래그앤드롭으로 자유 조정.

    주의: `_cached_preprocess` 경유하지 않고 매번 `preprocess_usage_file`
    을 직접 호출한다. 과거 Streamlit `@st.cache_data` 캐시가 코드 변경
    이후에도 낡은 결과를 반환해 UI 에 일부 SKU(예: Elevation) 가 누락되는
    사고가 있었고, 이 목록은 사용자가 "실제로 뭐가 잡혔나"를 확인하는
    진실 소스이기 때문에 성능보다 정확성을 우선한다.
    """
    rows = preprocess_usage_file(
        tmp_path, billing_month, company_filter=account
    )
    usage_by_name: dict[str, int] = {}
    for r in rows:
        nm = (r.get("sku_name") or "").strip()
        if not nm or _is_tax_sku(nm):
            continue
        usage_by_name[nm] = usage_by_name.get(nm, 0) + int(r.get("usage_amount") or 0)
    return sorted(nm for nm, u in usage_by_name.items() if u > 0)

# ── 페이지 설정 ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SPH GMP 정산 시스템",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 전역 CSS ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css');

html, body, [class*="css"] {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont,
                 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif !important;
}

/* 배경 */
.stApp { background-color: #f0f4f6; }
.main .block-container { padding-top: 1.5rem; padding-bottom: 3rem; }

/* ── 사이드바 ── */
[data-testid="stSidebar"] {
    background: linear-gradient(175deg, #00788a 0%, #00505e 100%) !important;
    border-right: none !important;
}
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label { color: rgba(255,255,255,0.88) !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #ffffff !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.18) !important; }

/* ── 탭 ── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    gap: 6px;
    border-bottom: 2px solid #d4e4e8;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border-radius: 10px 10px 0 0 !important;
    padding: 10px 24px !important;
    font-size: 0.94rem !important;
    font-weight: 600 !important;
    color: #7a9ea8 !important;
    border: none !important;
    transition: background 0.15s, color 0.15s !important;
}
.stTabs [aria-selected="true"] {
    background: white !important;
    color: #00788a !important;
    border-bottom: 2px solid white !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: white;
    border-radius: 0 16px 16px 16px;
    padding: 28px 32px;
    box-shadow: 0 4px 24px rgba(0,120,138,0.07);
}

/* ── 텍스트 인풋 ── */
/* 활성 상태(달러 모드) 에서는 흰 배경. 원화 모드에서는 disabled=true
   로 바뀌면서 Streamlit 기본 회색 배경이 적용되므로 별도 처리 불필요. */
div[data-testid="stTextInput"] input:not(:disabled) {
    background-color: #ffffff !important;
}

/* 환율 인풋: 달러 모드에서 비어있으면 빨간 테두리 (입력하면 자동 해제). */
.st-key-_rate_raw_input div[data-baseweb="input"]:has(input:not(:disabled):placeholder-shown),
.st-key-_rate_raw_input div[data-baseweb="base-input"]:has(input:not(:disabled):placeholder-shown),
.st-key-_batch_rate_input div[data-baseweb="input"]:has(input:not(:disabled):placeholder-shown),
.st-key-_batch_rate_input div[data-baseweb="base-input"]:has(input:not(:disabled):placeholder-shown) {
    border-color: #ef4444 !important;
    box-shadow: 0 0 0 1px #ef4444 !important;
}

/* ── 버튼 ── */
button[kind="primary"] {
    background: linear-gradient(135deg, #00788a 0%, #00596a 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    font-size: 0.98rem !important;
    padding: 0.65rem 2rem !important;
    box-shadow: 0 4px 14px rgba(0,120,138,0.28) !important;
    letter-spacing: -0.1px !important;
    transition: box-shadow 0.15s, transform 0.1s !important;
}
button[kind="primary"]:hover {
    box-shadow: 0 6px 20px rgba(0,120,138,0.42) !important;
    transform: translateY(-1px) !important;
}
button[kind="secondary"] {
    border-radius: 10px !important;
    border: 1.5px solid #bdd8de !important;
    font-weight: 600 !important;
    color: #00788a !important;
    background: white !important;
}

/* ── 다운로드 버튼 (녹색 액센트) ── */
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #a5d15a 0%, #7dbb26 100%) !important;
    border: none !important;
    color: #1b3d06 !important;
    font-weight: 700 !important;
    border-radius: 12px !important;
    padding: 0.65rem 2rem !important;
    box-shadow: 0 4px 14px rgba(165,209,90,0.32) !important;
    font-size: 0.98rem !important;
    transition: box-shadow 0.15s, transform 0.1s !important;
}
[data-testid="stDownloadButton"] button:hover {
    box-shadow: 0 6px 20px rgba(165,209,90,0.48) !important;
    transform: translateY(-1px) !important;
}

/* ── 메트릭 카드 ── */
[data-testid="stMetric"] {
    background: white !important;
    border-radius: 16px !important;
    padding: 20px 24px !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.05) !important;
    border: 1px solid #e5eff2 !important;
}
[data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 800 !important;
    color: #00788a !important;
    letter-spacing: -0.5px !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.78rem !important;
    color: #7a9ea8 !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.3px !important;
}

/* ── 정산 대상 선택 — 셀렉트박스 영역 너비/위치 ── */
/* 평소에는 헤딩("정산 대상 선택") + 라벨("결제 계정...") 그대로 노출.
   floating 상태일 때만 헤딩·라벨을 숨겨 셀렉트박스만 컴팩트하게 보여준다.
   너비는 JS 가 좌측 SKU 컬럼 rect 를 측정해 inline style 로 적용. */
.st-key-sticky_account_select.is-floating h4,
.st-key-sticky_account_select.is-floating label,
.st-key-sticky_account_select.is-floating [data-testid="stMarkdownContainer"]:first-child,
.st-key-sticky_account_select.is-floating [data-testid="stWidgetLabel"] {
    display: none !important;
}
.st-key-sticky_account_select.is-floating [data-testid="stVerticalBlock"] {
    gap: 0 !important;
}
/* CSS fallback — JS 가 inline 으로 덮어쓰기 전까지 절반 너비 좌측 정렬 */
.st-key-sticky_account_select {
    max-width: calc(50% - 0.5rem) !important;
    margin-right: auto !important;
}


/* floating: 좌측 컬럼 위치/너비 유지 + 최상단 부착 + 컴팩트 패딩 */
.st-key-sticky_account_select.is-floating {
    position: fixed !important;
    top: 0 !important;
    transform: none !important;
    z-index: 999 !important;
    background: #ffffff !important;
    padding: 4px 8px !important;
    border-radius: 0 0 10px 10px !important;
    box-shadow: 0 4px 10px -4px rgba(0, 60, 70, 0.18) !important;
    animation: account-float-in 160ms ease-out;
}
@keyframes account-float-in {
    from { transform: translateY(-8px); opacity: 0; }
    to   { transform: translateY(0);    opacity: 1; }
}

/* ── Selectbox (결제 계정) 흰색 배경 + teal 보더 ── */
div[data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border-color: #bdd8de !important;
    border-radius: 10px !important;
    min-height: 44px !important;
}
div[data-baseweb="select"] > div:hover {
    border-color: #00788a !important;
}
div[data-baseweb="select"] input,
div[data-baseweb="select"] span {
    color: #1a3540 !important;
    background-color: transparent !important;
}

/* ── 데이터프레임 ── */
[data-testid="stDataFrame"],
[data-testid="stDataEditor"] {
    border-radius: 12px !important;
    overflow: hidden !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important;
}

/* ── 프로그레스 바 ── */
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #00788a, #a5d15a) !important;
    border-radius: 8px !important;
}

/* ── 구분선 ── */
hr { border-color: #e0eaed !important; margin: 1.5rem 0 !important; }

/* ── 섹션 헤더 ── */
h4 { color: #1a3540 !important; letter-spacing: -0.3px !important; font-size: 1.05rem !important; }

/* ── 파일 업로더 라벨 검정색 ── */
[data-testid="stFileUploader"] > label,
[data-testid="stFileUploader"] > label p {
    color: #111111 !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
}

/* ── 드래그앤드롭 영역 스타일 ── */
[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed #00788a !important;
    border-radius: 14px !important;
    background: linear-gradient(135deg, #f8fcfd 0%, #edf6f8 100%) !important;
    transition: border-color 0.2s, background 0.2s !important;
    padding: 8px 12px !important;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #005a6a !important;
    background: linear-gradient(135deg, #e8f5f8 0%, #d8eef2 100%) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: #1a3540 !important;
    font-weight: 600 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stFileUploaderDropzoneInstructions"] span small {
    color: #5a8290 !important;
}

/* ── number_input / date_input 배경을 흰색으로 (기본 회색 제거) ── */
.stNumberInput input,
.stDateInput input,
[data-testid="stDateInput"] input,
[data-testid="stNumberInput"] input {
    background-color: #ffffff !important;
}
.stNumberInput [data-baseweb="input"],
.stDateInput [data-baseweb="input"],
[data-testid="stDateInput"] [data-baseweb="input"],
[data-testid="stNumberInput"] [data-baseweb="input"] {
    background-color: #ffffff !important;
}

/* ── date_input 우측에 달력 아이콘 (유니코드 📅) 강제 표시 ── */
[data-testid="stDateInput"] [data-baseweb="input"] {
    position: relative;
}
[data-testid="stDateInput"] [data-baseweb="input"]::after {
    content: "📅";
    position: absolute;
    right: 10px;
    top: 50%;
    transform: translateY(-50%);
    pointer-events: none;
    font-size: 1rem;
    opacity: 0.75;
}
[data-testid="stDateInput"] [data-baseweb="input"] input {
    padding-right: 32px !important;
}

/* ── Streamlit 상단 영역 숨김 (Deploy / 햄버거 / 러닝 인디케이터) ── */
/* 헤더 전체와 데코레이션 바를 제거하고 본문을 상단까지 끌어올린다.
   상단 러닝 인디케이터가 사라지므로, 정산 중 진행 표시는
   화면 전체 딤(.sph-loading-overlay)으로 대체한다 (아래 정의). */
[data-testid="stHeader"]     { display: none !important; }
[data-testid="stToolbar"]    { display: none !important; }
[data-testid="stToolbarActions"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
[data-testid="stStatusWidget"] { display: none !important; }
#MainMenu                    { visibility: hidden !important; display: none !important; }
header[role="banner"]        { display: none !important; }
footer                       { visibility: hidden !important; }
.main .block-container       { padding-top: 0.6rem !important; }
[data-testid="stAppViewContainer"] > .main { padding-top: 0 !important; }
/* Streamlit 1.4x+ 본문 컨테이너 (stHeader 숨겼으므로 상단 거의 0) */
[data-testid="stMain"]            { padding-top: 0 !important; }
[data-testid="stMainBlockContainer"],
.stMainBlockContainer {
    padding-top: 0.6rem !important;
}
/* 사이드바 block-container 기본 상단 패딩 축소 (Streamlit 기본 ~6rem) */
[data-testid="stSidebar"] > div:first-child,
[data-testid="stSidebarUserContent"],
[data-testid="stSidebarContent"],
section[data-testid="stSidebar"] .block-container,
section[data-testid="stSidebar"] > div,
[data-testid="stSidebar"] [data-testid="stVerticalBlock"]:first-child {
    padding-top: 0 !important;
    margin-top: 0 !important;
}
/* 사이드바 첫 요소(로고 markdown) 자체의 상단 여백도 제거 */
[data-testid="stSidebar"] [data-testid="stElementContainer"]:first-child,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"]:first-child {
    margin-top: 0 !important;
    padding-top: 0 !important;
}
/* 사이드바 접기/펴기 버튼 숨기기 */
[data-testid="stSidebarCollapseButton"],
[data-testid="stSidebarCollapsedControl"],
[data-testid="collapsedControl"],
button[kind="headerNoPadding"] {
    display: none !important;
    visibility: hidden !important;
}
/* 사이드바 강제 표시 (접기 버튼 사용 후에도 항상 펼친 상태 유지) */
[data-testid="stSidebar"],
section[data-testid="stSidebar"] {
    display: flex !important;
    visibility: visible !important;
    transform: none !important;
    margin-left: 0 !important;
    min-width: 244px !important;
    width: 244px !important;
}
[data-testid="stSidebar"][aria-expanded="false"] {
    margin-left: 0 !important;
    transform: none !important;
}

/* ── 정산 진행 로딩 오버레이 (화면 딤) ── */
.sph-loading-overlay {
    position: fixed;
    inset: 0;
    background: rgba(15, 30, 35, 0.55);
    backdrop-filter: blur(3px);
    -webkit-backdrop-filter: blur(3px);
    z-index: 999999;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-direction: column;
    animation: sphFadeIn 0.18s ease-out;
}
@keyframes sphFadeIn { from { opacity: 0; } to { opacity: 1; } }
.sph-loading-spinner {
    width: 64px;
    height: 64px;
    border: 5px solid rgba(255,255,255,0.22);
    border-top-color: #00bcd4;
    border-radius: 50%;
    animation: sphSpin 0.9s linear infinite;
}
@keyframes sphSpin { to { transform: rotate(360deg); } }
.sph-loading-text {
    color: #fff;
    margin-top: 22px;
    font-size: 1.02rem;
    font-weight: 700;
    letter-spacing: -0.2px;
    text-shadow: 0 1px 4px rgba(0,0,0,0.3);
}
.sph-loading-bar {
    width: 280px;
    height: 6px;
    background: rgba(255,255,255,0.18);
    border-radius: 3px;
    margin-top: 16px;
    overflow: hidden;
}
.sph-loading-bar-fill {
    height: 100%;
    background: linear-gradient(90deg, #00bcd4, #00788a);
    transition: width 0.3s ease;
}
.sph-loading-pct {
    color: rgba(255,255,255,0.75);
    margin-top: 10px;
    font-size: 0.82rem;
    font-weight: 600;
    letter-spacing: 0.3px;
}

/* ── 일괄 정산 — 체크박스 (SPH teal #00788a) ──────────────────────────────
   접근 방식: Streamlit/BaseWeb 의 DOM 구조 추측 대신 transform:scale 로
   원본 박스만 시각적으로 키운다. 라벨 텍스트(label 의 두번째 자식)는
   건드리지 않아 "체크박스 두 개" 처럼 보이던 사고가 없다.
   대상: 회사 체크박스(`_batch_chk_*`), 다운로드(`_batch_dl_xlsx/_pdf`). */

/* 박스만 1.55배 확대 — label > 첫번째 자식이 시각 박스 wrapper.
   transform 은 레이아웃을 안 바꾸므로 margin-right 로 라벨과의 간격 확보. */
[class*="st-key-_batch_chk_"] [data-testid="stCheckbox"] label > span:first-child,
[class*="st-key-_batch_chk_"] [data-testid="stCheckbox"] label > div:first-child,
.st-key-_batch_dl_xlsx [data-testid="stCheckbox"] label > span:first-child,
.st-key-_batch_dl_xlsx [data-testid="stCheckbox"] label > div:first-child,
.st-key-_batch_dl_pdf [data-testid="stCheckbox"] label > span:first-child,
.st-key-_batch_dl_pdf [data-testid="stCheckbox"] label > div:first-child {
    transform: scale(1.45);
    transform-origin: center left;
    margin-right: 14px !important;
}

/* 체크 상태 — SPH teal 배경. :has(input:checked) 만 사용해 정확히 박스만 색칠.
   이전엔 `:first-of-type` 같은 광범위 셀렉터가 라벨 자식까지 색을 바꿔
   "체크박스가 두 개" 처럼 보이게 했음. */
[class*="st-key-_batch_chk_"] [data-testid="stCheckbox"] label:has(input:checked) > span:first-child,
[class*="st-key-_batch_chk_"] [data-testid="stCheckbox"] label:has(input:checked) > div:first-child,
.st-key-_batch_dl_xlsx [data-testid="stCheckbox"] label:has(input:checked) > span:first-child,
.st-key-_batch_dl_xlsx [data-testid="stCheckbox"] label:has(input:checked) > div:first-child,
.st-key-_batch_dl_pdf [data-testid="stCheckbox"] label:has(input:checked) > span:first-child,
.st-key-_batch_dl_pdf [data-testid="stCheckbox"] label:has(input:checked) > div:first-child {
    background-color: #00788a !important;
    border-color: #00788a !important;
}

/* 다운로드 옵션 + 회사 체크박스 라벨 — 동일 간격/크기로 통일. */
.st-key-_batch_dl_xlsx label,
.st-key-_batch_dl_pdf label,
[class*="st-key-_batch_chk_"] label {
    font-size: 0.92rem !important;
    line-height: 1.5 !important;
}
.st-key-_batch_dl_xlsx label,
.st-key-_batch_dl_pdf label {
    white-space: nowrap !important;
}
.st-key-_batch_dl_xlsx label p,
.st-key-_batch_dl_pdf label p,
[class*="st-key-_batch_chk_"] label p {
    font-size: 0.92rem !important;
    font-weight: 500 !important;
    color: #1a3540 !important;
    margin: 0 !important;
}

/* 회사명 아래 메타 라인 — 체크박스 박스 너비만큼 들여쓰기. */
._batch_company_meta {
    color: #7a8a90;
    font-size: 0.76rem;
    margin: -6px 0 6px 38px;  /* 위로 살짝 당기고, 체크박스 폭만큼 들여쓰기 */
    line-height: 1.4;
}

/* 회사 선택 박스 — fragment 3개를 한 container 안에 두고 시각적으로 하나의
   박스처럼 보이도록 fragment 사이 gap 축소. */
.st-key-_batch_main_container [data-testid="stVerticalBlock"] {
    gap: 0.5rem !important;
}
/* fragment 자체 영역의 추가 margin/padding 제거 */
.st-key-_batch_main_container [data-testid="stElementContainer"] {
    margin-top: 0 !important;
    margin-bottom: 0 !important;
}

/* 환율 요약 텍스트 (접힘 상태에서 현재 설정값을 한눈에). 가벼운 markdown
   div — selectbox/input 같은 무거운 위젯 대신 텍스트만 표시. */
._batch_rate_summary {
    color: #1a3540;
    font-size: 0.82rem;
    line-height: 1.5;
    padding: 6px 10px;
    background: #f6fafc;
    border-left: 3px solid #00788a;
    border-radius: 6px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* ✏️ 수정 / ✓ 접기 버튼 — 컴팩트하게 30px 높이로 통일 */
[class*="st-key-_batch_edit_btn_"] button {
    min-height: 30px !important;
    height: 30px !important;
    padding: 0 8px !important;
    font-size: 0.82rem !important;
    border-radius: 6px !important;
    border: 1px solid #d4dce0 !important;
    background: #ffffff !important;
    color: #00788a !important;
    font-weight: 600 !important;
    box-shadow: none !important;
}
[class*="st-key-_batch_edit_btn_"] button:hover {
    background: #f0f8fa !important;
    border-color: #00788a !important;
}

/* 펼친 상태 라벨 caption — 인풋 바로 위 작게 표시 */
[class*="st-key-_batch_rate_row_"] [data-testid="stCaptionContainer"],
[class*="st-key-_batch_rate_row_"] .st-emotion-cache-caption {
    color: #7a8a90 !important;
    font-size: 0.72rem !important;
    margin-bottom: 2px !important;
    padding: 0 !important;
}

/* 일괄 정산 키워드 검색 입력 — 둥근 검색바 (SPH teal focus) */
.st-key-_batch_search input {
    border-radius: 12px !important;
    border: 1.5px solid #d4dce0 !important;
    background-color: #ffffff !important;
    padding: 0.55rem 0.9rem !important;
    font-size: 0.95rem !important;
    transition: border-color 0.15s, box-shadow 0.15s !important;
}
.st-key-_batch_search input:focus {
    border-color: #00788a !important;
    box-shadow: 0 0 0 3px rgba(0,120,138,0.14) !important;
}

/* 회사 row 안 인라인 환율 입력 — 우측 컬럼 안에 들어가므로 별도 들여쓰기
   없이 컨테이너 폭 100% 사용. overflow:visible 로 두 번째 줄(selectbox/
   직접 입력 text_input) 과 selectbox dropdown 메뉴가 정상 노출되도록 함. */
[class*="st-key-_batch_rate_row_"] {
    margin: 0 !important;
    padding: 6px 8px !important;
    background: #f6fafc !important;
    border-left: 3px solid #00788a !important;
    border-radius: 6px !important;
    width: 100% !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
    overflow: visible !important;
}
/* 둘째/셋째 줄 사이 간격 축소 */
[class*="st-key-_batch_rate_row_"] [data-testid="stVerticalBlock"] {
    gap: 0.3rem !important;
}
/* 4컬럼 gap 최소화 — selectbox 텍스트 공간 확보 */
[class*="st-key-_batch_rate_row_"] [data-testid="stHorizontalBlock"] {
    gap: 0.25rem !important;
    align-items: center !important;  /* 4필드 세로 중앙 정렬 */
}

/* ─── 4개 필드 (은행/날짜/문구/환율) 높이·padding·배경 통일 ───
   적용 단계: stXxxInput wrapper → baseweb input wrapper → input 자체
   모두 height 30px, padding 0, margin 0, box-sizing border-box.
   이전엔 input 만 30px 였고 wrapper 가 그대로라 wrapper 의 기본 padding/
   background 이 "input 아래 회색 영역" 으로 보였음. */

/* (1) Streamlit 위젯 wrapper — 위젯 컨테이너 자체 */
[class*="st-key-_batch_rate_row_"] [data-testid="stTextInput"],
[class*="st-key-_batch_rate_row_"] [data-testid="stNumberInput"],
[class*="st-key-_batch_rate_row_"] [data-testid="stDateInput"],
[class*="st-key-_batch_rate_row_"] [data-testid="stSelectbox"] {
    height: 30px !important;
    min-height: 30px !important;
    max-height: 30px !important;
    padding: 0 !important;
    margin: 0 !important;
    background: transparent !important;
    box-sizing: border-box !important;
    display: flex !important;
    align-items: center !important;
}

/* (2) baseweb input wrapper — input 을 감싸는 div (회색 영역의 정체) */
[class*="st-key-_batch_rate_row_"] [data-testid="stTextInput"] div[data-baseweb="input"],
[class*="st-key-_batch_rate_row_"] [data-testid="stNumberInput"] div[data-baseweb="input"],
[class*="st-key-_batch_rate_row_"] [data-testid="stDateInput"] div[data-baseweb="input"],
[class*="st-key-_batch_rate_row_"] [data-testid="stTextInput"] div[data-baseweb="base-input"],
[class*="st-key-_batch_rate_row_"] [data-testid="stNumberInput"] div[data-baseweb="base-input"],
[class*="st-key-_batch_rate_row_"] [data-testid="stDateInput"] div[data-baseweb="base-input"] {
    height: 30px !important;
    min-height: 30px !important;
    max-height: 30px !important;
    padding: 0 !important;
    margin: 0 !important;
    background-color: #ffffff !important;
    border: 1px solid #d4dce0 !important;
    border-radius: 6px !important;
    box-sizing: border-box !important;
    display: flex !important;
    align-items: center !important;
}

/* (3) <input> 자체 — replaced element 이므로 flex 적용 X, 좌우 padding 만 */
[class*="st-key-_batch_rate_row_"] [data-testid="stTextInput"] input,
[class*="st-key-_batch_rate_row_"] [data-testid="stNumberInput"] input,
[class*="st-key-_batch_rate_row_"] [data-testid="stDateInput"] input {
    height: 28px !important;       /* wrapper 30px - border 2px = 28px */
    min-height: 28px !important;
    max-height: 28px !important;
    padding: 0 8px !important;     /* 좌우 8px, 상하 0 — wrapper flex 가 정렬 */
    margin: 0 !important;
    font-size: 0.8rem !important;
    line-height: 1 !important;
    border: none !important;
    background: transparent !important;
    box-sizing: border-box !important;
    vertical-align: middle !important;
}

/* (4) Selectbox 의 baseweb wrapper — 다른 위젯과 height/border 통일 */
[class*="st-key-_batch_rate_row_"] [data-testid="stSelectbox"] div[data-baseweb="select"] {
    height: 30px !important;
    min-height: 30px !important;
    max-height: 30px !important;
    background-color: #ffffff !important;
    border-radius: 6px !important;
    box-sizing: border-box !important;
    /* baseweb 자체 보더 — 다른 input 과 동일하게 */
}
[class*="st-key-_batch_rate_row_"] [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    border: 1px solid #d4dce0 !important;
    border-radius: 6px !important;
}

[class*="st-key-_batch_rate_row_"] [data-testid="stWidgetLabel"] {
    display: none !important;
}

/* selectbox 텍스트 — 수직 정확 중앙 정렬 (재작성).
   ─────────────────────────────────────────────────────────────────
   이전 원인: <input> 에 display:flex; align-items:center; height:30px
   를 강제했는데, input 은 replaced element 라 flex 가 무시되고
   height/line-height 만 적용 → 텍스트가 박스 위쪽 baseline 에 정렬.
   해결 원칙(사용자 지시):
     1) padding-top == padding-bottom 으로 대칭 패딩
     2) line-height: 1 (멀티라인 회피 위해 height==line-height 금지)
     3) flex align-items: center 는 텍스트 *컨테이너* 에만 (input X)
   ───────────────────────────────────────────────────────────────── */
[class*="st-key-_batch_rate_row_"] div[data-baseweb="select"] {
    min-width: 0 !important;
    width: 100% !important;
}
/* 가장 바깥 wrapper — 박스 자체 높이 30px + 좌우 padding 만 */
[class*="st-key-_batch_rate_row_"] div[data-baseweb="select"] > div {
    /* 이전: height/min/max 30px, padding-top:0 / padding-bottom:0,
             display:flex; align-items:center
       → 이 자체는 OK. 유지. */
    height: 30px !important;
    min-height: 30px !important;
    max-height: 30px !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    padding-left: 8px !important;
    padding-right: 22px !important;
    margin: 0 !important;
    min-width: 0 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    box-sizing: border-box !important;
}
/* 값 표시 inner div — 텍스트가 들어있는 컨테이너(여기에 flex 중앙) */
[class*="st-key-_batch_rate_row_"] div[data-baseweb="select"] > div > div {
    /* 이전: height:30px + inline-flex+align-items:center
       → height 빼고 부모 flex 가 알아서 중앙 잡도록. */
    white-space: nowrap !important;
    overflow: visible !important;
    text-overflow: clip !important;
    font-size: 0.78rem !important;
    line-height: 1 !important;
    /* 상하 대칭 패딩 ((30 - 12.5) / 2 ≈ 8.75 → 8.5px). 박스 30px,
       텍스트 line-height ≈ 12.5px(0.78rem ≈ 12.5px). */
    padding: 8.5px 0 !important;
    margin: 0 !important;
    min-width: 0 !important;
    display: flex !important;
    align-items: center !important;
    box-sizing: border-box !important;
    /* height 강제 제거 — 부모(30px flex) 가 정렬 책임. */
}
/* placeholder 표시용 span (있는 경우) 은 텍스트 컨테이너와 동일 처리 */
[class*="st-key-_batch_rate_row_"] div[data-baseweb="select"] > div > div span {
    /* 이전: span 에 height:30px + display:flex 강제 → 텍스트 위쪽 정렬 원인 일부
       → height/flex 제거. 인라인 그대로 두고 부모 flex 에 맡김. */
    line-height: 1 !important;
    padding: 0 !important;
    margin: 0 !important;
    vertical-align: middle !important;
}
/* <input> 은 replaced element — flex 강제 금지 (display:flex 가 무시되어
   결국 line-height/height 만 먹는데 그게 위쪽 정렬 원인이었음).
   대신 좌우 패딩만 정리하고 height 와 line-height 를 자연스럽게. */
[class*="st-key-_batch_rate_row_"] div[data-baseweb="select"] input {
    /* 이전: display:flex; align-items:center; height:30px → 안 통함
       → flex 제거. height 도 부모가 책임. */
    height: auto !important;
    line-height: 1 !important;
    padding: 0 !important;
    margin: 0 !important;
    vertical-align: middle !important;
    background: transparent !important;
}
/* combobox role 요소 — BaseWeb 가 별도 wrapper 로 사용하는 경우 */
[class*="st-key-_batch_rate_row_"] div[data-baseweb="select"] [role="combobox"] {
    /* 이전: height:30px + display:flex 강제 → 부모 flex 와 중복
       → 부모에 맡기고 자체 padding 만 대칭. */
    padding: 0 !important;
    margin: 0 !important;
    line-height: 1 !important;
    display: flex !important;
    align-items: center !important;
}
/* selectbox 오른쪽 화살표 아이콘 더 작게 (12px) 해 텍스트 공간 추가 확보 */
[class*="st-key-_batch_rate_row_"] div[data-baseweb="select"] svg {
    width: 12px !important;
    height: 12px !important;
}
/* selectbox dropdown 메뉴(펼친 상태)는 부모 컨테이너 폭과 무관하게
   전체 옵션 텍스트가 정상 노출되도록 자동 너비. */
div[data-baseweb="popover"] li {
    white-space: nowrap !important;
    font-size: 0.82rem !important;
}
/* NumberInput ± 스피너 숨김 (좁은 공간에서 입력 영역만 노출) */
[class*="st-key-_batch_rate_row_"] [data-testid="stNumberInput"] [data-testid="stNumberInputStepDown"],
[class*="st-key-_batch_rate_row_"] [data-testid="stNumberInput"] [data-testid="stNumberInputStepUp"] {
    display: none !important;
}
/* DateInput 의 ::after 달력 아이콘 위치 + 패딩 보정 (좁은 폭 대응) */
[class*="st-key-_batch_rate_row_"] [data-testid="stDateInput"] [data-baseweb="input"]::after {
    font-size: 0.75rem !important;
    right: 4px !important;
}
[class*="st-key-_batch_rate_row_"] [data-testid="stDateInput"] [data-baseweb="input"] input {
    padding-right: 18px !important;
}

/* 첫 로딩 오버레이 — 데이터 로딩 동안 빈 화면 대신 표시. */
#loading-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(255, 255, 255, 0.9);
    z-index: 9999;
    display: flex;
    justify-content: center;
    align-items: center;
    flex-direction: column;
}
#loading-overlay .spinner {
    width: 50px;
    height: 50px;
    border: 4px solid #e0e0e0;
    border-top: 4px solid #00788a;  /* 앱 포인트 컬러 (SPH teal) */
    border-radius: 50%;
    animation: lo-spin 1s linear infinite;
}
#loading-overlay .loading-text {
    font-size: 1rem;
    color: #666;
    margin-top: 16px;
    font-weight: 500;
}
@keyframes lo-spin {
    0%   { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}
</style>
""", unsafe_allow_html=True)

# 첫 로딩 오버레이는 제거됨 — 사용자가 파일을 업로드한 직후 클라이언트
# 측 JS 가 오버레이를 표시하고, 페이지 렌더 완료 감지(MutationObserver) 시
# fade out. 서버에서 메시지를 받지 못하는 첫 진입엔 아무것도 안 표시됨.
from streamlit.components.v1 import html as _upload_listener_html
_upload_listener_html(
    """
<script>
(function() {
    const parentDoc = window.parent.document;
    // 마커 — 매 rerun 마다 components.html 이 재실행되지만 listener 는 1회만.
    if (parentDoc._uploadOverlayInit) return;
    parentDoc._uploadOverlayInit = true;

    function ensureOverlay() {
        let overlay = parentDoc.getElementById('loading-overlay');
        if (!overlay) {
            overlay = parentDoc.createElement('div');
            overlay.id = 'loading-overlay';
            overlay.innerHTML =
                '<div class="spinner"></div>' +
                '<div class="loading-text">데이터를 불러오는 중입니다...</div>';
            parentDoc.body.appendChild(overlay);
        }
        return overlay;
    }

    function showOverlay() {
        const overlay = ensureOverlay();
        overlay.style.display = 'flex';
        overlay.style.opacity = '1';

        // MutationObserver — 페이지 DOM 변경이 멈추면 렌더 완료로 판단 후 hide.
        let lastChange = Date.now();
        const root = parentDoc.querySelector('[data-testid="stAppViewContainer"]')
            || parentDoc.body;
        if (parentDoc._uploadObserver) {
            parentDoc._uploadObserver.disconnect();
        }
        const observer = new MutationObserver(() => { lastChange = Date.now(); });
        observer.observe(root, { childList: true, subtree: true });
        parentDoc._uploadObserver = observer;

        function checkDone() {
            if (Date.now() - lastChange > 800) {  // 800ms 무변경 → 완료
                observer.disconnect();
                overlay.style.transition = 'opacity 0.3s ease-out';
                overlay.style.opacity = '0';
                setTimeout(() => { overlay.style.display = 'none'; }, 300);
            } else {
                setTimeout(checkDone, 200);
            }
        }
        // 첫 체크는 1초 후 (streamlit 의 첫 rerun 메시지 도착까지 시간 확보).
        setTimeout(checkDone, 1000);
        // 안전망 — 30초 후엔 무조건 hide (파싱이 매우 오래 걸리는 경우 한계).
        setTimeout(() => {
            observer.disconnect();
            if (overlay) overlay.style.display = 'none';
        }, 30000);
    }

    // 파일 입력 change 이벤트 — capture 단계로 등록.
    // 모든 st.file_uploader 위젯(단가표/사용고지서) 에서 동작.
    parentDoc.addEventListener('change', (e) => {
        if (e.target && e.target.type === 'file'
            && e.target.closest('[data-testid="stFileUploader"]')
            && e.target.files && e.target.files.length > 0) {
            showOverlay();
        }
    }, true);
})();
</script>
    """,
    height=0,
)


# ── 헬퍼 함수 ─────────────────────────────────────────────────────────────────

def _render_loading(placeholder, percent: int, text: str) -> None:
    """정산 진행 중 화면 전체를 어둡게 덮는 fixed 오버레이 표시.

    상단 stHeader 를 CSS 로 숨겼기 때문에 Streamlit 기본 러닝 인디케이터가
    보이지 않는다. 그 대체로 화면 딤 + 스피너 + 진행 바를 placeholder 한
    곳에서 갱신한다. 진행률 갱신 시마다 같은 placeholder 에 다시 그려
    여러 오버레이가 쌓이지 않도록 한다.
    """
    pct = max(0, min(100, int(percent)))
    safe_text = (text or "").replace("<", "&lt;").replace(">", "&gt;")
    placeholder.markdown(
        f"""
<div class="sph-loading-overlay">
  <div class="sph-loading-spinner"></div>
  <div class="sph-loading-text">{safe_text}</div>
  <div class="sph-loading-bar"><div class="sph-loading-bar-fill" style="width:{pct}%"></div></div>
  <div class="sph-loading-pct">{pct}%</div>
</div>
""",
        unsafe_allow_html=True,
    )


def _load_master_df() -> pd.DataFrame:
    """master_data.csv → DataFrame. 없으면 빈 틀 반환."""
    if not MASTER_CSV.exists():
        return pd.DataFrame(columns=[
            "sku_id", "sku_name", "is_billable", "category",
            "free_usage_cap", "tier_number", "tier_limit", "tier_cpm",
        ])
    df = pd.read_csv(MASTER_CSV, dtype=str)
    df["is_billable"] = df["is_billable"].map(
        {"True": True, "False": False, "true": True, "false": False}
    ).fillna(False)
    df["free_usage_cap"] = (
        pd.to_numeric(df["free_usage_cap"], errors="coerce").fillna(0).astype(int)
    )
    for col in ("tier_number", "tier_limit", "tier_cpm"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def _df_to_sku_rows(df: pd.DataFrame) -> list[dict]:
    """DataFrame → load_sku_master() 소비 형식."""
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "sku_id":        str(r["sku_id"]),
            "sku_name":      str(r["sku_name"]),
            "is_billable":   bool(r["is_billable"]),
            "category":      str(r.get("category", "")),
            "free_usage_cap": int(r["free_usage_cap"]) if pd.notna(r.get("free_usage_cap")) else 0,
            "tier_number":   int(r["tier_number"]) if pd.notna(r.get("tier_number")) else None,
            "tier_limit":    int(r["tier_limit"])  if pd.notna(r.get("tier_limit"))  else None,
            "tier_cpm":      Decimal(str(r["tier_cpm"])) if pd.notna(r.get("tier_cpm")) else None,
        })
    return rows


def _save_master_df(df: pd.DataFrame) -> None:
    MASTER_CSV.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(MASTER_CSV, index=False)


# ── 캐시된 전처리: 같은 파일+월+계정 조합이면 pandas 파싱을 재사용 ────────────
# 정산 정확성 > 성능.
# 과거엔 @st.cache_data 로 감쌌지만 (파일 stat 기반 캐시 무효화까지 넣어도)
# Streamlit cache 가 코드 변경 / Hot reload 후 stale 결과를 반환하는 사례가
# 반복됐다 — Beyless Elevation(16,179) 누락 사고가 그 예. 정산 결과의 신뢰성이
# 최우선이므로 캐시를 쓰지 않고 매 호출 `preprocess_usage_file` 을 직접 실행.
# 호출부(`_cached_preprocess`) 시그니처는 하위호환을 위해 유지.
def _cached_preprocess_impl(tmp_path: str, billing_month: str,
                            company_filter: str | None,
                            file_stat: tuple[int, int]) -> list[dict]:
    del file_stat
    return preprocess_usage_file(
        tmp_path, billing_month, company_filter=company_filter
    )


def _cached_preprocess(tmp_path: str, billing_month: str,
                       company_filter: str | None) -> list[dict]:
    p = Path(tmp_path)
    try:
        st_ = p.stat()
        stat_key = (st_.st_mtime_ns, st_.st_size)
    except OSError:
        stat_key = (0, 0)
    return _cached_preprocess_impl(tmp_path, billing_month, company_filter, stat_key)


@st.cache_data(show_spinner=False)
def _cached_companies(tmp_path: str) -> list[str]:
    return extract_company_names(tmp_path)


@st.cache_data(show_spinner=False, ttl=600)
def _pdf_export_available() -> bool:
    """세션 시작 시 한 번만 Excel COM 가용성 체크 (15초 이내, 10분 캐시)."""
    try:
        from pdf_export import is_available
        return bool(is_available())
    except Exception:
        return False


@st.cache_data(show_spinner=False)
def _detect_billing_month(tmp_path: str) -> str | None:
    """CSV/Excel 상단 메타 영역에서 '인보이스 날짜' 를 찾아 YYYY-MM 으로 반환."""
    p = Path(tmp_path)
    if not p.exists():
        return None
    suffix = p.suffix.lower()

    def _parse_line(line: str) -> str | None:
        # "인보이스 날짜,2026-03-31," 또는 "인보이스 날짜,2026-03-31"
        # "Invoice date,2026-03-31" 도 허용
        lower = line.lower()
        if "인보이스 날짜" in line or "invoice date" in lower:
            for tok in line.replace("\t", ",").split(","):
                tok = tok.strip().strip('"')
                if len(tok) >= 7 and tok[4] == "-" and tok[:4].isdigit():
                    return tok[:7]
        return None

    if suffix in (".csv",):
        for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
            try:
                with open(p, encoding=enc, errors="strict") as f:
                    for i, line in enumerate(f):
                        if i > 30:   # 헤더 안에만 있음
                            break
                        ym = _parse_line(line)
                        if ym: return ym
                return None
            except (UnicodeDecodeError, LookupError):
                continue
    elif suffix in (".xlsx", ".xls"):
        try:
            df = pd.read_excel(p, header=None, nrows=20, dtype=str)
            for _, row in df.iterrows():
                line = ",".join("" if pd.isna(v) else str(v) for v in row)
                ym = _parse_line(line)
                if ym: return ym
        except Exception:
            return None
    return None


# ── 세션 상태 초기화 ──────────────────────────────────────────────────────────
if "master_df" not in st.session_state:
    st.session_state.master_df = _load_master_df()

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:0 0 12px;">
        <div style="
            background:rgba(255,255,255,0.13); border-radius:16px;
            padding:12px 0; margin-bottom:4px;
        ">
            <div style="font-size:2.2rem; margin-bottom:5px;">🗺️</div>
            <div style="font-weight:800; font-size:1.08rem; color:white; letter-spacing:-0.3px;">
                SPH GMP 정산
            </div>
            <div style="font-size:0.7rem; color:rgba(255,255,255,0.52); margin-top:3px;">
                Google Maps Platform Billing
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**📋 정산 기준 설정**")

    # ── 정산월: CSV 자동 감지 (직접 수정 불가) ──────────────────────────────
    _auto_bm = st.session_state.get("_auto_billing_month")
    billing_month = _auto_bm or ""
    st.text_input(
        "📅 정산월 (CSV 자동 감지)",
        value=billing_month or "CSV 업로드 필요",
        disabled=True,
        help="CSV 파일 상단의 '인보이스 날짜' 에서 자동 추출됩니다.",
    )

    # 은행·환율 표기 문구·날짜 는 본문 우측 "📝 환율 표기" 컨테이너로 이동.
    # (계정별 저장 + 선택·직접입력 + 달력 지원)
    margin_rate = 1.0

    st.divider()
    sku_count = st.session_state.master_df["sku_id"].nunique()
    st.caption(f"등록 SKU: {sku_count} 종")

# ── 통화 / 환율 기본값 — 메인 영역 위젯이 덮어씀 ──────────────────────────────
# 아래 col_right 블록에서 widget이 렌더될 때 실제 값으로 갱신되지만,
# uploaded_file 이 없거나 최초 렌더 시에도 NameError 방지용 sentinel 을 둔다.
currency      = "USD" if st.session_state.get("_detected_currency", "USD") == "USD" else "KRW"
exchange_rate = 0.0 if currency == "USD" else 1.0

# ── 환율 입력란 JS 포맷터 (4-digit auto-dot + decimal zero-pad on blur) ───
st.html(
    """
    <script>
    (function(){
      function patchRate(){
        try {
          var doc = window.parent.document;
          var inputs = doc.querySelectorAll('input[aria-label*="환율"]');
          inputs.forEach(function(inp){
            if (inp._ratePatched) return;
            inp._ratePatched = true;
            var nativeSetter = Object.getOwnPropertyDescriptor(
              window.parent.HTMLInputElement.prototype, 'value'
            ).set;
            function setVal(v){
              nativeSetter.call(inp, v);
              inp.dispatchEvent(new Event('input', { bubbles: true }));
            }
            inp.addEventListener('input', function(){
              var digits = inp.value.replace(/\\D/g,'');
              if (digits.length > 6) digits = digits.substring(0, 6);
              var formatted;
              if (digits.length <= 4) {
                formatted = digits;
              } else {
                formatted = digits.substring(0, 4) + '.' + digits.substring(4);
              }
              if (formatted !== inp.value) setVal(formatted);
            });
            inp.addEventListener('blur', function(){
              var v = inp.value;
              if (!v) return;
              var digits = v.replace(/\\D/g,'');
              if (digits.length === 0) return;
              var intp = digits.substring(0, 4);
              var decp = digits.substring(4, 6);
              if (intp.length < 4) return;   // 정수부 미완성이면 그대로 둠
              decp = (decp + '00').substring(0, 2);
              var formatted = intp + '.' + decp;
              if (formatted !== v) setVal(formatted);
            });
          });
        } catch(e) {}
      }
      new MutationObserver(patchRate).observe(
        window.parent.document.body, { subtree: true, childList: true }
      );
      patchRate();
    })();
    </script>
    """,
    unsafe_allow_javascript=True,
)

# ── 드롭다운(selectbox) 키보드 네비게이션 스크롤 패치 ────────────────────────
# BaseWeb 의 listbox 는 키보드 방향키로 하이라이트가 이동해도 화면 밖으로
# 나가면 자동 스크롤이 안 된다. listbox 가 열릴 때마다 하이라이트 변경을
# 관찰하고 scrollIntoView 로 시야 안으로 끌어온다.
st.html(
    """
    <script>
    (function(){
      function highlightedItem(listbox){
        return listbox.querySelector('li[aria-selected="true"]')
            || listbox.querySelector('[role="option"][aria-selected="true"]')
            || listbox.querySelector('[data-highlighted="true"]')
            || listbox.querySelector('li.active');
      }
      function attachScroll(listbox){
        if (listbox._kbdScrollPatched) return;
        listbox._kbdScrollPatched = true;
        var scroll = function(){
          var item = highlightedItem(listbox);
          if (item && typeof item.scrollIntoView === 'function') {
            item.scrollIntoView({ block: 'nearest', inline: 'nearest' });
          }
        };
        var obs = new MutationObserver(scroll);
        obs.observe(listbox, {
          subtree: true,
          attributes: true,
          attributeFilter: ['aria-selected', 'data-highlighted', 'class']
        });
        scroll();
      }
      function scan(){
        try {
          var doc = window.parent.document;
          var listboxes = doc.querySelectorAll(
            'ul[role="listbox"], div[role="listbox"]'
          );
          listboxes.forEach(attachScroll);
        } catch(e){}
      }
      new MutationObserver(scan).observe(
        window.parent.document.body, { subtree: true, childList: true }
      );
      scan();
    })();
    </script>
    """,
    unsafe_allow_javascript=True,
)

# 설정 변경 감지는 col_right 내부(통화·환율 위젯 직후)로 이동됨.

# ── 메인 헤더 ─────────────────────────────────────────────────────────────────
st.markdown("""
<div style="
    display:flex; align-items:center; gap:18px;
    padding:0 2px 16px;
    border-bottom:2px solid #d4e2e6;
    margin-bottom:16px;
">
    <div style="
        background:linear-gradient(135deg,#00788a,#005060);
        color:white; font-size:1rem; font-weight:900;
        width:56px; height:56px; border-radius:16px;
        display:flex; align-items:center; justify-content:center;
        box-shadow:0 6px 18px rgba(0,120,138,0.32); flex-shrink:0;
        letter-spacing:-0.5px;
    ">SPH</div>
    <div>
        <div style="
            font-size:1.48rem; font-weight:800; color:#1a3540;
            line-height:1.2; letter-spacing:-0.5px;
        ">GMP 정산 자동화 시스템</div>
        <div style="font-size:0.82rem; color:#5a8290; margin-top:3px;">
            Google Maps Platform Billing Automation · SPH Infosolution
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Selectbox 포커스 시 텍스트 전체 선택 (Ctrl+A 지원) ───────────────────────
st.html("""
<script>
(function () {
    function injectValueAndSelectAll(inp) {
        if (inp.value !== '') return false;
        var selectEl = inp.closest('[data-baseweb="select"]');
        if (!selectEl) return false;
        var nodes = selectEl.querySelectorAll('span, div');
        var displayText = '';
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (el.children.length === 0 && !el.contains(inp) && el !== inp) {
                var t = el.textContent.trim();
                if (t) { displayText = t; break; }
            }
        }
        if (!displayText) return false;
        var nativeSetter = Object.getOwnPropertyDescriptor(
            window.parent.HTMLInputElement.prototype, 'value'
        ).set;
        nativeSetter.call(inp, displayText);
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        setTimeout(function () { inp.select(); }, 0);
        return true;
    }

    function patch() {
        try {
            var doc = window.parent.document;
            doc.querySelectorAll('[data-baseweb="select"] input').forEach(function (inp) {
                if (inp._salPatched) return;
                inp._salPatched = true;
                inp.addEventListener('keydown', function (e) {
                    if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
                        if (injectValueAndSelectAll(inp)) {
                            e.preventDefault();
                            e.stopPropagation();
                        }
                    }
                });
            });
        } catch (e) {}
    }

    new MutationObserver(patch).observe(
        window.parent.document.body,
        { subtree: true, childList: true }
    );
    patch();
})();
</script>
""", unsafe_allow_javascript=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 통합 정산 실행 (flat — 탭 제거)
# ═══════════════════════════════════════════════════════════════════════════════
if True:

    # ── ① 단가표(GMP Price List) 첨부 — USD/KRW 통화별 분리 ───────────────
    # Streamlit Community Cloud 휴면 wake 시 컨테이너 로컬 디스크가 초기화되므로,
    # 세션 시작 시 GitHub repo 에서 단가표를 복원한다(설정 있을 때만).
    if github_storage.is_configured() and "_price_list_pulled" not in st.session_state:
        for _local, _remote in _PRICE_LIST_REMOTES.items():
            github_storage.pull(_remote, _local)
        st.session_state["_price_list_pulled"] = True

    # 레거시(saved_price_list.xlsx) 가 있으면 감지 통화에 맞춰 신규 파일로 자동
    # 이관. 두 신규 파일 중 하나라도 이미 있으면 이관 스킵.
    if PRICE_LIST_SAVED.exists() and not (
        PRICE_LIST_SAVED_USD.exists() or PRICE_LIST_SAVED_KRW.exists()
    ):
        try:
            _legacy_cur = detect_price_list_currency(PRICE_LIST_SAVED)
        except Exception:
            _legacy_cur = "USD"
        _target = (PRICE_LIST_SAVED_USD if _legacy_cur == "USD"
                   else PRICE_LIST_SAVED_KRW)
        try:
            _target.parent.mkdir(parents=True, exist_ok=True)
            _target.write_bytes(PRICE_LIST_SAVED.read_bytes())
        except Exception:
            pass

    def _persist_price_list(uploaded, save_path, session_tag: str) -> None:
        """업로드된 단가표를 지정 경로에 저장하고 한 번만 flash 메시지.

        설정이 되어 있으면 GitHub repo 에도 함께 커밋해 휴면 wake 후에도 단가표
        가 유지되도록 한다.
        """
        _key = f"{uploaded.name}_{uploaded.size}"
        if st.session_state.get(f"_saved_price_key_{session_tag}") != _key:
            uploaded.seek(0)
            save_path.parent.mkdir(parents=True, exist_ok=True)
            save_path.write_bytes(uploaded.read())
            remote = _PRICE_LIST_REMOTES.get(save_path)
            if remote and github_storage.is_configured():
                github_storage.push(
                    save_path, remote, f"단가표 업데이트: {uploaded.name}"
                )
            st.session_state[f"_saved_price_key_{session_tag}"] = _key
            st.session_state[f"_price_flash_{session_tag}"] = uploaded.name
            st.rerun()

    # 좌→우 순서: 사용고지서, 달러 단가표, 원화 단가표.
    # 사용자 워크플로우상 사용고지서 업로드가 매월 가장 먼저 일어나므로
    # 가장 왼쪽(시선 진입 위치)에 배치한다.
    _col_invoice, _col_pl_usd, _col_pl_krw = st.columns(3, gap="medium")

    with _col_pl_usd:
        _uploaded_price_usd = st.file_uploader(
            "📋 달러($) 단가표",
            type=["xlsx"],
            key="price_list_uploader_usd",
            help="USD 기준 GMP Price List. 통화를 달러로 선택하면 사용됩니다.",
        )
        if _uploaded_price_usd is not None:
            _persist_price_list(_uploaded_price_usd, PRICE_LIST_SAVED_USD, "usd")
        if PRICE_LIST_SAVED_USD.exists() and _uploaded_price_usd is None:
            _c1, _c2 = st.columns([5, 1], vertical_alignment="center")
            with _c1:
                st.caption(f"📂 저장됨: `{PRICE_LIST_SAVED_USD.name}`")
            with _c2:
                if st.button("🗑", key="del_price_list_usd",
                             help="저장된 달러 단가표 삭제"):
                    PRICE_LIST_SAVED_USD.unlink(missing_ok=True)
                    if github_storage.is_configured():
                        github_storage.delete(
                            _PRICE_LIST_REMOTES[PRICE_LIST_SAVED_USD],
                            "달러 단가표 삭제",
                        )
                    st.session_state.pop("_saved_price_key_usd", None)
                    st.rerun()
        if _f_usd := st.session_state.pop("_price_flash_usd", None):
            st.success(f"✅ 달러 단가표 저장 — {_f_usd}")

    with _col_pl_krw:
        _uploaded_price_krw = st.file_uploader(
            "📋 원화(₩) 단가표",
            type=["xlsx"],
            key="price_list_uploader_krw",
            help="KRW 기준 GMP Price List. 통화를 원화로 선택하면 사용됩니다.",
        )
        if _uploaded_price_krw is not None:
            _persist_price_list(_uploaded_price_krw, PRICE_LIST_SAVED_KRW, "krw")
        if PRICE_LIST_SAVED_KRW.exists() and _uploaded_price_krw is None:
            _c1, _c2 = st.columns([5, 1], vertical_alignment="center")
            with _c1:
                st.caption(f"📂 저장됨: `{PRICE_LIST_SAVED_KRW.name}`")
            with _c2:
                if st.button("🗑", key="del_price_list_krw",
                             help="저장된 원화 단가표 삭제"):
                    PRICE_LIST_SAVED_KRW.unlink(missing_ok=True)
                    if github_storage.is_configured():
                        github_storage.delete(
                            _PRICE_LIST_REMOTES[PRICE_LIST_SAVED_KRW],
                            "원화 단가표 삭제",
                        )
                    st.session_state.pop("_saved_price_key_krw", None)
                    st.rerun()
        if _f_krw := st.session_state.pop("_price_flash_krw", None):
            st.success(f"✅ 원화 단가표 저장 — {_f_krw}")

    # price_list_file 의 **임시** 초기값 — billable_skus 계산 등 currency 확정
    # 전에 참조되는 경로를 위해 존재하는 아무 파일이든 사용. 실제 정산에 쓰일
    # 최종 파일은 "통화·환율" 라디오 확정 직후에 재할당된다.
    if _uploaded_price_usd is not None:
        _uploaded_price_usd.seek(0)
        price_list_file = _uploaded_price_usd
    elif _uploaded_price_krw is not None:
        _uploaded_price_krw.seek(0)
        price_list_file = _uploaded_price_krw
    elif PRICE_LIST_SAVED_USD.exists():
        price_list_file = PRICE_LIST_SAVED_USD
    elif PRICE_LIST_SAVED_KRW.exists():
        price_list_file = PRICE_LIST_SAVED_KRW
    elif PRICE_LIST_SAVED.exists():
        price_list_file = PRICE_LIST_SAVED
    else:
        price_list_file = None
        st.caption("단가표 미첨부 — 통화에 맞는 단가표를 업로드하세요.")

    # ── ② 사용고지서 업로드 (원화 단가표 옆) ──────────────────────────────────
    with _col_invoice:
        uploaded_file = st.file_uploader(
            "📂 구글 Maps 플랫폼 사용고지서",
            type=["csv"],
            help="구글 Maps 플랫폼 콘솔 → 결제 → 청구서 내보내기 파일을 업로드하세요.",
        )

    if uploaded_file is not None:
        # ── 새 파일 감지 → 임시 저장 ──────────────────────────────────────────
        file_key = f"{uploaded_file.name}_{uploaded_file.size}"
        if st.session_state.get("_file_key") != file_key:
            old = st.session_state.get("_tmp_path")
            if old:
                Path(old).unlink(missing_ok=True)

            suffix = Path(uploaded_file.name).suffix
            uploaded_file.seek(0)
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = Path(tmp.name)

            st.session_state._file_key  = file_key
            st.session_state._tmp_path  = str(tmp_path)
            st.session_state._companies = _cached_companies(str(tmp_path))
            # 정산월 자동 감지 — CSV '인보이스 날짜' 에서 YYYY-MM 추출
            _detected_bm = _detect_billing_month(str(tmp_path))
            if _detected_bm:
                st.session_state._auto_billing_month = _detected_bm
            st.session_state.pop("_last_result", None)  # 이전 결과 초기화
            st.session_state.pop("_pending_result", None)
            st.session_state.pop("_pending_auto_dl_key", None)
            st.session_state.pop("_pending_hidden_paid", None)
            st.rerun()   # 사이드바의 자동 감지 정산월 표시를 즉시 반영

        tmp_input_path = Path(st.session_state._tmp_path)
        companies      = st.session_state._companies

        # ── 🗂 전체 일괄 정산 모드 토글 ───────────────────────────────────────
        # 켜져 있으면 회사 셀렉트박스/단일 정산 UI 를 모두 우회하고 일괄
        # 정산 UI 로 진입한다. 정산 엔진은 단일 모드와 동일 함수 체인 사용
        # → 같은 입력에 대해 같은 결과 보장.
        _batch_mode = st.toggle(
            "전체 일괄 정산 모드",
            value=st.session_state.get("_batch_mode_toggle", True),
            key="_batch_mode_toggle",
            help=(
                "체크한 회사들을 한 번에 정산해 회사별 폴더로 zip 다운로드. "
                "각 회사의 저장된 설정(과금방식/소수점/최소사용비용/환율표기/"
                "미노출SKU/직접등록/SKU순서) 을 그대로 사용합니다."
            ),
        )
        if _batch_mode:
            # 사이드바에서 이미 결정된 currency / price_list_file 을 그대로 사용.
            # billable_skus 만 추가로 계산 (단일 모드 흐름과 동일 함수).
            _bm_billable: set | None = None
            if price_list_file is not None:
                try:
                    _bm_billable = get_billable_sku_names(price_list_file)
                except Exception:
                    _bm_billable = None

            render_batch_billing_ui(
                tmp_input_path = tmp_input_path,
                companies      = sorted(companies or [], key=lambda c: str(c).lower()),
                billing_month  = billing_month,
                price_list_file= price_list_file,
                currency       = currency,
                billable_skus  = _bm_billable,
            )
            st.stop()

        # ── 결제 계정 선택 (전체 폭) ─────────────────────────────────────────
        # 스크롤로 셀렉트박스가 화면 밖으로 나가면 상단에 떠있는(floating) 모드.
        # 화면에 다시 들어오면 원래 자리로 복귀. IntersectionObserver 로 원본
        # 위치를 추적하고 CSS class 토글만 한다 — element 자체는 그대로 두므로
        # streamlit selectbox 의 검색/선택 동작이 floating 상태에서도 정상.
        with st.container(key="sticky_account_select"):
            st.markdown("#### 정산 대상 선택")
            if companies:
                companies = sorted(companies, key=lambda c: str(c).lower())
                # 검색창 — substring 매칭 (정규화 키 + 소문자) 으로 일괄 정산 UI 와 동일.
                # Streamlit selectbox 기본은 부분수열(subsequence) 매칭이라 의도치 않은
                # 회사가 결과에 끼는 문제가 있어, 직접 필터링 후 옵션을 좁혀준다.
                _acc_search = st.text_input(
                    "🔍 결제 계정 검색",
                    value="",
                    key="_account_search",
                    placeholder="회사명 일부를 입력하면 해당 회사만 표시됩니다",
                    label_visibility="collapsed",
                )
                if _acc_search:
                    _qn = _norm_account_key(_acc_search)
                    _ql = _acc_search.lower()
                    _filtered = [
                        c for c in companies
                        if (_qn and _qn in _norm_account_key(c))
                        or _ql in c.lower()
                    ]
                else:
                    _filtered = companies
                if not _filtered:
                    st.caption(f"🔍 '{_acc_search}' 와 일치하는 결제계정이 없습니다.")
                    selected_company = None
                else:
                    # selectbox 에 key 가 있으면 session_state 의 기존 값이 옵션에
                    # 없을 때 StreamlitAPIException — 필터 변경으로 사라진 값은
                    # 미리 제거해 첫 항목이 선택되도록 한다.
                    _saved_acc = st.session_state.get("_account_select")
                    if _saved_acc is not None and _saved_acc not in _filtered:
                        st.session_state.pop("_account_select", None)
                    selected_company = st.selectbox(
                        "결제 계정 (Billing Account Name)",
                        options=_filtered,
                        key="_account_select",
                    )
            else:
                selected_company = None
                st.info("파일에서 결제 계정 정보를 찾을 수 없습니다. 전체 데이터를 처리합니다.")

        # 회사 변경 감지 → 다음 rerun 에 SKU 노출 순서 패널로 자동 스크롤.
        # floating 상태에서 검색해도 결과 위치가 sku_order_panel 시작점으로
        # 정렬되어, 상단 floating bar 바로 아래에 SKU 영역이 보인다.
        _prev_cmp = st.session_state.get("_prev_selected_company")
        if selected_company != _prev_cmp:
            st.session_state["_prev_selected_company"] = selected_company
            if _prev_cmp is not None and selected_company is not None:
                st.session_state["_scroll_to_sku_panel"] = True

        # 회사 변경 직후 1회 발사 — floating bar 바로 아래에 sku_order_panel 시작점이
        # 오도록 자동 스크롤. sku_order_panel 이 DOM 에 들어올 때까지 짧게 polling.
        if st.session_state.pop("_scroll_to_sku_panel", False):
            st.html(
                """
                <script>
                (function(){
                  const doc = window.parent.document;
                  const scroller = window.parent;
                  function go(attempt){
                    const panel = doc.querySelector('.st-key-sku_order_panel');
                    if (!panel){
                      if (attempt < 30) setTimeout(() => go(attempt + 1), 80);
                      return;
                    }
                    const orig = doc.querySelector('.st-key-sticky_account_select');
                    const isFloating = orig && orig.classList.contains('is-floating');
                    const offset = (isFloating ? orig.offsetHeight : 0) + 8;
                    const rect = panel.getBoundingClientRect();
                    const top  = (scroller.pageYOffset || 0) + rect.top - offset;
                    scroller.scrollTo({top: top, behavior: 'smooth'});
                  }
                  setTimeout(() => go(0), 120);
                })();
                </script>
                """,
                unsafe_allow_javascript=True,
            )

        st.html(
            """
            <script>
            (function(){
              const doc = window.parent.document;

              // ───── 좌측 SKU 컬럼 너비를 측정해 floating/평소 양쪽에 적용 ─────
              function applyDims(){
                const orig = doc.querySelector('.st-key-sticky_account_select');
                if (!orig) return;
                const cols = doc.querySelectorAll('[data-testid="stColumn"]');
                if (cols.length >= 1){
                  const rect = cols[0].getBoundingClientRect();
                  if (rect.width > 50){
                    if (orig.classList.contains('is-floating')){
                      orig.style.left  = rect.left + 'px';
                      orig.style.width = rect.width + 'px';
                      orig.style.maxWidth = rect.width + 'px';
                    } else {
                      orig.style.left = '';
                      orig.style.width = '';
                      orig.style.maxWidth = rect.width + 'px';
                    }
                  }
                }
              }

              // ───── 검색 input 드래그 시 페이지 본문 selection 차단 ─────
              // 입력창의 텍스트를 마우스로 드래그할 때 마우스가 selectbox 밖으로
              // 나가면 페이지 본문까지 selection 이 확장되어, Del 키로 input
              // 텍스트를 지울 수 없게 된다. input mousedown→mouseup 동안만 body
              // user-select 를 비활성화해 본문 selection 만 차단한다. input
              // 자체의 native 텍스트 선택은 user-select 와 무관하게 정상 동작.
              function bindInputGuard(inp){
                if (inp.__selGuardBound) return;
                inp.__selGuardBound = true;
                inp.addEventListener('mousedown', function(){
                  const body = doc.body;
                  const prev = body.style.userSelect;
                  const prevWk = body.style.webkitUserSelect;
                  body.style.userSelect = 'none';
                  body.style.webkitUserSelect = 'none';
                  function release(){
                    body.style.userSelect = prev || '';
                    body.style.webkitUserSelect = prevWk || '';
                    doc.removeEventListener('mouseup', release);
                    doc.removeEventListener('blur', release, true);
                  }
                  doc.addEventListener('mouseup', release);
                  // input 포커스가 빠지는 경우(드롭다운 옵션 클릭 등)도 정리.
                  doc.addEventListener('blur', release, true);
                });
              }
              function scanInputs(){
                doc
                  .querySelectorAll('.st-key-sticky_account_select input')
                  .forEach(bindInputGuard);
              }

              function setup(){
                const orig = doc.querySelector('.st-key-sticky_account_select');
                if (!orig || orig.dataset._floatInit) return;

                // 원본 자리 비움 방지용 placeholder (floating 모드에서만 노출).
                let ph = orig.previousElementSibling;
                if (!ph || !ph.classList || !ph.classList.contains('account-select-ph')){
                  ph = doc.createElement('div');
                  ph.className = 'account-select-ph';
                  ph.style.display = 'none';
                  orig.parentNode.insertBefore(ph, orig);
                }

                // sentinel: 원본의 "원래 위치" 를 표시 — sentinel 이 화면 밖이면 floating.
                let sent = orig.previousElementSibling.previousElementSibling;
                if (!sent || !sent.classList || !sent.classList.contains('account-select-sentinel')){
                  sent = doc.createElement('div');
                  sent.className = 'account-select-sentinel';
                  sent.style.cssText = 'height:1px;pointer-events:none;';
                  orig.parentNode.insertBefore(sent, ph);
                }

                const io = new IntersectionObserver(function(entries){
                  entries.forEach(function(e){
                    if (e.isIntersecting){
                      orig.classList.remove('is-floating');
                      ph.style.display = 'none';
                    } else {
                      // 원본 높이만큼 placeholder 채워 layout shift 방지.
                      ph.style.height = orig.offsetHeight + 'px';
                      ph.style.display = 'block';
                      orig.classList.add('is-floating');
                    }
                    applyDims();
                  });
                }, { threshold: 0, rootMargin: '0px 0px 0px 0px' });
                io.observe(sent);
                // 윈도우 리사이즈 시 너비 재측정
                window.parent.addEventListener('resize', applyDims);
                orig.dataset._floatInit = '1';
              }
              setup();
              applyDims();
              scanInputs();
              // Streamlit rerun 으로 DOM 교체되면 다시 바인딩 + 너비 재측정.
              setInterval(function(){ setup(); applyDims(); scanInputs(); }, 700);
            })();
            </script>
            """,
            unsafe_allow_javascript=True,
        )

        # ── 좌: 드래그앤드롭 | 우: 다운로드 체크박스 + 정산 시작 ────────────
        _order_account_key = selected_company or "__ALL__"
        # 단가표에서 추출한 이름 집합 — invoice_generator 내부에서 tier 단가
        # 매핑에 쓰일 수 있도록 넘겨 줄 용도. UI 필터로는 쓰지 않는다
        # (회사별로 사용 SKU 가 달라 교집합을 걸면 실측 사용 SKU 가 숨겨짐).
        _billable_skus: set[str] | None = None
        if price_list_file is not None:
            _billable_skus = get_billable_sku_names(price_list_file)
        # 선택 계정의 CSV 에서 usage > 0 인 모든 SKU 를 그대로 노출 —
        # 하드코딩 화이트리스트 없음. 회사마다 쓰는 제품이 다르므로
        # 실측 사용 데이터가 유일한 기준.
        _found_skus = _unique_skus_for_account(
            str(tmp_input_path), billing_month, selected_company
        )

        # ── 진단(토글) — 업로드 CSV 에서 SKU 가 어디서 사라지는지 추적 ──
        # 사용자가 "CSV에 있는데 목록에 안 뜬다" 를 제기할 때 원본 vs 전처리
        # 결과를 화면에서 즉시 비교할 수 있도록 한다. 캐시된 경로가 아닌
        # 방금 업로드된 임시파일을 직접 다시 읽는다 (캐시 우회).
        with st.expander("🔍 진단 — 업로드 CSV 원본 / 전처리 결과 (SKU가 빠지는 경우 열어보세요)", expanded=False):
            st.caption(f"업로드 임시 경로: `{tmp_input_path}`")

            # 1) 원본에서 sku_name 후보 문자열을 전체 검색 (대소문자 무시)
            _search_term = st.text_input(
                "원본에서 찾을 문자열 (sku_name 또는 그 일부)",
                value="elevation",
                key="_diag_search_term",
            )
            _raw_hits: list[str] = []
            if _search_term:
                try:
                    with open(str(tmp_input_path), "rb") as _f:
                        _raw_bytes = _f.read()
                    # CSV 인코딩 후보 순서: UTF-8 BOM → UTF-8 → CP949 → latin-1
                    _raw_text = None
                    for _enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr", "latin-1"):
                        try:
                            _raw_text = _raw_bytes.decode(_enc)
                            break
                        except UnicodeDecodeError:
                            continue
                    if _raw_text is None:
                        _raw_text = _raw_bytes.decode("latin-1", errors="replace")
                    _tl = _search_term.lower()
                    _raw_hits = [
                        ln for ln in _raw_text.splitlines() if _tl in ln.lower()
                    ]
                except Exception as _e:
                    st.error(f"원본 읽기 실패: {_e}")
                st.write(f"원본에서 **'{_search_term}'** 포함 행: **{len(_raw_hits)}개**")
                if _raw_hits:
                    st.code("\n".join(_raw_hits[:30]), language="text")

            # 2) 전처리 결과(선택 계정 필터 적용) SKU 별 usage 합
            st.markdown("---")
            try:
                from billing.preprocessor import preprocess_usage_file as _pp_fn
                _dbg_rows = _pp_fn(
                    str(tmp_input_path), billing_month,
                    company_filter=selected_company,
                )
                _dbg_agg: dict[str, int] = {}
                for _r in _dbg_rows:
                    _nm = (_r.get("sku_name") or "").strip() or "(빈값)"
                    _dbg_agg[_nm] = _dbg_agg.get(_nm, 0) + int(
                        _r.get("usage_amount") or 0
                    )
                import pandas as _pd
                _dbg_df = _pd.DataFrame(
                    [{"sku_name": _k, "usage_sum": _v}
                     for _k, _v in sorted(
                         _dbg_agg.items(), key=lambda x: -x[1]
                     )]
                )
                st.write(
                    f"전처리 결과 (계정 필터: **{selected_company or '전체'}**) "
                    f"— 행 {len(_dbg_rows)}개, 고유 SKU **{len(_dbg_agg)}종**"
                )
                st.dataframe(_dbg_df, use_container_width=True, height=260)
            except Exception as _e:
                st.error(f"전처리 재실행 실패: {_e}")

        col_left, col_right = st.columns([1, 1], vertical_alignment="top", gap="large")

        # ═══ 좌측: SKU 순서 드래그앤드롭 ═══════════════════════════════════
        # 세 영역(SKU 순서 / 통화·환율 / 다운로드 옵션) 을 각각 bordered
        # container 로 명확히 구분해서 시각적 그룹핑을 만든다.
        with col_left, st.container(border=True, key="sku_order_panel"):
            # 마스터 SKU 직접등록 목록 (이번 CSV 사용량 0 이어도 노출하고 싶은
            # 항목 — 사용자가 multiselect 로 수동 선택). 계정별 저장.
            _saved_manual_all   = _load_manual_skus_map()
            _manual_skus_saved  = _lookup_account(
                _saved_manual_all, _order_account_key) or []

            # CSV 사용량과 직접등록이 모두 비어 있어도 드래그 영역·직접등록
            # 영역은 노출 — 사용자가 마스터에서 SKU 를 직접 추가할 수 있게.
            if True:
                _saved_orders = _load_saved_orders()
                _saved_for_this = _lookup_account(
                    _saved_orders, _order_account_key) or []

                # 미노출(hidden) SKU 는 sku_order 패널 자체에서도 빼서 출력
                # 대상만 보이게 한다 — 출력 결과(엑셀 3개)와 패널 표시를 일관.
                # hidden 패널에서 복원(X)하면 다음 rerun 에 다시 패널에 등장.
                _saved_hidden_pre = _lookup_account(
                    _load_hidden_skus_map(), _order_account_key) or []
                _hidden_set_for_order = set(_saved_hidden_pre)

                # saved_orders(=GMP 순서)를 walk 하면서 각 항목을 existing/manual 로 분류
                # → 사용량 없는 SKU 도 GMP 순서상 원래 위치에 끼워넣어 노출
                _found_set   = set(_found_skus)
                _manual_set  = set(_manual_skus_saved)
                _saved_set   = set(_saved_for_this)
                _new_items   = [
                    n for n in _found_skus
                    if n not in _saved_set and n not in _hidden_set_for_order
                ]
                _NEW_COUNT   = len(_new_items)
                # saved_order 에 없는 manual 만 별도(거의 안 생기는 edge case — 끝에 붙임)
                _manual_extra = [
                    m for m in _manual_skus_saved
                    if m not in _saved_set
                    and m not in _found_set
                    and m not in _hidden_set_for_order
                ]
                # 디스플레이상 manual 로 분류될 항목 = CSV 에 없는 manual SKU 전체
                # (X 버튼 UI 에서도 사용)
                _manual_only  = [
                    m for m in _manual_skus_saved
                    if m not in _found_set and m not in _hidden_set_for_order
                ]
                _MANUAL_COUNT = len(_manual_only)

                _has_existing = any(
                    s in _found_set and s not in _hidden_set_for_order
                    for s in _saved_for_this
                )

                # 패널 표시 항목 수 — hidden 제외 기준으로 카운트 (출력 엑셀과
                # 일치). _found_skus 중 hidden 인 항목은 빼고 + CSV 사용량 0
                # 직접등록(_manual_only, 이미 hidden 제외됨) 더함.
                _visible_count = (
                    sum(1 for s in _found_skus if s not in _hidden_set_for_order)
                    + _MANUAL_COUNT
                )

                st.markdown(f"#### 📋 엑셀 SKU 노출 순서 ({_visible_count}개)")
                if not _found_skus and not _manual_skus_saved:
                    st.caption(
                        "이번 CSV 에 사용량이 없습니다. 아래 **마스터에서 직접 "
                        "SKU 추가** 로 항목을 직접 등록하거나 CSV 를 확인해 주세요."
                    )
                elif _new_items and _has_existing:
                    st.caption(
                        "연한 초록 배경 = **[신규 항목]**. 드래그로 위치 조정 후 "
                        "**현재 순서 저장** 을 누르세요."
                    )
                elif _new_items and not _has_existing:
                    st.caption(
                        "모두 새로 발견된 SKU입니다. 드래그로 순서 지정 후 "
                        "**현재 순서 저장** 을 누르세요."
                    )
                else:
                    st.caption("드래그하여 엑셀에 나올 순서를 조정하세요.")

                # 신규 항목은 라벨에 [신규 항목] prefix, 직접등록은 [직접등록]
                # prefix 로 표시 (저장 전 제거).
                _NEW_PREFIX    = "[신규 항목] "
                _MANUAL_PREFIX = "[직접등록] "
                # GMP(saved_order) 순서대로 walk → CSV에 있으면 existing, 아니면 manual
                # hidden 안 항목은 sku_order 패널에서 제외 (위에서 _new_items
                # / _manual_extra 도 이미 hidden 제외).
                _initial: list[str] = []
                for _s in _saved_for_this:
                    if _s in _hidden_set_for_order:
                        continue
                    if _s in _found_set:
                        _initial.append(_s)
                    elif _s in _manual_set:
                        _initial.append(f"{_MANUAL_PREFIX}{_s}")
                    # saved_order 에 있으나 CSV·manual 어디에도 없으면 스킵
                # CSV 에 새로 등장한 SKU 는 끝에 부착(GMP 순서를 모름)
                _initial += [f"{_NEW_PREFIX}{n}" for n in _new_items]
                # saved_order 에 없고 manual 에만 존재하는 edge case 도 끝에
                _initial += [f"{_MANUAL_PREFIX}{m}" for m in _manual_extra]

                # sortable 컴포넌트 key: 입력 "항목 집합" 이 변경되면 새 key로
                # 캐시 초기화. 순서까지 fingerprint 에 넣으면 자동 저장 → rerun
                # 사이클에서 순서가 바뀔 때마다 key 가 바뀌어 streamlit_sortables
                # 가 컴포넌트를 재초기화 → 두 번째 드래그 결과가 손실된다.
                # sorted() 로 순서 영향을 제거하고 항목 추가/제거에만 반응.
                _items_fingerprint = hashlib.md5(
                    "\x1f".join(sorted(_initial)).encode("utf-8")
                ).hexdigest()[:10]
                _order_state_key = (
                    f"_sku_order::{_order_account_key}::{_items_fingerprint}"
                )

                # 사이트 룩앤필 스타일 (teal / light-green 팔레트)
                # ⚠ 떨림 방지 최종 대응:
                #   (a) 모든 transition/animation 완전 억제 — SortableJS 내부
                #       기본 transition 까지 덮어씀.
                #   (b) contain: strict → 자식 요소의 어떤 변경도 바깥 레이아웃에
                #       전파되지 않도록 CSS containment 로 격리.
                #   (c) item 에 고정 min-height + box-sizing:border-box →
                #       텍스트·폰트 렌더링 차이로 인한 픽셀 단위 변화 차단.
                #   (d) :hover 룰 전면 제거 — cursor 만 변경.
                # 색상은 항목의 텍스트 prefix 기반(.is-new / .is-manual). 위치 무관.
                # 클래스는 parent 페이지에서 JS 로 iframe 안 sortable-item 에 주입.
                _new_css = """
                .sortable-item.is-new,
                .sortable-item.is-new:hover,
                .sortable-item.is-new:focus {
                    background: linear-gradient(135deg, #f2fae3 0%, #e4f2cd 100%) !important;
                    background-color: transparent !important;
                    border-color: #cfe7a8 !important;
                    border-left-color: #7dbb26 !important;
                    color: #1b3d06 !important;
                }
                .sortable-item.is-manual,
                .sortable-item.is-manual:hover,
                .sortable-item.is-manual:focus {
                    background: linear-gradient(135deg, #f5eafa 0%, #ead7f4 100%) !important;
                    background-color: transparent !important;
                    border-color: #d4b8e3 !important;
                    border-left-color: #8e44ad !important;
                    color: #3d1854 !important;
                }
                """

                # 라이브러리 기본 CSS (.sortable-item, .sortable-item:hover)가
                #   background-color: var(--primary-color) (= Streamlit RED)
                #   color: #fff
                #   padding: 3px; margin: 5px; height: 100%
                # 를 설정해서, 우리 커스텀 CSS가 :hover 를 안 잡으면 그 속성이 그대로 노출됨.
                # => 모든 속성에 !important + :hover 룰도 동일하게 선언해 라이브러리 완전 무력화.
                _SORT_CSS = """
                .sortable-component,
                .sortable-component *,
                .sortable-item,
                .sortable-item * {
                    transition: none !important;
                    animation: none !important;
                }
                .sortable-component {
                    background: linear-gradient(135deg,#f8fcfd 0%,#edf6f8 100%) !important;
                    border: 2px dashed #00788a !important;
                    border-radius: 14px !important;
                    padding: 14px !important;
                    gap: 8px !important;
                    contain: layout style;
                }
                .sortable-item,
                .sortable-item:hover,
                .sortable-item:focus {
                    background: #ffffff !important;
                    background-color: #ffffff !important;
                    color: #1a3540 !important;
                    border: 1.5px solid #bdd8de !important;
                    border-left: 4px solid #00788a !important;
                    border-radius: 10px !important;
                    padding: 10px 14px !important;
                    margin: 0 !important;
                    height: auto !important;
                    font-weight: 600 !important;
                    font-size: 0.92rem !important;
                    cursor: grab !important;
                    user-select: none !important;
                    min-height: 42px !important;
                    box-sizing: border-box !important;
                    width: 100% !important;
                    contain: layout style paint;
                    display: flex !important;
                    align-items: center !important;
                    justify-content: flex-start !important;
                    text-align: left !important;
                }
                .sortable-item:active { cursor: grabbing !important; }
                """ + _new_css

                _reordered = sort_items(
                    _initial,
                    direction="vertical",
                    custom_style=_SORT_CSS,
                    key=_order_state_key,
                )

                # sortable iframe 안 각 .sortable-item 의 텍스트 prefix 를 보고
                # .is-new / .is-manual 클래스를 주입 → 색상이 위치가 아닌 항목 자체를 따라감.
                # Streamlit 컴포넌트 iframe 은 동일 오리진이라 parent → iframe 접근 가능.
                # MutationObserver 로 드래그 직후에도 자동 재적용.
                st.html(
                    """
                    <script>
                    (function(){
                      if (window.__sphSkuColorInit) return;
                      window.__sphSkuColorInit = true;
                      const NEW_TAG = '[신규 항목]';
                      const MAN_TAG = '[직접등록]';
                      function tagItems(doc){
                        const items = doc.querySelectorAll('.sortable-item');
                        items.forEach(it => {
                          const t = it.textContent || '';
                          it.classList.toggle('is-new', t.indexOf(NEW_TAG) !== -1);
                          it.classList.toggle('is-manual', t.indexOf(MAN_TAG) !== -1);
                        });
                      }
                      function attachToFrame(frame){
                        if (frame.__sphSkuTagged) return;
                        let doc;
                        try { doc = frame.contentDocument; } catch(e) { return; }
                        if (!doc || !doc.querySelector('.sortable-item')) return;
                        frame.__sphSkuTagged = true;
                        tagItems(doc);
                        const mo = new MutationObserver(() => tagItems(doc));
                        mo.observe(doc.body, {childList:true, subtree:true, characterData:true});
                      }
                      function scan(){
                        document.querySelectorAll('iframe').forEach(f => {
                          try {
                            if (f.contentDocument &&
                                f.contentDocument.querySelector('.sortable-component')) {
                              attachToFrame(f);
                            }
                          } catch(e) {}
                        });
                      }
                      scan();
                      // iframe 이 늦게 mount 되거나 rerun 으로 교체될 수 있으니 주기 재스캔
                      setInterval(scan, 700);
                    })();
                    </script>
                    """,
                    unsafe_allow_javascript=True,
                )

                # prefix 제거해 실제 SKU 순서 확정 — 신규/직접등록 양쪽 처리
                def _strip_prefix(_x: str) -> str:
                    if _x.startswith(_NEW_PREFIX):
                        return _x[len(_NEW_PREFIX):]
                    if _x.startswith(_MANUAL_PREFIX):
                        return _x[len(_MANUAL_PREFIX):]
                    return _x
                sku_order = [_strip_prefix(x) for x in _reordered]

                # ── 순서 자동 저장 ───────────────────────────────────────
                # 드래그/신규 항목 등으로 패널 순서가 바뀌면 즉시 saved_orders
                # 반영. 단, hidden 으로 패널에서 빠진 saved 항목은 보존(끝에
                # 유지)해야 사용자가 hidden 패널 X 클릭으로 복원 시 원래
                # 위치 흔적이 사라지지 않는다. 비교 기준은 "패널 가시 순서".
                _visible_saved = [
                    s for s in _saved_for_this if s not in _hidden_set_for_order
                ]
                _hidden_in_saved = [
                    s for s in _saved_for_this if s in _hidden_set_for_order
                ]
                if sku_order and sku_order != _visible_saved:
                    _new_saved = list(sku_order) + _hidden_in_saved
                    _save_order_for_account(_order_account_key, _new_saved)

                # 현재 순서 저장 버튼 (secondary)
                if st.button(
                    "💾  현재 순서 저장",
                    key=f"_save_order_btn_{_order_account_key}",
                    type="secondary",
                    use_container_width=True,
                    help=f"'{_order_account_key}' 계정의 현재 SKU 순서를 저장합니다.",
                ):
                    _save_order_for_account(_order_account_key, sku_order)
                    st.toast(
                        f"✅ '{_order_account_key}' SKU 순서 저장 완료",
                        icon="💾",
                    )
                    st.rerun()

                # ── 마스터 SKU 직접 등록 (현재 CSV 에 없어도 노출) ────────
                # 후보 소스 = master_data.csv ∪ Price List (둘 다 합쳐 누락 0).
                # 메모리상 master_data.csv 는 단일 진실 소스가 아니므로 Price
                # List 의 SKU 도 반드시 포함.
                _all_known_skus: set[str] = set()
                # (a) master_data.csv 의 sku_name
                try:
                    _all_known_skus |= {
                        str(n).strip()
                        for n in st.session_state.master_df["sku_name"].dropna().tolist()
                        if str(n).strip()
                    }
                except Exception:
                    pass
                # (b) Price List (xlsx) 의 SKU 명 — 마스터에 없는 SKU 포함
                if price_list_file is not None:
                    try:
                        from billing.loader import get_sku_tiers_from_price_list
                        _all_known_skus |= set(
                            get_sku_tiers_from_price_list(price_list_file).keys()
                        )
                    except Exception:
                        pass
                _all_master_skus = sorted(_all_known_skus)
                # 이미 CSV 에서 발견됐거나, 이미 직접등록된 항목은 후보에서 제외
                _manual_candidates = [
                    s for s in _all_master_skus
                    if s not in _found_skus and s not in _manual_skus_saved
                ]

                # ── 콜백: 멀티셀렉트 변경 시 즉시 저장 + 선택 해제 ──────
                _manual_add_key = f"_manual_add_ms::{_order_account_key}"

                def _on_manual_add(
                    acc: str = _order_account_key,
                    ms_key: str = _manual_add_key,
                ) -> None:
                    _sel = list(st.session_state.get(ms_key, []) or [])
                    if not _sel:
                        return
                    _curr = _load_manual_skus_map().get(acc, [])
                    _new  = list(dict.fromkeys(_curr + _sel))
                    _save_manual_skus_for_account(acc, _new)
                    st.session_state[ms_key] = []  # 선택 해제 — 패널에 chip 만 남도록
                    st.toast(f"✏️ 직접등록 {len(_sel)}개 추가", icon="➕")

                st.multiselect(
                    f"➕ 마스터에서 직접 SKU 추가 (총 {len(_all_master_skus)}개 후보 · CSV 사용량 0 도 노출)",
                    options=_manual_candidates,
                    default=[],
                    key=_manual_add_key,
                    on_change=_on_manual_add,
                    help=(
                        "마스터(master_data.csv) ∪ Price List 의 SKU 중 이번 CSV 에 "
                        "사용량이 없는 항목을 선택하면, 순서 조정 패널 가장 하단에 "
                        "[직접등록] 으로 추가됩니다. 인보이스에 빈 라인으로 출력됩니다."
                    ),
                )

                # ── 콜백: X 버튼 클릭 시 즉시 1개 제거 ────────────────────
                def _on_manual_remove(
                    acc: str, sku: str,
                ) -> None:
                    _curr = _load_manual_skus_map().get(acc, [])
                    _new  = [x for x in _curr if x != sku]
                    _save_manual_skus_for_account(acc, _new)
                    st.toast(f"🗑 '{sku}' 직접등록 해제", icon="✏️")

                # 직접등록된 SKU 개별 X 버튼 (현재 CSV 에 없는 것만 표시)
                if _manual_only:
                    st.caption("✏️ 직접등록된 SKU — X 클릭 시 즉시 제거")
                    for _ms in _manual_only:
                        _c1, _c2 = st.columns([6, 1])
                        with _c1:
                            st.markdown(
                                f"<div style='padding:6px 10px;background:#f5eafa;"
                                f"border-left:4px solid #8e44ad;border-radius:8px;"
                                f"font-size:0.88rem;color:#3d1854;font-weight:600;'>"
                                f"{_ms}</div>",
                                unsafe_allow_html=True,
                            )
                        with _c2:
                            st.button(
                                "✕",
                                key=f"_rm_manual::{_order_account_key}::{_ms}",
                                help=f"'{_ms}' 직접등록 해제",
                                use_container_width=True,
                                on_click=_on_manual_remove,
                                args=(_order_account_key, _ms),
                            )

                # ── 엑셀 미노출 SKU (수동) ──────────────────────────────
                # 엔진 계산(waterfall/sku_master/line_items) 에는 영향 주지 않음.
                # `generate_formatted_invoice` 호출 직전에 line_items / proj_results
                # 에서 sku_name 매칭 항목만 제거 → 출력물에서만 빠진다.
                # UI 패턴: 직접등록 SKU 와 동일 — multiselect(빈 default + on_change
                # 누적 저장) + 등록 항목별 chip + X 버튼.
                _saved_hidden_all = _load_hidden_skus_map()
                _saved_hidden_for_this = _lookup_account(
                    _saved_hidden_all, _order_account_key) or []
                # 후보 풀 = 이번 정산에 의미 있는 모든 SKU
                #         = CSV 발견 ∪ 직접등록 SKU (중복 제거, found 우선 순서)
                # sku_order 자체는 hidden 을 이미 제외했기 때문에 후보 계산에
                # 직접 쓰지 않는다 — 그러면 새 hidden 추가 불가가 된다.
                _hidden_pool = list(
                    dict.fromkeys(_found_skus + list(_manual_skus_saved))
                )
                # 후보 풀 안에 없는 저장값은 stale(다른 회사·다른 CSV) → 자동 정리.
                _hidden_for_this = [
                    s for s in _saved_hidden_for_this if s in _hidden_pool
                ]
                # 이후 단계(계산 키 / 출력 필터)에서 사용되는 hidden_skus 변수.
                hidden_skus = list(_hidden_for_this)

                # 후보 = 후보 풀 중 아직 hidden 으로 지정되지 않은 항목
                _hidden_candidates = [
                    s for s in _hidden_pool if s not in _hidden_for_this
                ]

                _hidden_add_key = f"_hidden_add_ms::{_order_account_key}"

                def _on_hidden_add(
                    acc: str = _order_account_key,
                    ms_key: str = _hidden_add_key,
                ) -> None:
                    _sel = list(st.session_state.get(ms_key, []) or [])
                    if not _sel:
                        return
                    _curr = _load_hidden_skus_map().get(acc, [])
                    _new  = list(dict.fromkeys(_curr + _sel))
                    _save_hidden_skus_for_account(acc, _new)
                    st.session_state[ms_key] = []  # 선택 해제 — 패널에 chip 만 남도록
                    st.toast(f"🚫 미노출 {len(_sel)}개 추가", icon="🚫")

                st.multiselect(
                    f"🚫 엑셀에서 제외할 SKU (총 {len(_hidden_pool)}개 중 선택)",
                    options=_hidden_candidates,
                    default=[],
                    key=_hidden_add_key,
                    on_change=_on_hidden_add,
                    help=(
                        "선택한 SKU 는 Invoice / Project 시트에 출력되지 않습니다.\n"
                        "엔진 계산(waterfall·무료 배분) 은 그대로 유지되며, **출력 직전**에만 제거됩니다.\n"
                        "완전 무료(subtotal $0) 항목을 숨기면 총액 변동 없음.\n"
                        "유료 항목을 숨기면 그만큼 청구 총액이 줄어드니 주의."
                    ),
                )

                def _on_hidden_remove(acc: str, sku: str) -> None:
                    _curr = _load_hidden_skus_map().get(acc, [])
                    _new  = [x for x in _curr if x != sku]
                    _save_hidden_skus_for_account(acc, _new)
                    st.toast(f"♻ '{sku}' 노출 복원", icon="♻")

                # [비용발생] 판정: 직전 정산 결과(_last_result) 의 line_items
                # 와 매칭해, hidden 안 SKU 중 final_krw > 0 인 항목은 무료
                # 한도를 넘어 실제 청구 대상이 된 것 → 엑셀에서 빠지면 총액
                # 불일치 위험. UI 에서 prefix 와 진한 색상으로 강조한다.
                # 회사 키가 다르면 정확하지 않으니 동일 회사일 때만 적용.
                _paid_hidden_set: set[str] = set()
                _last_res = st.session_state.get("_last_result") or {}
                if (
                    _last_res
                    and _hidden_for_this
                    and _last_res.get("company") == selected_company
                ):
                    for _it in _last_res.get("line_items") or []:
                        _nm = getattr(_it, "sku_name", "")
                        if _nm in _hidden_for_this:
                            try:
                                _krw = int(getattr(_it, "final_krw", 0) or 0)
                            except (TypeError, ValueError):
                                _krw = 0
                            if _krw > 0:
                                _paid_hidden_set.add(_nm)

                # 미노출 지정된 SKU 개별 X 버튼
                if _hidden_for_this:
                    if _paid_hidden_set:
                        st.caption(
                            "🚫 미노출 SKU — **[비용발생]** 항목은 무료 한도를 "
                            "초과해 엑셀 총액이 줄어듭니다 (X 클릭 시 노출 복원)"
                        )
                    else:
                        st.caption("🚫 미노출 SKU — X 클릭 시 즉시 노출 복원")
                    for _hs in _hidden_for_this:
                        _c1, _c2 = st.columns([6, 1])
                        with _c1:
                            if _hs in _paid_hidden_set:
                                # [비용발생]: 진한 붉은색
                                _label = f"<strong>[비용발생]</strong> {_hs}"
                                _bg = "#f5b7b1"
                                _bd = "#922b21"
                                _fc = "#641e16"
                                _bw = "5px"
                            else:
                                _label = _hs
                                _bg = "#fde8e8"
                                _bd = "#c0392b"
                                _fc = "#5b1a1a"
                                _bw = "4px"
                            st.markdown(
                                f"<div style='padding:6px 10px;background:{_bg};"
                                f"border-left:{_bw} solid {_bd};border-radius:8px;"
                                f"font-size:0.88rem;color:{_fc};font-weight:600;'>"
                                f"{_label}</div>",
                                unsafe_allow_html=True,
                            )
                        with _c2:
                            st.button(
                                "✕",
                                key=f"_rm_hidden::{_order_account_key}::{_hs}",
                                help=f"'{_hs}' 노출 복원",
                                use_container_width=True,
                                on_click=_on_hidden_remove,
                                args=(_order_account_key, _hs),
                            )

        # ═══ 우측: 과금 방식 / 통화·환율 / 다운로드 옵션(+정산 시작) ═══
        with col_right:
            # ── 과금 방식 영역 ──
            # 회사별로 "회사 통합 waterfall"(기본) vs "프로젝트별 독립
            # waterfall" 중 선택. 라디오 변경 시 자동 저장되어 다음 정산
            # 때 자동 로드된다.
            _billing_modes_all = _load_billing_modes()
            _saved_mode = _billing_modes_all.get(
                _order_account_key, BILLING_MODE_ACCOUNT
            )
            with st.container(border=True):
                st.markdown("#### 🧮 과금 방식")
                _mode_options = {
                    "회사 통합 (Google 실제 청구 방식)": BILLING_MODE_ACCOUNT,
                    "프로젝트별 독립 waterfall":          BILLING_MODE_PER_PROJECT,
                }
                _mode_labels = list(_mode_options.keys())
                _saved_idx = next(
                    (i for i, lb in enumerate(_mode_labels)
                     if _mode_options[lb] == _saved_mode),
                    0,
                )
                _mode_label = st.radio(
                    "계정별 과금 방식 선택",
                    options=_mode_labels,
                    index=_saved_idx,
                    key=f"_billing_mode_radio::{_order_account_key}",
                    help=(
                        "● 회사 통합: 결제계정 전체 usage 로 tier waterfall "
                        "(Google 실제 청구 방식과 동일).\n"
                        "● 프로젝트별 독립: 각 프로젝트가 자기 usage 만으로 "
                        "waterfall — 소규모 프로젝트도 tier1 부터 시작하므로 "
                        "할인 혜택이 적게 적용되어 총액이 소폭 높아진다."
                    ),
                )
                billing_mode = _mode_options[_mode_label]
                # 선택이 변경되면 자동 저장
                if billing_mode != _saved_mode:
                    _save_billing_mode_for_account(_order_account_key, billing_mode)
                    st.toast(
                        f"💾 '{_order_account_key}' 과금 방식 저장: {_mode_label}",
                        icon="🧮",
                    )

                # Project(요약) 시트 포함 여부 — 계정별로 저장되어 다음 정산 시 자동 로드.
                _saved_include_proj = _load_include_project_flags().get(
                    _order_account_key, True
                )
                include_project_sheet = st.checkbox(
                    "📑 Project(요약) 시트 포함",
                    value=_saved_include_proj,
                    key=f"_include_proj_chk::{_order_account_key}",
                    help="계정별로 저장됩니다. 해제 시 엑셀에 Project 시트가 생성되지 않습니다.",
                )
                if include_project_sheet != _saved_include_proj:
                    _save_include_project_for_account(
                        _order_account_key, include_project_sheet
                    )
                    st.toast(
                        f"💾 '{_order_account_key}' Project 시트 "
                        f"{'포함' if include_project_sheet else '제외'} 저장",
                        icon="📑",
                    )

                # 반올림 자리수 — =ROUND(SUM(I..:I..),N) 의 N 변경 +
                # 구간별(tier) **금액(amount, I 열)** 의 표시 포맷도 동일
                # 자리수로 맞춤. **단가(H 열) 는 영향 없음** (통화 기본 유지).
                # 계정별로 저장, 미저장 시 통화 기본(KRW=0, USD=2).
                # 표시 포맷만 변경 — 셀 수식/값은 그대로라 결과값 변동 없음.
                _default_round = 0 if currency == "KRW" else 2
                _saved_round = _load_subtotal_round_map().get(
                    _order_account_key, _default_round
                )
                if _saved_round not in (0, 2):
                    _saved_round = _default_round
                _round_options = {
                    "정수 (,0)":      0,
                    "소수 2자리 (,2)": 2,
                }
                _round_labels = list(_round_options.keys())
                _round_idx = next(
                    (i for i, lb in enumerate(_round_labels)
                     if _round_options[lb] == _saved_round),
                    0 if _saved_round == 0 else 1,
                )
                _round_label = st.radio(
                    "반올림 자리수",
                    options=_round_labels,
                    index=_round_idx,
                    horizontal=True,
                    key=f"_subtotal_round::{_order_account_key}",
                    help=(
                        "Invoice 시트의 **금액(amount, I 열)** 표시 자리수 — "
                        "소계와 구간별(tier) 금액에 동일 적용. 단가(H 열) 는 "
                        "통화 기본 자리수가 유지됩니다(영향 없음). 계정별로 "
                        "저장되며 셀 수식/값은 그대로라 결과값 변동 없습니다."
                    ),
                )
                subtotal_round = _round_options[_round_label]
                if subtotal_round != _saved_round:
                    _save_subtotal_round_for_account(
                        _order_account_key, subtotal_round
                    )
                    st.toast(
                        f"💾 '{_order_account_key}' 반올림 자리수: "
                        f"{subtotal_round}",
                        icon="🔢",
                    )

            # ── 최소사용비용 영역 ──
            # Google Maps Platform 기본 월 ₩500,000 규정이지만, 회사별 계약에
            # 따라 달라질 수 있음 (금액·통화). 0 으로 두면 적용 안 함.
            _saved_min_amt, _saved_min_cur = _min_charge_for_account(_order_account_key)
            with st.container(border=True):
                st.markdown("#### 💵 최소사용비용")
                _mc_col1, _mc_col2 = st.columns([2, 1])
                with _mc_col2:
                    _mc_cur_label = st.radio(
                        "통화",
                        options=["원 (₩)", "달러 ($)"],
                        index=(0 if _saved_min_cur == "KRW" else 1),
                        horizontal=False,
                        key=f"_min_charge_cur::{_order_account_key}",
                    )
                    min_charge_currency = "KRW" if _mc_cur_label.startswith("원") else "USD"
                with _mc_col1:
                    # KRW 일 때 정수 입력, USD 일 때 소수점 둘째자리 허용
                    _mc_step = 1000.0 if min_charge_currency == "KRW" else 0.01
                    _mc_fmt  = "%.0f"  if min_charge_currency == "KRW" else "%.2f"
                    min_charge_amount = st.number_input(
                        f"최소사용비용 금액 ({'₩' if min_charge_currency == 'KRW' else '$'})",
                        min_value=0.0,
                        value=float(_saved_min_amt),
                        step=_mc_step,
                        format=_mc_fmt,
                        key=f"_min_charge_amt::{_order_account_key}",
                        help="0 으로 설정하면 최소사용비용 룰을 적용하지 않습니다.",
                    )
                # 변경 시 자동 저장
                _mc_changed = (
                    float(min_charge_amount) != float(_saved_min_amt)
                    or min_charge_currency != _saved_min_cur
                )
                if _mc_changed:
                    _save_min_charge_for_account(
                        _order_account_key, float(min_charge_amount),
                        min_charge_currency,
                    )
                    _mc_display = (
                        f"₩{int(min_charge_amount):,}" if min_charge_currency == "KRW"
                        else f"${float(min_charge_amount):,.2f}"
                    )
                    st.toast(
                        f"💾 '{_order_account_key}' 최소사용비용 저장: {_mc_display}",
                        icon="💵",
                    )

            # ── 통화 / 환율 영역 ──
            with st.container(border=True):
                st.markdown("#### 💰 통화 · 환율")
                _detected = st.session_state.get("_detected_currency", "USD")
                _cur_options = ["달러 ($)", "원화 (₩)"]
                _cur_idx     = 0 if _detected == "USD" else 1
                currency_label = st.radio(
                    "단가표 통화",
                    options=_cur_options,
                    index=_cur_idx,
                    horizontal=True,
                    key="_currency_radio",
                    help="단가표가 USD 기준이면 '달러', KRW 기준이면 '원화'를 선택하세요. "
                         "단가표 업로드 시 자동 감지됩니다.",
                )
                currency = "USD" if currency_label.startswith("달러") else "KRW"

                # currency 확정 후 price_list_file 최종 선택.
                #   - USD 선택 → 달러 단가표 (업로드 우선 > 저장파일 > 레거시)
                #   - KRW 선택 → 원화 단가표 (업로드 우선 > 저장파일)
                # 해당 통화 단가표가 없으면 다른 통화 파일로 fallback 하지 않는다
                # (잘못된 단위로 단가가 찍히는 것을 방지 — 경고 표시).
                if currency == "USD":
                    if _uploaded_price_usd is not None:
                        _uploaded_price_usd.seek(0)
                        price_list_file = _uploaded_price_usd
                    elif PRICE_LIST_SAVED_USD.exists():
                        price_list_file = PRICE_LIST_SAVED_USD
                    elif PRICE_LIST_SAVED.exists():
                        # 레거시 파일이 아직 남아있고 USD 로 감지되면 사용.
                        try:
                            _legacy_is_usd = detect_price_list_currency(PRICE_LIST_SAVED) == "USD"
                        except Exception:
                            _legacy_is_usd = True
                        price_list_file = PRICE_LIST_SAVED if _legacy_is_usd else None
                    else:
                        price_list_file = None
                else:
                    if _uploaded_price_krw is not None:
                        _uploaded_price_krw.seek(0)
                        price_list_file = _uploaded_price_krw
                    elif PRICE_LIST_SAVED_KRW.exists():
                        price_list_file = PRICE_LIST_SAVED_KRW
                    else:
                        price_list_file = None
                if price_list_file is None:
                    st.warning(
                        f"⚠️ 선택한 통화({'달러($)' if currency == 'USD' else '원화(₩)'}) 에 "
                        "해당하는 단가표가 업로드되어 있지 않습니다. 상단에서 해당 통화 단가표를 첨부해 주세요."
                    )

                # 환율: 달러 모드에서만 입력 가능 (원화 모드에선 사용 안 함)
                # disabled=True 일 때만 회색 배경 — 달러 모드 활성 상태는 전역
                # CSS 로 흰 배경 처리(인풋: not(:disabled)).
                # 세션 진입 시 빈 값으로 시작 (이전 입력값 자동 복원 안 함).
                _rate_raw = st.text_input(
                    "환율 (USD → KRW)",
                    value="",
                    max_chars=7,
                    placeholder="예: 1427.87" if currency == "USD" else "원화 모드 — 입력 불필요",
                    key="_rate_raw_input",
                    disabled=(currency == "KRW"),
                )
                # 정산 시작 시 환율 미입력 에러 메시지가 채워질 자리.
                _rate_error_ph = st.empty()

                # 사용자가 입력하면 invalid flash 자동 해제.
                if _rate_raw and st.session_state.get("_rate_invalid_flash"):
                    st.session_state.pop("_rate_invalid_flash", None)

                # 이전 rerun 에서 flash 가 켜진 상태면 메시지 유지(타이핑 전까지).
                if currency == "USD" and st.session_state.get("_rate_invalid_flash"):
                    _rate_error_ph.markdown(
                        '<div style="color:#ef4444; font-size:0.85rem; '
                        'margin-top:-10px; padding-left:4px;">'
                        '환율을 입력해 주세요</div>',
                        unsafe_allow_html=True,
                    )

                if currency == "KRW":
                    exchange_rate = 1.0   # placeholder — 원화 모드에선 실제 미사용
                else:
                    try:
                        exchange_rate = float(_rate_raw) if _rate_raw else 0.0
                    except ValueError:
                        exchange_rate = 0.0

            # ── 환율 표기 설정 (은행·문구·날짜) ──
            # Invoice 하단 "환율(하나은행 2026.02.27 최종 송금환율 기준)" 줄의
            # 세 요소를 계정별로 자유 조정 + 저장. 변경 시 자동 저장.
            _rl = _rate_label_for_account(_order_account_key)
            # 기본 날짜 = billing_month 마지막 날 (미저장 시)
            import calendar as _cal
            try:
                _bm_year  = int((billing_month or "2026-01")[:4])
                _bm_month = int((billing_month or "2026-01")[5:7])
                _bm_last  = _cal.monthrange(_bm_year, _bm_month)[1]
                _default_rate_date = date(_bm_year, _bm_month, _bm_last)
            except Exception:
                _default_rate_date = date.today()
            if _rl.get("date"):
                try:
                    _y, _m, _d = [int(x) for x in str(_rl["date"]).split("-")]
                    _saved_rate_date = date(_y, _m, _d)
                except Exception:
                    _saved_rate_date = _default_rate_date
            else:
                _saved_rate_date = _default_rate_date

            with st.container(border=True):
                st.markdown("#### 📝 환율 표기")
                # KRW(원화) 단가표 정산은 환율 변환이 없어 환율 표기 자체가
                # 인보이스에 들어가지 않는다 → 입력 영역 전체 비활성화.
                _rate_disabled = (currency == "KRW")
                if _rate_disabled:
                    st.caption(
                        "ℹ️ 원화(KRW) 단가표 정산에는 환율 표기가 사용되지 않습니다. "
                        "입력이 비활성화됩니다."
                    )

                # 은행 — 셀렉트 + 직접입력 (한 글자 타이핑 시 자동 매칭)
                _bank_options = MAJOR_BANKS + ["직접입력"]
                _saved_bank = _rl["bank"]
                _bank_idx = (
                    _bank_options.index(_saved_bank)
                    if _saved_bank in _bank_options else
                    _bank_options.index("직접입력")
                )
                _bank_sel = st.selectbox(
                    "은행",
                    options=_bank_options,
                    index=_bank_idx,
                    key=f"_bank_sel::{_order_account_key}",
                    disabled=_rate_disabled,
                )
                if _bank_sel == "직접입력":
                    _bank_typed = st.text_input(
                        "은행명 직접입력",
                        value=(_saved_bank if _saved_bank not in MAJOR_BANKS else ""),
                        key=f"_bank_typed::{_order_account_key}",
                        help="한 글자만 입력해도 MAJOR_BANKS 에서 자동 매칭 ('하' → 하나은행)",
                        disabled=_rate_disabled,
                    )
                    _auto = _match_bank_prefix(_bank_typed)
                    if _auto and _auto != _bank_typed and len(_bank_typed.strip()) <= 2:
                        st.caption(f"💡 자동완성 제안: **{_auto}** (유지하려면 그대로 두세요)")
                        bank_name = _auto
                    else:
                        bank_name = _bank_typed.strip() or DEFAULT_BANK_NAME
                else:
                    bank_name = _bank_sel

                # 고정문구 — 셀렉트 + 직접입력 ("추가문구" 영역은 제거됨)
                _phrase_options = RATE_PHRASES + ["직접입력"]
                _saved_phrase = _rl["phrase"]
                _phrase_idx = (
                    _phrase_options.index(_saved_phrase)
                    if _saved_phrase in _phrase_options else
                    _phrase_options.index("직접입력")
                )
                _phrase_sel = st.selectbox(
                    "고정문구",
                    options=_phrase_options,
                    index=_phrase_idx,
                    key=f"_phrase_sel::{_order_account_key}",
                    disabled=_rate_disabled,
                )
                if _phrase_sel == "직접입력":
                    rate_phrase_text = st.text_input(
                        "고정문구 직접입력",
                        value=(_saved_phrase if _saved_phrase not in RATE_PHRASES else ""),
                        key=f"_phrase_typed::{_order_account_key}",
                        disabled=_rate_disabled,
                    ).strip() or DEFAULT_RATE_PHRASE
                else:
                    rate_phrase_text = _phrase_sel

                # 추가문구 영역은 제거 — 내부적으로는 빈 값 고정.
                rate_extra_text = ""

                # 날짜 — 달력 + 직접입력 (st.date_input 은 둘 다 지원)
                _picked_date = st.date_input(
                    "날짜",
                    value=_saved_rate_date,
                    key=f"_rate_date::{_order_account_key}",
                    format="YYYY-MM-DD",
                    disabled=_rate_disabled,
                )
                rate_date_str = _picked_date.strftime("%Y.%m.%d")
                _rate_date_iso = _picked_date.strftime("%Y-%m-%d")

                # 실시간 미리보기
                _preview_phrase = rate_phrase_text + (
                    f" {rate_extra_text.strip()}" if rate_extra_text.strip() else ""
                )
                st.caption(
                    f"🔎 미리보기: **환율({bank_name} {rate_date_str} {_preview_phrase})**"
                )

                # 변경 감지 → 자동 저장
                if (bank_name != _rl["bank"]
                        or rate_phrase_text != _rl["phrase"]
                        or rate_extra_text.strip() != _rl["extra"]
                        or _rate_date_iso != (_rl["date"] or _default_rate_date.strftime("%Y-%m-%d"))):
                    _save_rate_label_for_account(
                        _order_account_key, bank_name, rate_phrase_text,
                        rate_extra_text.strip(), _rate_date_iso,
                    )
                    st.toast(
                        f"💾 '{_order_account_key}' 환율 표기 저장",
                        icon="📝",
                    )

            # 설정 변경 감지 → 이전 결과 자동 무효화
            _current_calc_key = (
                f"{billing_month}|{exchange_rate}|{bank_name}|{currency}|"
                f"{billing_mode}|{min_charge_amount}|{min_charge_currency}|"
                f"{rate_phrase_text}|{rate_extra_text}|{rate_date_str}|"
                f"{include_project_sheet}|{subtotal_round}|"
                f"{','.join(sorted(hidden_skus or []))}"
            )
            if st.session_state.get("_calc_key") != _current_calc_key:
                st.session_state["_calc_key"] = _current_calc_key
                st.session_state.pop("_last_result", None)
                # 확인 대기 중인 결과도 설정 변경 시 stale → 폐기
                st.session_state.pop("_pending_result", None)
                st.session_state.pop("_pending_auto_dl_key", None)
                st.session_state.pop("_pending_hidden_paid", None)

            # ── 다운로드 옵션 영역 (정산 시작 버튼 포함) ──
            with st.container(border=True):
                st.markdown("#### ⚙️ 다운로드 옵션")
                dl_excel = st.checkbox(
                    "📗 엑셀 다운받기 (.xlsx)",
                    value=True,
                    key="_dl_excel_chk",
                    help="Invoice / Project / GMP Price List 3개 시트로 구성된 엑셀 파일",
                )
                _pdf_ok = _pdf_export_available()
                _pdf_label = (
                    "📄 PDF 다운받기 (Invoice 시트)"
                    if _pdf_ok else
                    "📄 PDF 다운받기 (사용 불가 — Excel/pywin32 확인 필요)"
                )
                dl_pdf = st.checkbox(
                    _pdf_label,
                    value=False,
                    key="_dl_pdf_chk",
                    disabled=not _pdf_ok,
                    help="엑셀의 Invoice 시트를 PDF로 변환해 동일한 레이아웃으로 출력합니다. "
                         "(Microsoft Excel 이 설치된 Windows 환경에서만 동작)",
                )

                # 정산 시작 버튼은 다운로드 옵션 영역 하단에 배치
                # (검증은 클릭 시점에 얼럿/토스트로 처리)
                run_button = st.button(
                    "▶  정산 시작",
                    type="primary",
                    use_container_width=True,
                )

        # 클릭 시점 검증 — 미충족 시 얼럿 후 run_button 을 False 로 재설정
        if run_button:
            _missing_msgs: list[str] = []
            if not billing_month:
                _missing_msgs.append("CSV 파일을 먼저 업로드해주세요.")
            if not (dl_excel or dl_pdf):
                _missing_msgs.append("다운로드 형식(엑셀 / PDF)을 하나 이상 선택해주세요.")

            # 환율 미입력은 인풋 바로 아래 빨간 메시지 + 포커싱으로 처리.
            _rate_missing = (
                currency == "USD" and (not exchange_rate or exchange_rate <= 0)
            )
            if _rate_missing:
                st.session_state["_rate_invalid_flash"] = True
                _rate_error_ph.markdown(
                    '<div style="color:#ef4444; font-size:0.85rem; '
                    'margin-top:-10px; padding-left:4px;">'
                    '환율을 입력해 주세요</div>',
                    unsafe_allow_html=True,
                )
                st.html(
                    """
                    <script>
                    (function(){
                      try {
                        var inp = window.parent.document.querySelector(
                          '.st-key-_rate_raw_input input'
                        );
                        if (inp && !inp.disabled) {
                          inp.focus();
                          inp.scrollIntoView({behavior:'smooth', block:'center'});
                        }
                      } catch(e) {}
                    })();
                    </script>
                    """,
                    unsafe_allow_javascript=True,
                )

            if _missing_msgs or _rate_missing:
                for _m in _missing_msgs:
                    st.toast(f"⚠ {_m}", icon="⚠️")
                if _missing_msgs:
                    st.error(" / ".join(_missing_msgs))
                run_button = False   # 아래 정산 블록 실행 차단

        # ── 정산 실행 ─────────────────────────────────────────────────────────
        # 컨펌 다이얼로그에서 "노출로 옮기고 재정산" 버튼을 누른 경우
        # _rerun_billing 플래그가 켜져 있어 사용자 클릭 없이 한 번 더 자동 실행.
        if st.session_state.pop("_rerun_billing", False):
            run_button = True
        if run_button:
            if price_list_file is None:
                st.error(
                    "Price List(xlsx) 가 없습니다. 우측 상단에서 업로드하거나 "
                    "billing/saved_price_list.xlsx 를 준비해 주세요. "
                    "SKU 정의(단가·무료한도)의 단일 소스입니다."
                )
            else:
                loading_ph = st.empty()
                _render_loading(loading_ph, 0, "⚙️ 파일 전처리 중...")
                try:
                    raw_rows = _cached_preprocess(
                        str(tmp_input_path), billing_month, selected_company
                    )
                    _render_loading(loading_ph, 20, "📦 SKU 마스터 구성 중 (CSV + Price List)...")

                    usage_rows = load_usage_rows(raw_rows)
                    # 단일 진실 소스:
                    #   - "어떤 SKU 가 사용됐는가" → 업로드된 CSV (usage_rows)
                    #   - "각 SKU 의 단가/무료한도"  → Price List (xlsx)
                    # master_data.csv 는 정산 경로에서 더 이상 참조하지 않는다
                    # (SKU 관리 탭의 수동 편집 전용으로만 남음).
                    sku_master = build_sku_master_from_usage(usage_rows, price_list_file)

                    # CSV 에 있지만 Price List 에 없어 매칭 실패한 SKU 탐지
                    # → 정산완료 배너 아래에서 노출하도록 결과에 함께 저장
                    # (사용자 가시성: 처리 중 한 자리 차지하지 말고 완료 후 노출).
                    _missing_skus = detect_missing_skus(usage_rows, sku_master)

                    _render_loading(loading_ph, 35, "🧮 Waterfall 과금 계산 중...")
                    _ex = Decimal(str(exchange_rate))
                    _mr = Decimal(str(margin_rate))

                    # per_project 모드: 계정 Free Usage Cap 을 **usage 큰 프로젝트
                    # 부터 순차 소진(rollover)** 해 각 프로젝트의 무료 배정량을 결정.
                    #   Case A=13K B=15K cap=10K  →  B 10K, A 0  → A billable 13K, B 5K
                    #   Case A=6K  B=3K  cap=10K  →  A 6K,  B 3K → 둘 다 billable 0
                    # 무료량은 프로젝트마다 다를 수 있어 Free Usage 셀은 계산된 값을
                    # 직접 기록(SUMIF 수식 공유 불가능 — 값이 프로젝트별 상이).
                    _per_proj_invoices = None
                    if billing_mode == "per_project":
                        from collections import defaultdict as _dd
                        _proj_rows_map = _dd(list)
                        _proj_name_map: dict[str, str] = {}
                        _proj_sku_usage: dict[str, dict[str, int]] = _dd(
                            lambda: _dd(int)
                        )
                        _sid_to_name: dict[str, str] = {}
                        for _r in usage_rows:
                            _proj_rows_map[_r.project_id].append(_r)
                            _proj_name_map.setdefault(
                                _r.project_id,
                                getattr(_r, "project_name", None) or _r.project_id,
                            )
                            _proj_sku_usage[_r.project_id][_r.sku_id] += int(
                                _r.usage_amount or 0
                            )
                            _nm = getattr(_r, "sku_name", None)
                            if _nm and _r.sku_id not in _sid_to_name:
                                _sid_to_name[_r.sku_id] = str(_nm).strip()

                        # free cap 소스 = Price List (master 에 없는 SKU 도 포함).
                        _price_caps: dict[str, int] = {}
                        if price_list_file is not None:
                            try:
                                _price_caps = get_free_caps_from_price_list(price_list_file)
                            except Exception:
                                _price_caps = {}

                        def _full_cap_for_sid(_sid: str) -> int:
                            _nm = _sid_to_name.get(_sid, "")
                            if _nm and _nm in _price_caps:
                                return int(_price_caps[_nm])
                            _sku = sku_master.get(_sid)
                            return int(getattr(_sku, "free_usage_cap", 0) or 0)

                        # SKU 별로 usage 큰 프로젝트부터 rollover 소진.
                        _all_sku_ids: set[str] = set()
                        for _sku_map in _proj_sku_usage.values():
                            _all_sku_ids.update(_sku_map.keys())

                        _proj_sku_free_cap: dict[str, dict[str, int]] = _dd(dict)
                        for _sid in _all_sku_ids:
                            _full_cap = _full_cap_for_sid(_sid)
                            if _full_cap <= 0:
                                continue
                            # (usage desc, proj_id asc) 로 결정적 정렬.
                            _rank = sorted(
                                [(pid, _proj_sku_usage[pid].get(_sid, 0))
                                 for pid in _proj_rows_map.keys()
                                 if _proj_sku_usage[pid].get(_sid, 0) > 0],
                                key=lambda x: (-x[1], x[0]),
                            )
                            _remaining = _full_cap
                            for _pid, _pu in _rank:
                                _take = min(_remaining, _pu)
                                _proj_sku_free_cap[_pid][_sid] = int(_take)
                                _remaining -= _take

                        _per_proj_invoices = []
                        for _pid in sorted(_proj_rows_map.keys()):
                            _items = calculate_billing(
                                _proj_rows_map[_pid], sku_master, _ex, _mr,
                                mode="account",
                                free_cap_override=_proj_sku_free_cap.get(_pid),
                            )
                            _per_proj_invoices.append({
                                "proj_name":  _proj_name_map[_pid],
                                "line_items": _items,
                            })

                    line_items   = calculate_billing(
                        usage_rows, sku_master, _ex, _mr, mode=billing_mode
                    )
                    proj_results = calculate_billing_by_project(
                        usage_rows, sku_master, _ex, _mr, mode=billing_mode,
                        proj_sku_free_cap=(
                            dict(_proj_sku_free_cap)
                            if billing_mode == "per_project" else None
                        ),
                    )

                    # ── 직접등록 SKU 빈 라인 주입 (출력 단계 한정) ──────────
                    # 사용자가 사이드 패널에서 마스터로부터 직접 추가한 SKU 중
                    # 이번 CSV 사용량이 0 인 항목은 엔진 결과에 없다 →
                    # 인보이스에 빈 라인으로 라도 노출하기 위해 BillingLineItem
                    # stub 을 만들어 line_items / per_project_invoices 양쪽에 주입.
                    # 추가로 generate_formatted_invoice 의 canonical 필터가
                    # usage=0 항목을 제거하므로 `force_keep_skus` 로 화이트리스트.
                    try:
                        _manual_for_inject = _lookup_account(
                            _load_manual_skus_map(), _order_account_key
                        ) or []
                    except Exception:
                        _manual_for_inject = []
                    _manual_keep_set: set[str] = set()
                    if _manual_for_inject:
                        from billing.models import BillingLineItem as _BLI
                        from decimal import Decimal as _D

                        def _make_stub(_nm: str) -> "_BLI":
                            return _BLI(
                                billing_month   = billing_month or "",
                                project_id      = "",
                                project_name    = "",
                                sku_id          = "",
                                sku_name        = _nm,
                                total_usage     = 0,
                                free_usage_cap  = 0,
                                free_cap_applied= 0,
                                billable_usage  = 0,
                                tier_breakdown  = [],
                                subtotal_usd    = _D("0"),
                                exchange_rate   = _ex,
                                margin_rate     = _mr,
                                final_krw       = _D("0"),
                            )

                        # account 모드 line_items 주입
                        _existing_names = {
                            getattr(_it, "sku_name", "") for _it in line_items
                        }
                        _missing_manual = [
                            m for m in _manual_for_inject
                            if m and m not in _existing_names
                        ]
                        for _nm in _missing_manual:
                            line_items.append(_make_stub(_nm))
                        _manual_keep_set |= set(_manual_for_inject)

                        # per_project 모드: 각 프로젝트 line_items 에도 주입
                        # (해당 프로젝트에 없는 manual SKU 만)
                        if _per_proj_invoices:
                            for _entry in _per_proj_invoices:
                                _proj_items = _entry.get("line_items") or []
                                _proj_names = {
                                    getattr(_it, "sku_name", "") for _it in _proj_items
                                }
                                for _nm in _manual_for_inject:
                                    if _nm and _nm not in _proj_names:
                                        _proj_items.append(_make_stub(_nm))
                                _entry["line_items"] = _proj_items

                    # ── 수동 미노출 SKU 필터 (출력 단계 한정) ──────────────
                    # 엔진 결과(line_items / proj_results / _per_proj_invoices) 는
                    # 그대로 두고, 엑셀 생성 함수에 넘겨줄 **사본**에서 지정된
                    # sku_name 만 제거한다. waterfall 계산·무료 배분은 영향 없음.
                    _hidden_set = set(hidden_skus or [])
                    if _hidden_set:
                        _hidden_impact_krw = sum(
                            int(getattr(_it, "final_krw", 0) or 0)
                            for _it in line_items
                            if getattr(_it, "sku_name", "") in _hidden_set
                        )
                        if _hidden_impact_krw > 0:
                            st.warning(
                                f"⚠️ 미노출 지정한 SKU 중 **유료 항목**이 포함돼 있어 "
                                f"엑셀 총액이 약 **₩{_hidden_impact_krw:,}** 감소합니다. "
                                "완전 무료 SKU 만 제외하려면 해당 항목을 선택에서 빼세요."
                            )

                        _line_items_out = [
                            _it for _it in line_items
                            if getattr(_it, "sku_name", "") not in _hidden_set
                        ]
                        _proj_results_out = []
                        for _pr in (proj_results or []):
                            _skus_filtered = {
                                _nm: _v for _nm, _v in (_pr.get("skus") or {}).items()
                                if _nm not in _hidden_set
                            }
                            if not _skus_filtered:
                                continue
                            _new_pr = dict(_pr)
                            _new_pr["skus"] = _skus_filtered
                            _new_pr["total_usd"] = sum(
                                (_v.get("subtotal_usd") or 0)
                                for _v in _skus_filtered.values()
                            )
                            _new_pr["total_krw"] = sum(
                                (_v.get("final_krw") or 0)
                                for _v in _skus_filtered.values()
                            )
                            _proj_results_out.append(_new_pr)
                        _per_proj_invoices_out = None
                        if _per_proj_invoices is not None:
                            _per_proj_invoices_out = []
                            for _entry in _per_proj_invoices:
                                _items_f = [
                                    _it for _it in (_entry.get("line_items") or [])
                                    if getattr(_it, "sku_name", "") not in _hidden_set
                                ]
                                _per_proj_invoices_out.append({
                                    "proj_name":  _entry.get("proj_name"),
                                    "line_items": _items_f,
                                })
                    else:
                        _line_items_out        = line_items
                        _proj_results_out      = proj_results
                        _per_proj_invoices_out = _per_proj_invoices

                    _render_loading(loading_ph, 55, "📄 Excel 인보이스 생성 중...")
                    _safe        = (selected_company or "전체").replace("/", "_").replace("\\", "_")
                    # 임시 — 파일명 앞에 's' prefix 부여 (테스트 산출물 구분용)
                    _fname_xlsx  = f"sGMP_Invoice_{_safe}.xlsx"
                    _fname_pdf   = f"sGMP_Invoice_{_safe}.pdf"
                    # sku_order 에서도 미노출 항목 제거 — Invoice 시트 순서
                    # 렌더 시 빈 섹션이 끼지 않도록 깔끔하게 정리.
                    _sku_order_out = [
                        _n for _n in (sku_order or [])
                        if _n not in (_hidden_set if hidden_skus else set())
                    ]

                    _excel_bytes = generate_formatted_invoice(
                        line_items           = _line_items_out,
                        company_name         = selected_company or "전체",
                        billing_month        = billing_month,
                        exchange_rate        = _ex,
                        margin_rate          = _mr,
                        bank_name            = bank_name,
                        proj_results         = _proj_results_out,
                        price_list_file      = price_list_file,
                        sku_order            = _sku_order_out or None,
                        currency             = currency,
                        billable_skus        = _billable_skus,
                        billing_mode         = billing_mode,
                        per_project_invoices = _per_proj_invoices_out,
                        min_charge_amount    = float(min_charge_amount),
                        min_charge_currency  = min_charge_currency,
                        rate_date_str        = rate_date_str,
                        rate_phrase          = rate_phrase_text,
                        rate_extra           = rate_extra_text.strip(),
                        include_project_sheet= include_project_sheet,
                        subtotal_round       = subtotal_round,
                        force_keep_skus      = _manual_keep_set or None,
                    )

                    # 엑셀 자체 정합성 검사 (단일 정산용) — 결과는 _result_dict
                    # 에 동봉하고 결과 영역에서 빨간 박스로 노출.
                    _val_warns_single: list[str] = []
                    if _excel_bytes:
                        try:
                            _val_warns_single = validate_invoice_excel(
                                _excel_bytes,
                                line_items=_line_items_out,
                                company_name=selected_company or "전체",
                            )
                        except Exception as _ve:
                            _val_warns_single = [f"검사 함수 오류: {type(_ve).__name__}: {_ve}"]
                        if _val_warns_single:
                            print(f"[정산] {selected_company} ⚠ 정합성 경고 {len(_val_warns_single)}건")

                    # PDF 변환 (체크된 경우만)
                    _pdf_bytes = None
                    _pdf_error = None
                    if dl_pdf:
                        _render_loading(loading_ph, 75, "📄 PDF 변환 중 (Excel 실행)...")
                        from pdf_export import xlsx_sheet_to_pdf
                        # per_project 모드: 시트명이 프로젝트명이라 "Invoice"가 없음.
                        # 첫 프로젝트 시트를 PDF 로 변환(단일 시트 PDF 제약 때문).
                        if billing_mode == "per_project" and _per_proj_invoices:
                            from invoice_generator import _safe_sheet_title
                            _pdf_sheet = _safe_sheet_title(
                                _per_proj_invoices[0]["proj_name"], used=[]
                            )
                        else:
                            _pdf_sheet = "Invoice"
                        _pdf_bytes, _pdf_error = xlsx_sheet_to_pdf(
                            _excel_bytes, _pdf_sheet
                        )

                    _render_loading(loading_ph, 100, "✅ 완료!")
                    loading_ph.empty()

                    _result_dict = {
                        "line_items":      line_items,
                        "proj_results":    proj_results,
                        "company":         selected_company,
                        "billing_month":   billing_month,
                        "exchange_rate":   _ex,
                        "margin_rate":     _mr,
                        "bank_name":       bank_name,
                        "excel_bytes":     _excel_bytes if dl_excel else None,
                        "excel_filename":  _fname_xlsx,
                        "pdf_bytes":       _pdf_bytes,
                        "pdf_filename":    _fname_pdf,
                        "pdf_error":       _pdf_error,
                        "missing_skus":    _missing_skus,
                        "validation_warnings": _val_warns_single,
                    }
                    # 자동 다운로드용 키 — 같은 결과를 재 다운로드하지 않도록
                    _auto_dl_key_value = (
                        f"{selected_company}|{billing_month}|{len(_excel_bytes)}|"
                        f"{'X' if dl_excel else '-'}{'P' if dl_pdf else '-'}"
                    )

                    # ── 미노출 SKU 비용 확인 게이트 ─────────────────────────
                    # hidden_skus 안에 무료 한도 초과 SKU(final_krw > 0)가 있으면
                    # 엑셀 총액이 실제 청구액과 어긋난다. 다운로드 직전에 사용자
                    # 명시 확인을 받는다. 계산 결과는 _pending_result 에 임시
                    # 캐시되고, 사용자가 확인하면 _last_result 로 promote 된다.
                    _paid_in_hidden = [
                        (
                            getattr(_it, "sku_name", ""),
                            int(getattr(_it, "final_krw", 0) or 0),
                        )
                        for _it in line_items
                        if getattr(_it, "sku_name", "") in _hidden_set
                        and int(getattr(_it, "final_krw", 0) or 0) > 0
                    ]
                    if _paid_in_hidden:
                        st.session_state._pending_result        = _result_dict
                        st.session_state._pending_auto_dl_key   = _auto_dl_key_value
                        st.session_state._pending_hidden_paid   = _paid_in_hidden
                        st.rerun()
                    else:
                        st.session_state._last_result  = _result_dict
                        st.session_state._auto_dl_key  = _auto_dl_key_value
                        st.session_state.pop("_auto_dl_fired", None)

                except Exception as exc:
                    loading_ph.empty()
                    st.error(f"정산 중 오류가 발생했습니다:\n\n```\n{exc}\n```")

        # ── 미노출 SKU 비용 확인 dialog ───────────────────────────────────────
        # _pending_result 가 있으면 사용자에게 모달로 확인을 받는다. 계산은
        # 이미 끝났으므로 dialog 는 게이트 역할만 — 확인 시 _last_result 로
        # promote 되어 결과 영역(자동 다운로드 포함)이 정상 진행된다.
        if st.session_state.get("_pending_result"):
            _pp = st.session_state.get("_pending_hidden_paid") or []
            _pp_total = sum(int(k) for _, k in _pp)

            @st.dialog("⚠ 미노출 SKU 비용 확인")
            def _confirm_hidden_paid_dialog():
                st.markdown(
                    f"**엑셀에서 제외한 SKU 중 무료 사용량을 초과한 유료 SKU 가 "
                    f"{len(_pp)}개** 있습니다."
                )
                st.markdown(
                    "이대로 진행하면 엑셀 총액이 실제 청구액보다 "
                    f"**₩{_pp_total:,}** 만큼 적게 표시됩니다."
                )
                st.markdown("**[비용발생] 항목:**")
                for _nm, _kw in _pp:
                    st.markdown(
                        f"- <span style='color:#641e16;font-weight:600;'>"
                        f"[비용발생] {_nm}</span> &nbsp; ₩{int(_kw):,}",
                        unsafe_allow_html=True,
                    )
                st.divider()
                # 권장 액션을 가장 잘 보이게: 노출로 옮기고 재정산 (primary)
                # 그 아래에 부정적 액션(취소 / 강행) 을 작게 배치.
                if st.button(
                    "📤  노출로 옮기고 재정산",
                    type="primary",
                    use_container_width=True,
                    key="_dlg_hidden_move",
                    help=(
                        "비용발생 SKU 를 미노출 목록에서 빼고 SKU 노출 순서 "
                        "끝에 추가한 뒤, 자동으로 한 번 더 정산을 돌립니다."
                    ),
                ):
                    # 1) hidden 에서 비용발생 SKU 제거
                    _paid_names = [_nm for _nm, _ in _pp]
                    _curr_hidden = _load_hidden_skus_map().get(
                        _order_account_key, []
                    )
                    _new_hidden = [s for s in _curr_hidden if s not in _paid_names]
                    _save_hidden_skus_for_account(
                        _order_account_key, _new_hidden
                    )
                    # 2) saved_orders 끝에 추가 (이미 있으면 그대로 유지)
                    _curr_orders = _load_saved_orders().get(
                        _order_account_key, []
                    )
                    _merged_orders = list(_curr_orders) + [
                        n for n in _paid_names if n not in _curr_orders
                    ]
                    _save_order_for_account(
                        _order_account_key, _merged_orders
                    )
                    # 3) pending 정리 후 재정산 신호
                    st.session_state.pop("_pending_result",      None)
                    st.session_state.pop("_pending_auto_dl_key", None)
                    st.session_state.pop("_pending_hidden_paid", None)
                    st.session_state["_rerun_billing"] = True
                    st.toast(
                        f"📤 비용발생 SKU {len(_paid_names)}개를 노출로 옮겼습니다. "
                        "재정산을 시작합니다.",
                        icon="🔁",
                    )
                    st.rerun()

                _c1, _c2 = st.columns(2)
                with _c1:
                    if st.button(
                        "취소",
                        use_container_width=True,
                        key="_dlg_hidden_cancel",
                    ):
                        st.session_state.pop("_pending_result",      None)
                        st.session_state.pop("_pending_auto_dl_key", None)
                        st.session_state.pop("_pending_hidden_paid", None)
                        st.toast("정산을 취소했습니다.", icon="↩")
                        st.rerun()
                with _c2:
                    if st.button(
                        "계속 진행 (총액 불일치)",
                        use_container_width=True,
                        key="_dlg_hidden_confirm",
                    ):
                        st.session_state._last_result = (
                            st.session_state.pop("_pending_result")
                        )
                        st.session_state._auto_dl_key = (
                            st.session_state.pop("_pending_auto_dl_key", None)
                        )
                        st.session_state.pop("_pending_hidden_paid", None)
                        st.session_state.pop("_auto_dl_fired", None)
                        st.rerun()

            _confirm_hidden_paid_dialog()

        # ── 정산 완료 후: 자동 다운로드 트리거 + 수동 다운로드 버튼 ────────────
        result = st.session_state.get("_last_result")
        if result and result.get("line_items"):
            _excel_bytes = result.get("excel_bytes")
            _pdf_bytes   = result.get("pdf_bytes")
            _fname_xlsx  = result.get("excel_filename")
            _fname_pdf   = result.get("pdf_filename")
            _pdf_error   = result.get("pdf_error")

            # 정합성 경고 — 외부 발송 전 사용자가 반드시 확인하도록 빨간 박스로 노출
            _val_w = result.get("validation_warnings") or []
            if _val_w:
                st.error(
                    f"🚨 **엑셀 정합성 경고 {len(_val_w)}건** — 외부 발송 전 확인하세요."
                )
                with st.expander("🚨 정합성 경고 상세", expanded=True):
                    for _w in _val_w[:15]:
                        st.markdown(f"- {_w}")
                    if len(_val_w) > 15:
                        st.markdown(f"- … 외 {len(_val_w)-15}건")

            if _pdf_error:
                st.warning(f"⚠ {_pdf_error}")

            # 자동 다운로드 (이번 결과에 대해 1회만 발사)
            # ─────────────────────────────────────────────────────────────
            # 접근 방식: <a data:...> 직접 클릭은 DOMPurify 의 sanitize 영향을
            # 받아 메인 페이지에서 동작 안 함. 대신 아래에 렌더되는 Streamlit
            # 내장 st.download_button (검증된 blob 다운로드)을 JS 로 자동 클릭.
            #   - st.iframe 같은 1px 자리도 안 남음
            #   - 6월 1일 이후 deprecated API 의존성 없음
            _dl_key = st.session_state.get("_auto_dl_key")
            if _dl_key and st.session_state.get("_auto_dl_fired") != _dl_key:
                st.html(
                    f"""
                    <script>
                    (function(){{
                      const KEY = {json.dumps(_dl_key)};
                      if (window.__sphAutoDlKey === KEY) return;
                      window.__sphAutoDlKey = KEY;
                      function tryClick(){{
                        const btns = Array.from(document.querySelectorAll('button'))
                          .filter(b => (b.textContent || '').indexOf('다시 다운로드') !== -1);
                        if (!btns.length) return false;
                        // Excel(0) 즉시, PDF(1) 는 900ms 뒤 — 멀티 다운 차단 회피
                        btns.forEach((b, i) => setTimeout(() => b.click(), i * 900));
                        return true;
                      }}
                      let n = 0;
                      const iv = setInterval(() => {{
                        if (tryClick() || ++n > 40) clearInterval(iv);  // 최대 4s 대기
                      }}, 100);
                    }})();
                    </script>
                    """,
                    unsafe_allow_javascript=True,
                )
                st.session_state._auto_dl_fired = _dl_key

            # 수동 재다운로드 버튼
            _btn_cols = st.columns(2)
            if _excel_bytes:
                with _btn_cols[0]:
                    st.download_button(
                        label=f"⬇  {_fname_xlsx}  다시 다운로드",
                        data=_excel_bytes,
                        file_name=_fname_xlsx,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            if _pdf_bytes:
                with _btn_cols[1]:
                    st.download_button(
                        label=f"⬇  {_fname_pdf}  다시 다운로드",
                        data=_pdf_bytes,
                        file_name=_fname_pdf,
                        mime="application/pdf",
                        use_container_width=True,
                    )

        # ── 결과 출력 ─────────────────────────────────────────────────────────
        if result:
            line_items  = result["line_items"]
            company_out = result["company"]
            _res_bm     = result["billing_month"]
            _res_ex     = result["exchange_rate"]
            _res_mr     = result["margin_rate"]
            _res_bank   = result["bank_name"]
            _res_proj   = result["proj_results"]
            _res_price  = result.get("price_list_file")

            if not line_items:
                st.warning(
                    "처리된 청구 항목이 없습니다. "
                    "SKU 마스터의 SKU ID가 파일의 데이터와 일치하는지 확인하세요."
                )
            else:
                st.markdown(f"""
                <div style="
                    background:linear-gradient(135deg,#a5d15a,#7dbb26);
                    border-radius:14px; padding:16px 26px; margin-bottom:6px;
                    color:#1b3d06; font-weight:700; font-size:1.02rem;
                    box-shadow:0 4px 16px rgba(165,209,90,0.3);
                ">
                    🎉 정산 완료 &nbsp;·&nbsp; {company_out or '전체'}
                    &nbsp;·&nbsp; 총 {len(line_items)}개 항목
                </div>
                """, unsafe_allow_html=True)

                # 정산완료 배너 직하단에 누락 SKU 경고 (있을 때만).
                _missing_skus = result.get("missing_skus") or []
                if _missing_skus:
                    _lines = "\n".join(
                        f"• **{_nm or '(이름없음)'}** (`{_sid}`)"
                        for _sid, _nm in _missing_skus
                    )
                    st.warning(
                        f"⚠️ **CSV 에 사용량이 있지만 Price List 에서 단가를 찾지 못해 집계에서 제외된 SKU {len(_missing_skus)}건**\n\n"
                        f"{_lines}\n\n"
                        "→ Price List(xlsx) A열의 SKU 명과 CSV 'SKU 설명' 이 정확히 일치해야 매칭됩니다."
                    )


# (render_batch_billing_ui 정의는 호출 위치보다 위인 _run_batch_single_billing
# 직후로 이동했다. streamlit 은 스크립트를 top-down 실행하므로 호출 시점에
# 함수가 정의되어 있어야 NameError 가 안 난다.)
def _legacy_render_batch_billing_ui_DEPRECATED(
    *,
    tmp_input_path: Path,
    companies: list[str],
    billing_month: str,
    price_list_file,
    currency: str,
    billable_skus,
):
    """[더 이상 사용 안 함] 이 위치는 호출보다 아래라 NameError 가 발생함.
       실제 정의는 위쪽 _run_batch_single_billing 직후에 있음.
    """
    import datetime as _dt
    import zipfile as _zip
    import io as _io
    import re as _re

    st.markdown("#### 전체 일괄 정산")
    st.caption(
        "체크한 회사들을 한 번에 정산해 회사별 폴더로 zip 다운로드합니다. "
        "각 회사의 **저장된 설정**(과금방식·소수점·최소사용비용·환율 표기·"
        "미노출 SKU·직접등록·SKU 순서) 을 그대로 사용하므로 개별 정산과 결과가 일치합니다."
    )

    # ── 저장된 선택 상태 로드 + 신규 회사 표시용 known 계산 ───────────
    _saved_batch = _load_batch_selection()
    _known_set   = set(_saved_batch.get("known", []))
    _norm_known  = {_norm_account_key(k) for k in _known_set}
    # 회사명을 known 과 정규화 매칭 — 키 표기 다르면 신규로 보이지 않게.
    def _is_known(c: str) -> bool:
        return _norm_account_key(c) in _norm_known

    _new_companies = [c for c in companies if not _is_known(c)]

    # 이번 화면에서 본 회사들을 known 에 모두 누적 저장 (한 번이라도 노출되면 알려진 것).
    _updated_known = sorted(_known_set | set(companies))

    # 미리 회사별 저장 설정값 로드 (요약 표시 + 정산 호출용).
    _orders_all   = _load_saved_orders()
    _manual_all   = _load_manual_skus_map()
    _hidden_all   = _load_hidden_skus_map()
    _mode_all     = _load_billing_modes()
    _round_all    = _load_subtotal_round_map()
    _proj_flag_all= _load_include_project_flags()
    _rate_all     = _load_rate_labels()

    def _summary(c: str) -> str:
        """회사별 saved 설정 요약 (한 줄)."""
        _mode  = _lookup_account(_mode_all,  c) or BILLING_MODE_ACCOUNT
        _round = _round_all.get(c, 0 if currency == "KRW" else 2)
        _proj  = _lookup_account(_proj_flag_all, c)
        _proj  = True if _proj is None else bool(_proj)
        _amt, _cur = _min_charge_for_account(c)
        _hidden_n = len(_lookup_account(_hidden_all, c) or [])
        _manual_n = len(_lookup_account(_manual_all, c) or [])
        _mode_tag = "회사통합" if _mode == BILLING_MODE_ACCOUNT else "프로젝트별"
        _round_tag = ",0" if _round == 0 else ",2"
        _proj_tag  = "Proj✓" if _proj else "Proj✗"
        _min_tag = (
            f"최소 {_cur} {int(_amt):,}" if (_amt and _amt > 0) else "최소-"
        )
        return (
            f"{_mode_tag} · {_round_tag} · {_proj_tag} · {_min_tag} · "
            f"hidden {_hidden_n} · 직접등록 {_manual_n}"
        )

    # ── 일괄 입력 영역 (환율 · 환율 날짜) ─────────────────────────────
    _today = _dt.date.today()
    with st.container(border=True):
        st.markdown("#### 💱 일괄 입력 (USD 회사에만 적용)")
        c1, c2 = st.columns([1, 1])
        with c1:
            # 단일 모드와 동일 UX — 비어있으면 빨간 테두리(CSS placeholder-shown).
            batch_rate = st.number_input(
                "환율 (₩/$)",
                min_value=0.0, value=None, step=0.01, format="%.2f",
                placeholder="예: 1427.87",
                key="_batch_rate_input",
                help="USD 단가표 회사들에만 적용. KRW 회사는 환율 무관.",
            )
        with c2:
            batch_rate_date = st.date_input(
                "환율 날짜",
                value=_today,
                key="_batch_rate_date",
                format="YYYY-MM-DD",
            )
        _batch_rate_date_str = batch_rate_date.strftime("%Y.%m.%d")

    # ── 정책 라디오 ────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("#### ⚙️ 비용발생 SKU 정책")
        _policy_options = {
            "그대로 진행 (기본)": "as_is",
            "노출로 옮기고 재정산": "move",
            "건너뛰기 (해당 회사 정산 안 함)": "skip",
        }
        _policy_labels = list(_policy_options.keys())
        _saved_policy = _saved_batch.get("policy", "as_is")
        _policy_idx = next(
            (i for i, lb in enumerate(_policy_labels)
             if _policy_options[lb] == _saved_policy), 0,
        )
        _policy_label = st.radio(
            "hidden 안에 무료 한도 초과 SKU 가 발견되었을 때:",
            options=_policy_labels, index=_policy_idx, horizontal=False,
            key="_batch_policy",
            help=(
                "• 그대로 진행: 엑셀 총액이 실제 청구액보다 적게 표시될 수 있음(주의).\n"
                "• 노출로 옮김: 해당 SKU 를 hidden 에서 빼고 saved_orders 끝에 추가 후 재정산. "
                "saved 데이터가 영구 변경됩니다.\n"
                "• 건너뛰기: 그 회사는 정산 결과에서 제외됨."
            ),
        )
        batch_policy = _policy_options[_policy_label]

    # ── 다운로드 옵션 ──────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("#### 📥 다운로드 옵션")
        _pdf_ok = _pdf_export_available()
        dl_xlsx = st.checkbox(
            "📗 엑셀 (.xlsx)",
            value=bool(_saved_batch.get("dl_xlsx", True)),
            key="_batch_dl_xlsx",
        )
        dl_pdf = st.checkbox(
            ("📄 PDF" if _pdf_ok else "📄 PDF (현재 환경에서 변환 불가)"),
            value=bool(_saved_batch.get("dl_pdf", False)) and _pdf_ok,
            key="_batch_dl_pdf",
            disabled=not _pdf_ok,
        )

    # ── 회사 리스트 (체크박스 + 신규 마커 + 설정 요약) ───────────────
    with st.container(border=True):
        st.markdown(f"#### 📋 회사 선택 ({len(companies)}개 / 신규 {len(_new_companies)}개)")
        # 전체 선택/해제
        cc1, cc2, cc3 = st.columns([1, 1, 4])
        with cc1:
            if st.button("✅ 전체 선택", key="_batch_check_all", use_container_width=True):
                st.session_state["_batch_force_all"] = True
                st.rerun()
        with cc2:
            if st.button("⬜ 전체 해제", key="_batch_uncheck_all", use_container_width=True):
                st.session_state["_batch_force_none"] = True
                st.rerun()
        with cc3:
            st.caption(
                "🆕 = 이전에 본 적 없는 회사 (이번 CSV 에서 새로 발견). "
                "체크/해제 상태는 다음 방문 시 자동 복원됩니다."
            )

        _saved_selected_norm = {_norm_account_key(k) for k in _saved_batch.get("selected", [])}
        _force_all  = st.session_state.pop("_batch_force_all", False)
        _force_none = st.session_state.pop("_batch_force_none", False)

        _checked: list[str] = []
        for c in companies:
            _is_new = not _is_known(c)
            # 기본값: 강제 토글 → 저장된 선택 → 신규는 기본 ON
            if _force_all:
                _default = True
            elif _force_none:
                _default = False
            elif _norm_account_key(c) in _saved_selected_norm:
                _default = True
            elif _is_new:
                _default = True
            else:
                _default = False

            _col_chk, _col_txt = st.columns([1, 18])
            with _col_chk:
                _chk = st.checkbox(
                    "", value=_default,
                    key=f"_batch_chk_{_norm_account_key(c)}",
                    label_visibility="collapsed",
                )
            with _col_txt:
                _marker = " 🆕 **신규**" if _is_new else ""
                st.markdown(
                    f"**{c}**{_marker}  \n"
                    f"<span style='color:#7a8a90;font-size:0.78rem;'>"
                    f"{_summary(c)}</span>",
                    unsafe_allow_html=True,
                )
            if _chk:
                _checked.append(c)

    # ── 정산 시작 ─────────────────────────────────────────────────
    st.divider()
    if not _checked:
        st.info("정산할 회사를 1개 이상 체크해 주세요.")
        return

    if not (dl_xlsx or dl_pdf):
        st.info("다운로드 형식(엑셀/PDF) 을 1개 이상 선택해 주세요.")
        return

    # 환율 미입력 차단 — 빨간 테두리로 시각적 안내. 빈 상태로 시작 막음.
    if st.session_state.get("_batch_rate_input") is None:
        st.warning("환율을 입력해 주세요. (빨간 테두리 표시 영역)")
        return

    _start = st.button(
        f"▶ 전체 정산 시작 ({len(_checked)}개사)",
        type="primary", use_container_width=True,
        key="_batch_run_btn",
    )

    # 선택 상태 + 옵션 자동 저장 (start 클릭 무관, 매 rerun 마다 반영)
    _save_batch_selection({
        "selected": _checked,
        "known":    _updated_known,
        "dl_xlsx":  dl_xlsx,
        "dl_pdf":   dl_pdf,
        "policy":   batch_policy,
    })

    if not _start:
        return
    if price_list_file is None:
        st.error("Price List(xlsx) 가 없습니다. 사이드바에서 업로드해 주세요.")
        return

    # ── 일괄 실행 루프 ───────────────────────────────────────────
    overlay_ph = st.empty()    # 정산 중 화면 dim + 가운데 진행 카드
    _timings_lines: list[str] = []   # 결과 expander 표시용 (UI 실시간 표시는 오버레이가 담당)
    import time as _t_perf3
    _t_loop_start2 = _t_perf3.time()
    log_lines:   list[str] = []
    results:     list[dict] = []
    safe_re = _re.compile(r'[\\/*?:"<>|]')

    zip_buf = _io.BytesIO()
    _finished_count = 0  # 누적 완료 회사 수 (ETA 계산용)
    _ema_per_company: float | None = None
    _EMA_ALPHA = 0.3
    with _zip.ZipFile(zip_buf, "w", _zip.ZIP_DEFLATED) as zf:
        total = len(_checked)
        for idx, c in enumerate(_checked, 1):
            _remaining = (
                _ema_per_company * (total - _finished_count)
                if _ema_per_company is not None and total > _finished_count
                else None
            )
            _render_batch_overlay(
                overlay_ph, idx=idx, total=total, company=c,
                remaining_seconds=_remaining,
            )
            _t_company_start = _t_perf3.time()

            # 회사별 saved 값 로드 (lookup 폴백 포함)
            _saved_for   = _lookup_account(_orders_all,   c) or []
            _manual_for  = _lookup_account(_manual_all,   c) or []
            _hidden_for  = _lookup_account(_hidden_all,   c) or []
            _mode        = _lookup_account(_mode_all,     c) or BILLING_MODE_ACCOUNT
            _round_val   = _round_all.get(c, 0 if currency == "KRW" else 2)
            _proj_flag   = _lookup_account(_proj_flag_all, c)
            _proj_flag   = True if _proj_flag is None else bool(_proj_flag)
            _rl          = _lookup_account(_rate_all,     c) or {}
            _bank        = _rl.get("bank", "")  or DEFAULT_BANK_NAME
            _phrase      = _rl.get("phrase", "") or DEFAULT_RATE_PHRASE
            _extra       = _rl.get("extra", "")
            _min_amt, _min_cur = _min_charge_for_account(c)

            # 1차 정산
            res = _run_batch_single_billing(
                selected_company=c,
                billing_month=billing_month,
                tmp_input_path=tmp_input_path,
                price_list_file=price_list_file,
                currency=currency,
                exchange_rate=float(batch_rate or 0),
                margin_rate=1.0,
                rate_date_str=_batch_rate_date_str,
                billing_mode=_mode,
                include_project_sheet=_proj_flag,
                subtotal_round=int(_round_val),
                bank_name=_bank,
                rate_phrase_text=_phrase,
                rate_extra_text=_extra,
                min_charge_amount=float(_min_amt or 0),
                min_charge_currency=_min_cur or "KRW",
                sku_order=_saved_for,
                manual_skus=_manual_for,
                hidden_skus=_hidden_for,
                billable_skus=billable_skus,
                dl_xlsx=dl_xlsx,
                dl_pdf=dl_pdf,
            )

            # 비용발생 정책 처리
            _paid_names = [nm for nm, _ in (res.get("paid_in_hidden") or [])]
            policy_applied = None
            if _paid_names:
                if batch_policy == "skip":
                    results.append({
                        "company": c, "status": "skipped",
                        "paid_in_hidden": res.get("paid_in_hidden") or [],
                        "error": "정책=건너뛰기",
                    })
                    log_lines.append(f"⏭ **{c}** — 비용발생 SKU {len(_paid_names)}개 → 건너뜀")
                    _ema_per_company = _update_ema(
                        _ema_per_company, _t_perf3.time() - _t_company_start, _EMA_ALPHA,
                    )
                    _finished_count += 1
                    continue
                elif batch_policy == "move":
                    # hidden 에서 제거 + saved_orders 끝에 추가 → saved 영구 변경
                    _new_hidden = [s for s in _hidden_for if s not in _paid_names]
                    _save_hidden_skus_for_account(c, _new_hidden)
                    _new_saved = list(_saved_for) + [
                        n for n in _paid_names if n not in _saved_for
                    ]
                    _save_order_for_account(c, _new_saved)
                    policy_applied = "moved"
                    # 재정산
                    res = _run_batch_single_billing(
                        selected_company=c,
                        billing_month=billing_month,
                        tmp_input_path=tmp_input_path,
                        price_list_file=price_list_file,
                        currency=currency,
                        exchange_rate=float(batch_rate or 0),
                        margin_rate=1.0,
                        rate_date_str=_batch_rate_date_str,
                        billing_mode=_mode,
                        include_project_sheet=_proj_flag,
                        subtotal_round=int(_round_val),
                        bank_name=_bank,
                        rate_phrase_text=_phrase,
                        rate_extra_text=_extra,
                        min_charge_amount=float(_min_amt or 0),
                        min_charge_currency=_min_cur or "KRW",
                        sku_order=_new_saved,
                        manual_skus=_manual_for,
                        hidden_skus=_new_hidden,
                        billable_skus=billable_skus,
                        dl_xlsx=dl_xlsx,
                        dl_pdf=dl_pdf,
                    )

            if not res["ok"]:
                results.append({
                    "company": c, "status": "error",
                    "error": res.get("error"),
                    "paid_in_hidden": [],
                })
                log_lines.append(f"❌ **{c}** — {res.get('error')}")
                _ema_per_company = _update_ema(
                    _ema_per_company, _t_perf3.time() - _t_company_start, _EMA_ALPHA,
                )
                _finished_count += 1
                continue

            # zip 에 파일 저장 — 회사별 폴더
            _safe = safe_re.sub("_", c).strip() or "전체"
            _stem = f"sGMP_Invoice_{_safe}"
            if dl_xlsx and res.get("excel_bytes"):
                zf.writestr(f"{_safe}/{_stem}.xlsx", res["excel_bytes"])
            if dl_pdf and res.get("pdf_bytes"):
                zf.writestr(f"{_safe}/{_stem}.pdf", res["pdf_bytes"])

            _val_w = res.get("validation_warnings") or []
            results.append({
                "company": c,
                "status": "ok" if not res.get("paid_in_hidden") else "ok_with_paid",
                "paid_in_hidden": res.get("paid_in_hidden") or [],
                "pdf_error": res.get("pdf_error"),
                "policy_applied": policy_applied,
                "validation_warnings": _val_w,
            })
            _hint = ""
            if res.get("paid_in_hidden") and not policy_applied:
                _hint = f" · ⚠ 비용발생 {len(res['paid_in_hidden'])}개"
            elif policy_applied == "moved":
                _hint = f" · 🔁 노출이동·재정산 ({len(_paid_names)}개)"
            if _val_w:
                _hint += f" · 🚨 정합성 경고 {len(_val_w)}건"
            log_lines.append(f"✅ **{c}**{_hint}")
            _ema_per_company = _update_ema(
                _ema_per_company, _t_perf3.time() - _t_company_start, _EMA_ALPHA,
            )
            # 결과 expander 표시용 timing 라인만 수집 (UI 실시간 표시는 오버레이가 담당)
            _timings_lines.append(
                _format_billing_timings_line(c, res.get("timings") or {})
            )
            if _val_w:
                _timings_lines.append(
                    f"  🚨 **{c}** 정합성 경고: " + " / ".join(_val_w[:2])
                    + (f" (외 {len(_val_w)-2}건)" if len(_val_w) > 2 else "")
                )
            _finished_count += 1

        _t_loop_total2 = _t_perf3.time() - _t_loop_start2
        _timings_lines.append(
            f"**전체 완료: {_t_loop_total2:.2f}초 ({len(_checked)}개사)**"
        )

    # 루프 종료 — 오버레이 제거
    overlay_ph.empty()

    # ── 결과 요약 + 다운로드 버튼 ──────────────────────────────────
    n_ok     = sum(1 for r in results if r["status"] in ("ok", "ok_with_paid"))
    n_paid   = sum(1 for r in results if r["status"] == "ok_with_paid")
    n_skip   = sum(1 for r in results if r["status"] == "skipped")
    n_err    = sum(1 for r in results if r["status"] == "error")
    n_total  = len(results)

    st.markdown("---")
    st.markdown(f"### 결과 요약 ({n_ok}/{n_total} 성공)")
    if n_paid > 0:
        st.warning(
            f"⚠️ 비용발생 SKU 가 포함된 회사 **{n_paid}개** — 엑셀 총액이 실제 "
            "청구액보다 적게 표시되었을 수 있습니다. 아래 회사 행을 확인하세요."
        )
    if n_err > 0:
        st.error(f"❌ 정산 실패 {n_err}개사 — 아래 로그 확인")
    if n_skip > 0:
        st.info(f"⏭ 건너뛴 회사 {n_skip}개사 (정책=건너뛰기)")

    # 정합성 경고 — 외부 발송 전 사용자 확인 필요
    _val_alerts = [r for r in results if r.get("validation_warnings")]
    if _val_alerts:
        st.error(
            f"🚨 **엑셀 정합성 경고가 발생한 회사 {len(_val_alerts)}개** — "
            "외부 발송 전 반드시 확인하세요."
        )
        with st.expander("🚨 정합성 경고 상세", expanded=True):
            for r in _val_alerts:
                st.markdown(f"**{r['company']}**")
                for _w in r["validation_warnings"][:10]:
                    st.markdown(f"- {_w}")
                _rem = len(r["validation_warnings"]) - 10
                if _rem > 0:
                    st.markdown(f"- … 외 {_rem}건")

    # 회사별 로그
    with st.expander("📋 회사별 정산 로그", expanded=False):
        for ln in log_lines:
            st.markdown(ln)
        # 비용발생 상세
        _has_paid = [r for r in results if r.get("paid_in_hidden")]
        if _has_paid:
            st.markdown("---")
            st.markdown("**[비용발생] 상세:**")
            for r in _has_paid:
                _items = ", ".join(
                    f"{nm}(₩{kw:,})" for nm, kw in r["paid_in_hidden"]
                )
                st.markdown(f"- {r['company']}: {_items}")

    # 단계별 소요시간은 디버그용 expander 에 별도 노출 (기본 닫힘)
    if _timings_lines:
        with st.expander("⏱ 단계별 소요시간 (디버그)", expanded=False):
            st.markdown("  \n".join(_timings_lines))

    # zip 다운로드
    if n_ok > 0:
        _ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        _zip_name = f"전체정산_{billing_month or 'all'}_{_ts}.zip"
        st.download_button(
            f"📦 zip 다운로드 ({_zip_name})",
            data=zip_buf.getvalue(),
            file_name=_zip_name,
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )
    else:
        st.info("다운로드할 결과가 없습니다.")



