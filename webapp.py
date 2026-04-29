"""
webapp.py — SPH GMP 정산 자동화 시스템 v2

streamlit run webapp.py
"""
from __future__ import annotations

import base64
import hashlib
import json
import tempfile
from datetime import date
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
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
from invoice_generator import generate_formatted_invoice

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
    "최종 송금환율 기준",
    "최종 매매기준율 기준",
    "최종 전신환매도율 기준",
    "최종 전신환매입률 기준",
    "고시환율 기준",
]
DEFAULT_RATE_PHRASE = "최종 송금환율 기준"
DEFAULT_BANK_NAME   = "하나은행"

# ─── 테스트 모드 (기간 한정 — 아래 False 로 바꾸면 전부 해제됨) ──────────────
# 활성화 시:
#   * 환율 입력값을 1480.80 으로 자동 프리필
#   * 결제 계정 선택을 "hanatour" 가 포함된 항목으로 자동 선택
# 종료 시: _TEST_DEFAULTS = False 한 줄만 바꾸면 됨.
_TEST_DEFAULTS          = True
_TEST_DEFAULT_RATE      = "1480.80"
_TEST_DEFAULT_COMPANY_KW = "hanatour"


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


def _save_order_for_account(account: str, order: list[str]) -> None:
    data = _load_saved_orders()
    data[account] = [str(n) for n in order if n]
    SAVED_ORDERS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_ORDERS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


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
        out[acc] = {
            "bank":   str(v.get("bank")   or DEFAULT_BANK_NAME),
            "phrase": str(v.get("phrase") or DEFAULT_RATE_PHRASE),
            "extra":  str(v.get("extra")  or ""),
            "date":   v.get("date") if isinstance(v.get("date"), str) else None,
        }
    return out


def _save_rate_label_for_account(account: str, bank: str, phrase: str,
                                  extra: str, date_str: str | None) -> None:
    data = _load_rate_labels()
    data[account] = {
        "bank":   (bank or DEFAULT_BANK_NAME).strip(),
        "phrase": (phrase or DEFAULT_RATE_PHRASE).strip(),
        "extra":  (extra or "").strip(),
        "date":   (date_str or None),
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
    }


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
</style>
""", unsafe_allow_html=True)


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


@st.cache_data(show_spinner=False)
def _get_file_preview(tmp_path: str) -> dict:
    """파일 기본 통계 (전체 계정, preview 전용)."""
    try:
        rows = _cached_preprocess(tmp_path, "0000-00", None)
        return {
            "row_count":   len(rows),
            "proj_count":  len({r["project_id"] for r in rows}),
            "total_usage": sum(r["usage_amount"] for r in rows),
        }
    except Exception:
        return {"row_count": 0, "proj_count": 0, "total_usage": 0}


# ── 세션 상태 초기화 ──────────────────────────────────────────────────────────
if "master_df" not in st.session_state:
    st.session_state.master_df = _load_master_df()

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:22px 0 18px;">
        <div style="
            background:rgba(255,255,255,0.13); border-radius:16px;
            padding:16px 0; margin-bottom:8px;
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
components.html(
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
    height=0,
)

# 설정 변경 감지는 col_right 내부(통화·환율 위젯 직후)로 이동됨.

# ── 메인 헤더 ─────────────────────────────────────────────────────────────────
st.markdown("""
<div style="
    display:flex; align-items:center; gap:18px;
    padding:6px 2px 22px;
    border-bottom:2px solid #d4e2e6;
    margin-bottom:22px;
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
components.html("""
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
""", height=0)

# ═══════════════════════════════════════════════════════════════════════════════
# 통합 정산 실행 (flat — 탭 제거)
# ═══════════════════════════════════════════════════════════════════════════════
if True:

    # ── ① 단가표(GMP Price List) 첨부 — USD/KRW 통화별 분리 ───────────────
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
        """업로드된 단가표를 지정 경로에 저장하고 한 번만 flash 메시지."""
        _key = f"{uploaded.name}_{uploaded.size}"
        if st.session_state.get(f"_saved_price_key_{session_tag}") != _key:
            uploaded.seek(0)
            save_path.parent.mkdir(parents=True, exist_ok=True)
            save_path.write_bytes(uploaded.read())
            st.session_state[f"_saved_price_key_{session_tag}"] = _key
            st.session_state[f"_price_flash_{session_tag}"] = uploaded.name
            st.rerun()

    _col_pl_usd, _col_pl_krw = st.columns(2, gap="medium")

    with _col_pl_usd:
        _uploaded_price_usd = st.file_uploader(
            "📋 달러($) 단가표",
            type=["xlsx", "xls"],
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
                    st.session_state.pop("_saved_price_key_usd", None)
                    st.rerun()
        if _f_usd := st.session_state.pop("_price_flash_usd", None):
            st.success(f"✅ 달러 단가표 저장 — {_f_usd}")

    with _col_pl_krw:
        _uploaded_price_krw = st.file_uploader(
            "📋 원화(₩) 단가표",
            type=["xlsx", "xls"],
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

    # ── ② 사용고지서 업로드 ───────────────────────────────────────────────────
    uploaded_file = st.file_uploader(
        "📂 구글 Maps 플랫폼 사용고지서",
        type=["csv", "xlsx", "xls"],
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
            st.session_state._preview   = _get_file_preview(str(tmp_path))
            # 정산월 자동 감지 — CSV '인보이스 날짜' 에서 YYYY-MM 추출
            _detected_bm = _detect_billing_month(str(tmp_path))
            if _detected_bm:
                st.session_state._auto_billing_month = _detected_bm
            st.session_state.pop("_last_result", None)  # 이전 결과 초기화
            st.rerun()   # 사이드바의 자동 감지 정산월 표시를 즉시 반영

        tmp_input_path = Path(st.session_state._tmp_path)
        companies      = st.session_state._companies
        preview        = st.session_state._preview

        # ── 파일 분석 미리보기 배너 ───────────────────────────────────────────
        st.markdown(f"""
        <div style="
            background:linear-gradient(135deg,#00788a 0%,#005a6a 100%);
            border-radius:18px; padding:20px 28px; margin:18px 0 6px;
            box-shadow:0 6px 22px rgba(0,120,138,0.2);
        ">
            <div style="
                font-size:0.72rem; font-weight:700;
                color:rgba(255,255,255,0.6); margin-bottom:14px;
                letter-spacing:0.8px; text-transform:uppercase;
            ">📋 파일 분석 결과 &nbsp;·&nbsp; {uploaded_file.name}</div>
            <div style="display:flex; gap:36px; flex-wrap:wrap; align-items:flex-end;">
                <div>
                    <div style="font-size:1.75rem; font-weight:800; color:#ffffff; line-height:1.1;">
                        {preview['row_count']:,}
                    </div>
                    <div style="font-size:0.72rem; color:rgba(255,255,255,0.6); margin-top:5px;">
                        집계 데이터 행
                    </div>
                </div>
                <div>
                    <div style="font-size:1.75rem; font-weight:800; color:#a5d15a; line-height:1.1;">
                        {len(companies):,}
                    </div>
                    <div style="font-size:0.72rem; color:rgba(255,255,255,0.6); margin-top:5px;">
                        결제 계정 수
                    </div>
                </div>
                <div>
                    <div style="font-size:1.75rem; font-weight:800; color:#ffd97a; line-height:1.1;">
                        {preview['proj_count']:,}
                    </div>
                    <div style="font-size:0.72rem; color:rgba(255,255,255,0.6); margin-top:5px;">
                        프로젝트 수
                    </div>
                </div>
                <div>
                    <div style="font-size:1.75rem; font-weight:800; color:#ffffff; line-height:1.1;">
                        {preview['total_usage']:,}
                    </div>
                    <div style="font-size:0.72rem; color:rgba(255,255,255,0.6); margin-top:5px;">
                        총 사용량 (건)
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ── 결제 계정 선택 (전체 폭) ─────────────────────────────────────────
        st.markdown("#### 정산 대상 선택")
        if companies:
            # 테스트 모드: 지정 키워드(기본 'hanatour') 가 포함된 항목을 자동 선택.
            _default_company_idx = 0
            if _TEST_DEFAULTS and _TEST_DEFAULT_COMPANY_KW:
                for _i, _c in enumerate(companies):
                    if _TEST_DEFAULT_COMPANY_KW.lower() in str(_c).lower():
                        _default_company_idx = _i
                        break
            selected_company = st.selectbox(
                "결제 계정 (Billing Account Name)",
                options=companies,
                index=_default_company_idx,
            )
        else:
            selected_company = None
            st.info("파일에서 결제 계정 정보를 찾을 수 없습니다. 전체 데이터를 처리합니다.")

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
        with col_left, st.container(border=True):
            if _found_skus:
                _saved_orders = _load_saved_orders()
                _saved_for_this = _saved_orders.get(_order_account_key, [])

                # 저장된 순서 + 이번 CSV에 신규로 등장한 항목 분리
                _existing   = [n for n in _saved_for_this if n in _found_skus]
                _new_items  = [n for n in _found_skus if n not in _existing]
                _NEW_COUNT  = len(_new_items)

                st.markdown(f"#### 📋 엑셀 SKU 노출 순서 ({len(_found_skus)}개)")
                if _new_items and _existing:
                    st.caption(
                        "연한 초록 배경 = **[신규 항목]**. 드래그로 위치 조정 후 "
                        "**현재 순서 저장** 을 누르세요."
                    )
                elif _new_items and not _existing:
                    st.caption(
                        "모두 새로 발견된 SKU입니다. 드래그로 순서 지정 후 "
                        "**현재 순서 저장** 을 누르세요."
                    )
                else:
                    st.caption("드래그하여 엑셀에 나올 순서를 조정하세요.")

                # 신규 항목은 라벨에 [신규 항목] prefix로 표시 (저장 전 제거)
                _NEW_PREFIX = "[신규 항목] "
                _initial = _existing + [f"{_NEW_PREFIX}{n}" for n in _new_items]

                # sortable 컴포넌트 key: 입력 항목이 변경되면 새 key로 캐시 초기화
                _items_fingerprint = hashlib.md5(
                    "\x1f".join(_initial).encode("utf-8")
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
                _new_css = ""
                if _NEW_COUNT > 0:
                    _new_css = f"""
                    .sortable-item:nth-last-child(-n+{_NEW_COUNT}),
                    .sortable-item:nth-last-child(-n+{_NEW_COUNT}):hover,
                    .sortable-item:nth-last-child(-n+{_NEW_COUNT}):focus {{
                        background: linear-gradient(135deg, #f2fae3 0%, #e4f2cd 100%) !important;
                        background-color: transparent !important;
                        border-color: #cfe7a8 !important;
                        border-left-color: #7dbb26 !important;
                        color: #1b3d06 !important;
                    }}
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
                }
                .sortable-item:active { cursor: grabbing !important; }
                """ + _new_css

                _reordered = sort_items(
                    _initial,
                    direction="vertical",
                    custom_style=_SORT_CSS,
                    key=_order_state_key,
                )

                # prefix 제거해 실제 SKU 순서 확정
                sku_order = [
                    (x[len(_NEW_PREFIX):] if x.startswith(_NEW_PREFIX) else x)
                    for x in _reordered
                ]

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

                # ── 엑셀 미노출 SKU (수동) ──────────────────────────────
                # 엔진 계산(waterfall/sku_master/line_items) 에는 영향 주지 않음.
                # `generate_formatted_invoice` 호출 직전에 line_items / proj_results
                # 에서 sku_name 매칭 항목만 제거 → 출력물에서만 빠진다.
                _saved_hidden_all = _load_hidden_skus_map()
                _saved_hidden_for_this = _saved_hidden_all.get(_order_account_key, [])
                # 현재 CSV 에 존재하지 않는 항목은 자동 필터 (stale 저장값 방어)
                _hidden_default = [s for s in _saved_hidden_for_this if s in sku_order]

                hidden_skus = st.multiselect(
                    "🚫 엑셀에서 제외할 SKU (선택)",
                    options=sku_order,
                    default=_hidden_default,
                    key=f"_hidden_skus_ms::{_order_account_key}",
                    help=(
                        "선택한 SKU 는 Invoice / Project 시트에 출력되지 않습니다.\n"
                        "엔진 계산(waterfall·무료 배분) 은 그대로 유지되며, **출력 직전**에만 제거됩니다.\n"
                        "완전 무료(subtotal $0) 항목을 숨기면 총액 변동 없음.\n"
                        "유료 항목을 숨기면 그만큼 청구 총액이 줄어드니 주의."
                    ),
                )
                # 변경 즉시 자동 저장 (다음 세션에 복원)
                if sorted(hidden_skus) != sorted(_saved_hidden_for_this):
                    _save_hidden_skus_for_account(_order_account_key, hidden_skus)
                    st.toast(
                        f"💾 '{_order_account_key}' 미노출 SKU {len(hidden_skus)}개 저장",
                        icon="🚫",
                    )
            else:
                sku_order = []
                hidden_skus = []
                st.info("표시할 SKU가 없습니다. CSV를 확인하세요.")

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
                # 구간별(tier) 단가·금액의 표시 포맷도 동일 자리수로 맞춤.
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
                        "Invoice 시트의 소계와 구간별(tier) 단가·금액 표시 자리수. "
                        "계정별로 저장됩니다. 셀 수식/값은 그대로이고 표시 포맷만 "
                        "달라지므로 결과값에는 영향이 없습니다."
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
                # 테스트 모드: 최초 렌더 시 기본값 프리필 (사용자가 이후 지울 수 있음).
                if (_TEST_DEFAULTS
                        and "_rate_raw_input" not in st.session_state
                        and not st.session_state.get("_rate_raw")):
                    st.session_state["_rate_raw_input"] = _TEST_DEFAULT_RATE
                _rate_raw = st.text_input(
                    "환율 (USD → KRW)",
                    value=st.session_state.get("_rate_raw", ""),
                    max_chars=7,
                    placeholder="예: 1427.87" if currency == "USD" else "원화 모드 — 입력 불필요",
                    key="_rate_raw_input",
                    disabled=(currency == "KRW"),
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
                )
                if _bank_sel == "직접입력":
                    _bank_typed = st.text_input(
                        "은행명 직접입력",
                        value=(_saved_bank if _saved_bank not in MAJOR_BANKS else ""),
                        key=f"_bank_typed::{_order_account_key}",
                        help="한 글자만 입력해도 MAJOR_BANKS 에서 자동 매칭 ('하' → 하나은행)",
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
                )
                if _phrase_sel == "직접입력":
                    rate_phrase_text = st.text_input(
                        "고정문구 직접입력",
                        value=(_saved_phrase if _saved_phrase not in RATE_PHRASES else ""),
                        key=f"_phrase_typed::{_order_account_key}",
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
            # 환율은 달러 모드에서만 필수 (원화 모드는 환율 미사용)
            if currency == "USD" and (not exchange_rate or exchange_rate <= 0):
                _missing_msgs.append("환율을 입력해 주세요")
            if not (dl_excel or dl_pdf):
                _missing_msgs.append("다운로드 형식(엑셀 / PDF)을 하나 이상 선택해주세요.")

            if _missing_msgs:
                for _m in _missing_msgs:
                    st.toast(f"⚠ {_m}", icon="⚠️")
                st.error(" / ".join(_missing_msgs))
                run_button = False   # 아래 정산 블록 실행 차단

        # ── 정산 실행 ─────────────────────────────────────────────────────────
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
                    # → 엔진이 조용히 버리지 않도록 UI 에 경고.
                    _missing_skus = detect_missing_skus(usage_rows, sku_master)
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
                    _fname_xlsx  = f"정산리포트_{_safe}_{billing_month}.xlsx"
                    _fname_pdf   = f"정산리포트_{_safe}_{billing_month}.pdf"
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
                    )

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

                    st.session_state._last_result = {
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
                    }
                    # 자동 다운로드용 키 — 같은 결과를 재 다운로드하지 않도록
                    st.session_state._auto_dl_key = (
                        f"{selected_company}|{billing_month}|{len(_excel_bytes)}|"
                        f"{'X' if dl_excel else '-'}{'P' if dl_pdf else '-'}"
                    )
                    st.session_state.pop("_auto_dl_fired", None)

                except Exception as exc:
                    loading_ph.empty()
                    st.error(f"정산 중 오류가 발생했습니다:\n\n```\n{exc}\n```")

        # ── 정산 완료 후: 자동 다운로드 트리거 + 수동 다운로드 버튼 ────────────
        result = st.session_state.get("_last_result")
        if result and result.get("line_items"):
            _excel_bytes = result.get("excel_bytes")
            _pdf_bytes   = result.get("pdf_bytes")
            _fname_xlsx  = result.get("excel_filename")
            _fname_pdf   = result.get("pdf_filename")
            _pdf_error   = result.get("pdf_error")

            if _pdf_error:
                st.warning(f"⚠ {_pdf_error}")

            # 자동 다운로드 (이번 결과에 대해 1회만 발사)
            _dl_key = st.session_state.get("_auto_dl_key")
            if _dl_key and st.session_state.get("_auto_dl_fired") != _dl_key:
                _MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                _MIME_PDF  = "application/pdf"

                _anchors = []
                _clicks  = []
                if _excel_bytes:
                    _b64 = base64.b64encode(_excel_bytes).decode()
                    _anchors.append(
                        f'<a id="_dl_xlsx" href="data:{_MIME_XLSX};base64,{_b64}" '
                        f'download="{_fname_xlsx}" style="display:none">excel</a>'
                    )
                    _clicks.append(
                        'setTimeout(function(){var a=document.getElementById("_dl_xlsx");'
                        'if(a)a.click();}, 120);'
                    )
                if _pdf_bytes:
                    _b64 = base64.b64encode(_pdf_bytes).decode()
                    _anchors.append(
                        f'<a id="_dl_pdf" href="data:{_MIME_PDF};base64,{_b64}" '
                        f'download="{_fname_pdf}" style="display:none">pdf</a>'
                    )
                    # Excel 먼저, PDF는 800ms 뒤 (멀티 다운로드 차단 회피)
                    _clicks.append(
                        'setTimeout(function(){var a=document.getElementById("_dl_pdf");'
                        'if(a)a.click();}, 900);'
                    )

                if _anchors:
                    components.html(
                        "<html><body>"
                        + "".join(_anchors)
                        + "<script>" + "".join(_clicks) + "</script>"
                        + "</body></html>",
                        height=0,
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
                # 성공 배너 + balloons
                if run_button:
                    st.balloons()

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

                # DataFrame 구성
                df_result = pd.DataFrame([
                    {
                        "정산월":         it.billing_month,
                        "업체(프로젝트)": it.project_name,
                        "SKU명":          it.sku_name,
                        "총사용량":       it.total_usage,
                        "무료차감":       it.free_cap_applied,
                        "청구대상":       it.billable_usage,
                        "소계(USD)":      float(it.subtotal_usd),
                        "최종(KRW)":      float(it.final_krw),
                    }
                    for it in line_items
                ])

                # KPI 카드
                st.divider()
                st.markdown("#### 📊 정산 요약")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("프로젝트 수",  f"{df_result['업체(프로젝트)'].nunique():,} 개")
                c2.metric("유료 청구 건", f"{(df_result['청구대상'] > 0).sum():,} 건")
                c3.metric("합계 (USD)",   f"$ {df_result['소계(USD)'].sum():,.2f}")
                c4.metric("합계 (KRW)",   f"₩ {df_result['최종(KRW)'].sum():,.0f}")

                st.divider()

                # 프로젝트별 합계
                st.markdown("#### 🏢 프로젝트별 합계")
                df_proj = (
                    df_result
                    .groupby("업체(프로젝트)", as_index=False)
                    .agg(소계_USD=("소계(USD)", "sum"), 최종_KRW=("최종(KRW)", "sum"))
                    .sort_values("최종_KRW", ascending=False)
                    .rename(columns={"소계_USD": "소계(USD)", "최종_KRW": "최종(KRW)"})
                    [["소계(USD)", "최종(KRW)"]]
                )
                df_proj["소계(USD)"] = df_proj["소계(USD)"].map("$ {:,.4f}".format)
                df_proj["최종(KRW)"] = df_proj["최종(KRW)"].map("₩ {:,.0f}".format)
                st.dataframe(df_proj, width='stretch', hide_index=True)

                st.divider()

                # SKU별 세부 내역
                st.markdown("#### 📋 SKU별 세부 내역")
                df_disp = df_result[["SKU명", "총사용량", "무료차감", "청구대상", "소계(USD)", "최종(KRW)"]].copy()
                df_disp["소계(USD)"] = df_disp["소계(USD)"].map("$ {:,.4f}".format)
                df_disp["최종(KRW)"] = df_disp["최종(KRW)"].map("₩ {:,.0f}".format)
                for col in ("총사용량", "무료차감", "청구대상"):
                    df_disp[col] = df_disp[col].map("{:,}".format)
                st.dataframe(df_disp, width='stretch', hide_index=True)



