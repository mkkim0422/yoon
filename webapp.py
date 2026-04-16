"""
webapp.py — SPH GMP 정산 자동화 시스템 v2

streamlit run webapp.py
"""
from __future__ import annotations

import base64
import hashlib
import json
import tempfile
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from streamlit_sortables import sort_items

from billing.engine import calculate_billing, calculate_billing_by_project
from billing.loader import (
    detect_price_list_currency,
    load_sku_master,
    load_usage_rows,
)
from billing.preprocessor import extract_company_names, preprocess_usage_file
from invoice_generator import generate_formatted_invoice

# ── 경로 상수 ─────────────────────────────────────────────────────────────────
MASTER_CSV        = Path(__file__).parent / "billing" / "master_data.csv"
PRICE_LIST_SAVED  = Path(__file__).parent / "billing" / "saved_price_list.xlsx"
SAVED_ORDERS_FILE = Path(__file__).parent / "billing" / "saved_orders.json"


# ── SKU 순서 저장/로드 (계정별) ─────────────────────────────────────────────
def _load_saved_orders() -> dict[str, list[str]]:
    if not SAVED_ORDERS_FILE.exists():
        return {}
    try:
        return json.loads(SAVED_ORDERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_order_for_account(account: str, order: list[str]) -> None:
    data = _load_saved_orders()
    data[account] = order
    SAVED_ORDERS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SAVED_ORDERS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── tax/VAT SKU 판별 (인보이스 본문에서 제외되는 항목) ─────────────────────
_TAX_KW = ("세금", "tax", "vat")
def _is_tax_sku(name: str) -> bool:
    n = (name or "").lower()
    return any(kw in n for kw in _TAX_KW)


def _unique_skus_for_account(tmp_path: str, billing_month: str,
                              account: str | None) -> list[str]:
    """선택된 계정의 CSV에서 **실제 Invoice에 출력될** SKU명 리스트.
    engine 필터와 동일 기준: 사용량 > 0 또는 cost_krw ≠ 0 이 있는 SKU만 (tax 제외).
    """
    rows = _cached_preprocess(tmp_path, billing_month, account)
    # sku_name → "has any usage or cost" 플래그 집계
    nonzero: dict[str, bool] = {}
    for r in rows:
        nm = (r.get("sku_name") or "").strip()
        if not nm or _is_tax_sku(nm):
            continue
        if nonzero.get(nm):
            continue
        usage = int(r.get("usage_amount") or 0)
        cost  = float(r.get("cost_krw") or 0)
        if usage > 0 or cost != 0:
            nonzero[nm] = True
        else:
            nonzero.setdefault(nm, False)
    # True인 것만 반환, 알파벳순
    return sorted(nm for nm, has in nonzero.items() if has)

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
</style>
""", unsafe_allow_html=True)


# ── 헬퍼 함수 ─────────────────────────────────────────────────────────────────

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
@st.cache_data(show_spinner=False)
def _cached_preprocess(tmp_path: str, billing_month: str,
                       company_filter: str | None) -> list[dict]:
    return preprocess_usage_file(
        tmp_path, billing_month, company_filter=company_filter
    )


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

    # ── 단가표 통화 선택 (달러 / 원화) ──────────────────────────────────────
    # 단가표 업로드 시 detect_price_list_currency() 결과가 기본값으로 세팅되며
    # 사용자는 라디오로 수동 교정 가능.
    #   - 달러: 기존 로직 (환율 입력 필요, Invoice에 $ 단가/합계, Project에 toal($))
    #   - 원화: 환율 미사용, 모든 금액 ₩ 표기, 합계/환율 행 생략
    _detected = st.session_state.get("_detected_currency", "USD")
    _cur_options = ["달러 ($)", "원화 (₩)"]
    _cur_idx     = 0 if _detected == "USD" else 1
    currency_label = st.radio(
        "💰 단가표 통화",
        options=_cur_options,
        index=_cur_idx,
        horizontal=True,
        key="_currency_radio",
        help="단가표가 USD 기준이면 '달러', KRW 기준이면 '원화'를 선택하세요. "
             "단가표 업로드 시 자동 감지됩니다.",
    )
    currency = "USD" if currency_label.startswith("달러") else "KRW"

    # ── 환율: 달러 모드에서만 입력 가능 (원화 모드에선 사용 안 함) ───────────
    # 포맷 규칙 (JS):
    #   - 숫자만 남긴 뒤 앞 4자리=정수부 / 다음 2자리=소수부
    #   - 입력 중 4자리 초과 시 자동으로 "." 삽입
    #   - 포커스 아웃 시 소수부 2자리 0 패딩 ("1427" → "1427.00")
    _rate_raw = st.text_input(
        "💱 환율 (USD → KRW)",
        value=st.session_state.get("_rate_raw", ""),
        max_chars=7,
        placeholder="예: 1427.87" if currency == "USD" else "원화 모드 — 입력 불필요",
        key="_rate_raw_input",
        disabled=(currency == "KRW"),
    )
    # 숫자 파싱 — 유효한 수치만 사용 (원화 모드에선 1.0 으로 고정)
    if currency == "KRW":
        exchange_rate = 1.0   # placeholder — 원화 모드에선 실제로 사용 안 함
    else:
        try:
            exchange_rate = float(_rate_raw) if _rate_raw else 0.0
        except ValueError:
            exchange_rate = 0.0

    bank_name = st.text_input(
        "🏦 은행",
        value="하나은행",
    )
    margin_rate = 1.0

    st.divider()
    sku_count = st.session_state.master_df["sku_id"].nunique()
    if currency == "KRW":
        _rate_disp = "— 원화 모드 (환율 미사용)"
    else:
        _rate_disp = f"{exchange_rate:,.2f} KRW/USD" if exchange_rate > 0 else "— 입력 필요"
    st.caption(
        f"단가표 통화: **{currency}**\n"
        f"적용 환율: {_rate_disp}\n"
        f"등록 SKU: {sku_count} 종"
    )

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

# ── 설정 변경 감지 → 이전 결과 자동 무효화 ────────────────────────────────────
_current_calc_key = f"{billing_month}|{exchange_rate}|{bank_name}"
if st.session_state.get("_calc_key") != _current_calc_key:
    st.session_state["_calc_key"] = _current_calc_key
    st.session_state.pop("_last_result", None)

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

    # ── ① 단가표(GMP Price List) 첨부 — 저장 유지 ───────────────────────────
    _has_saved = PRICE_LIST_SAVED.exists()

    with st.container():
        _col_up, _col_del = st.columns([5, 1], vertical_alignment="bottom")
        with _col_up:
            _uploaded_price = st.file_uploader(
                "📋 단가표(GMP Price List) 첨부",
                type=["xlsx", "xls"],
                key="price_list_uploader",
                help="한 번 업로드하면 저장되어 새로 올리기 전까지 자동 사용됩니다.",
            )
        with _col_del:
            if _has_saved:
                if st.button("🗑 단가표 삭제", key="del_price_list",
                             help="저장된 단가표를 삭제합니다"):
                    PRICE_LIST_SAVED.unlink(missing_ok=True)
                    st.session_state.pop("_saved_price_key", None)
                    st.session_state.pop("_detected_currency", None)
                    st.success("단가표가 삭제되었습니다.")
                    st.rerun()

    # 새 파일이 올라왔으면 디스크에 저장 (중복 저장 방지 + rerun으로 상태 갱신)
    if _uploaded_price is not None:
        _price_key = f"{_uploaded_price.name}_{_uploaded_price.size}"
        if st.session_state.get("_saved_price_key") != _price_key:
            _uploaded_price.seek(0)
            PRICE_LIST_SAVED.parent.mkdir(parents=True, exist_ok=True)
            PRICE_LIST_SAVED.write_bytes(_uploaded_price.read())
            st.session_state["_saved_price_key"] = _price_key
            st.session_state["_price_flash"] = _uploaded_price.name
            # 단가표 통화(USD/KRW) 자동 감지 → 사이드바 라디오 기본값에 반영
            try:
                _detected_cur = detect_price_list_currency(PRICE_LIST_SAVED)
                st.session_state["_detected_currency"] = _detected_cur
            except Exception:
                pass
            st.rerun()
        _has_saved = True

    # 저장 완료 flash 메시지
    if _flash := st.session_state.pop("_price_flash", None):
        st.success(f"✅ 단가표 저장 완료 — {_flash}  (이후 정산에서 자동 사용)")

    # 실제로 사용할 price_list_file 결정 (업로드 > 저장파일 > None)
    if _uploaded_price is not None:
        _uploaded_price.seek(0)
        price_list_file = _uploaded_price
    elif _has_saved:
        price_list_file = PRICE_LIST_SAVED
        # 저장파일 기준으로도 통화 재확인 (세션 초기 로드 시)
        if "_detected_currency" not in st.session_state:
            try:
                st.session_state["_detected_currency"] = (
                    detect_price_list_currency(PRICE_LIST_SAVED)
                )
            except Exception:
                st.session_state["_detected_currency"] = "USD"
        _cur_tag = st.session_state.get("_detected_currency", "USD")
        st.info(
            f"📂 저장된 단가표 사용 중 — **{PRICE_LIST_SAVED.name}** "
            f"(감지 통화: **{_cur_tag}**) *(우측 🗑 버튼으로 삭제)*"
        )
    else:
        price_list_file = None
        st.caption("단가표 미첨부 — 업로드 시 3번째 시트로 이식됩니다.")

    # ── ② 사용고지서 업로드 ───────────────────────────────────────────────────
    uploaded_file = st.file_uploader(
        "📂 구글 Maps 플랫폼 사용고지서 파일을 여기에 드래그 앤 드롭하세요",
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
            selected_company = st.selectbox(
                "결제 계정 (Billing Account Name)",
                options=companies,
            )
        else:
            selected_company = None
            st.info("파일에서 결제 계정 정보를 찾을 수 없습니다. 전체 데이터를 처리합니다.")

        # ── 좌: 드래그앤드롭 | 우: 다운로드 체크박스 + 정산 시작 ────────────
        _order_account_key = selected_company or "__ALL__"
        _found_skus = _unique_skus_for_account(
            str(tmp_input_path), billing_month, selected_company
        )

        col_left, col_right = st.columns([1, 1], vertical_alignment="top", gap="large")

        # ═══ 좌측: SKU 순서 드래그앤드롭 ═══════════════════════════════════
        with col_left:
            if _found_skus:
                _saved_orders = _load_saved_orders()
                _saved_for_this = _saved_orders.get(_order_account_key, [])

                # 저장된 순서 + 이번 CSV에 신규로 등장한 항목 분리
                _existing   = [n for n in _saved_for_this if n in _found_skus]
                _new_items  = [n for n in _found_skus if n not in _existing]
                _NEW_COUNT  = len(_new_items)

                st.markdown("#### 📋 엑셀 SKU 노출 순서")
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
                    .sortable-item:nth-last-child(-n+{_NEW_COUNT}) {{
                        background: linear-gradient(135deg, #f2fae3 0%, #e4f2cd 100%);
                        border-color: #cfe7a8;
                        border-left-color: #7dbb26;
                        color: #1b3d06;
                    }}
                    """

                _SORT_CSS = """
                /* (a) 컴포넌트 내부 모든 전이/애니메이션 차단 — iframe 리플로우 원천 차단 */
                .sortable-component,
                .sortable-component *,
                .sortable-item,
                .sortable-item * {
                    transition: none !important;
                    animation: none !important;
                }
                /* (b) 레이아웃 격리 */
                .sortable-component {
                    background: linear-gradient(135deg,#f8fcfd 0%,#edf6f8 100%);
                    border: 2px dashed #00788a;
                    border-radius: 14px;
                    padding: 14px;
                    gap: 8px;
                    contain: layout style;
                }
                .sortable-item {
                    background: #ffffff;
                    color: #1a3540;
                    border: 1.5px solid #bdd8de;
                    border-left: 4px solid #00788a;
                    border-radius: 10px;
                    padding: 10px 14px;
                    font-weight: 600;
                    font-size: 0.92rem;
                    cursor: grab;
                    user-select: none;
                    /* (c) 고정 높이 + 박스 모델 안정화 */
                    min-height: 42px;
                    box-sizing: border-box;
                    width: 100%;
                    /* (b) 자식 변경 격리 */
                    contain: layout style paint;
                }
                .sortable-item:active { cursor: grabbing; }
                /* (d) :hover 상태 변경 없음 — 의도적으로 비움 */
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
            else:
                sku_order = []
                st.info("표시할 SKU가 없습니다. CSV를 확인하세요.")

        # ═══ 우측: 다운로드 옵션 + 정산 시작 ═══════════════════════════════
        with col_right:
            st.markdown("#### ⚙️ 다운로드 옵션")
            st.markdown(
                """
                <div style="
                    background: linear-gradient(135deg,#f8fcfd 0%,#edf6f8 100%);
                    border: 1.5px solid #bdd8de;
                    border-radius: 14px;
                    padding: 18px 22px;
                    margin-bottom: 14px;
                ">
                """,
                unsafe_allow_html=True,
            )
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
            st.markdown("</div>", unsafe_allow_html=True)

            # 버튼은 항상 클릭 가능 — 검증은 클릭 시점에 얼럿/토스트로 처리
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
            sku_rows = _df_to_sku_rows(st.session_state.master_df)
            if not sku_rows:
                st.error("SKU 마스터 데이터가 없습니다. billing/master_data.csv 를 확인하세요.")
            else:
                prog = st.progress(0, text="⚙️ 파일 전처리 중...")
                try:
                    raw_rows = _cached_preprocess(
                        str(tmp_input_path), billing_month, selected_company
                    )
                    prog.progress(20, text="📦 SKU 마스터 로드 중...")

                    sku_master = load_sku_master(sku_rows)
                    usage_rows = load_usage_rows(raw_rows)

                    prog.progress(35, text="🧮 Waterfall 과금 계산 중...")
                    _ex = Decimal(str(exchange_rate))
                    _mr = Decimal(str(margin_rate))
                    line_items   = calculate_billing(usage_rows, sku_master, _ex, _mr)
                    proj_results = calculate_billing_by_project(usage_rows, sku_master, _ex, _mr)

                    prog.progress(55, text="📄 Excel 인보이스 생성 중...")
                    _safe        = (selected_company or "전체").replace("/", "_").replace("\\", "_")
                    _fname_xlsx  = f"정산리포트_{_safe}_{billing_month}.xlsx"
                    _fname_pdf   = f"정산리포트_{_safe}_{billing_month}.pdf"
                    _excel_bytes = generate_formatted_invoice(
                        line_items      = line_items,
                        company_name    = selected_company or "전체",
                        billing_month   = billing_month,
                        exchange_rate   = _ex,
                        margin_rate     = _mr,
                        bank_name       = bank_name,
                        proj_results    = proj_results,
                        price_list_file = price_list_file,
                        sku_order       = sku_order or None,
                        currency        = currency,
                    )

                    # PDF 변환 (체크된 경우만)
                    _pdf_bytes = None
                    _pdf_error = None
                    if dl_pdf:
                        prog.progress(75, text="📄 PDF 변환 중 (Excel 실행)...")
                        from pdf_export import xlsx_sheet_to_pdf
                        _pdf_bytes, _pdf_error = xlsx_sheet_to_pdf(
                            _excel_bytes, "Invoice"
                        )

                    prog.progress(100, text="✅ 완료!")
                    prog.empty()

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
                    prog.empty()
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



