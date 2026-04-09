"""
webapp.py — SPH GMP 정산 자동화 시스템 v2

streamlit run webapp.py
"""
from __future__ import annotations

import tempfile
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st

from billing.engine import calculate_billing, calculate_billing_by_project
from billing.loader import load_sku_master, load_usage_rows, parse_gmp_price_excel
from billing.preprocessor import extract_company_names, preprocess_usage_file
from excel_formatter import create_report_excel

# ── 경로 상수 ─────────────────────────────────────────────────────────────────
MASTER_CSV = Path(__file__).parent / "billing" / "master_data.csv"

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


def _get_file_preview(tmp_path: Path) -> dict:
    """파일 기본 통계 (전처리 결과 기반, 전체 계정)."""
    try:
        rows = preprocess_usage_file(str(tmp_path), "0000-00")
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

    billing_month = st.text_input(
        "📅 정산월 (YYYY-MM)",
        value="2026-03",
        placeholder="예: 2026-03",
    )
    exchange_rate = st.number_input(
        "💱 환율 (USD → KRW)",
        min_value=500.0, max_value=3000.0,
        value=1350.0, step=1.0, format="%g",
    )
    margin_pct = st.number_input(
        "📈 마진율 (%)",
        min_value=0.0, max_value=100.0,
        value=12.0, step=0.1, format="%g",
        help="원가에 추가할 마진 백분율 (예: 12 → ×1.12 적용)",
    )
    margin_rate = (100.0 + margin_pct) / 100.0

    st.divider()
    sku_count = st.session_state.master_df["sku_id"].nunique()
    st.caption(
        f"적용 환율: {exchange_rate:,.0f} KRW/USD\n"
        f"마진 배수: ×{margin_rate:.4f}\n"
        f"등록 SKU: {sku_count} 종"
    )

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

# ── 탭 UI ─────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["🚀  통합 정산 실행", "⚙️  SKU 단가 관리"])


# ═══════════════════════════════════════════════════════════════════════════════
# Tab 1 ─ 통합 정산 실행
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:

    # ── 파일 업로드 드롭존 ────────────────────────────────────────────────────
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
            st.session_state._companies = extract_company_names(tmp_path)
            st.session_state._preview   = _get_file_preview(tmp_path)
            st.session_state.pop("_last_result", None)  # 이전 결과 초기화

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

        # ── 결제 계정 선택 + 실행 버튼 ───────────────────────────────────────
        st.markdown("#### 정산 대상 선택")
        col_sel, col_gap, col_btn = st.columns([3, 0.15, 1])
        with col_sel:
            if companies:
                selected_company = st.selectbox(
                    "결제 계정 (Billing Account Name)",
                    options=companies,
                    label_visibility="collapsed",
                )
            else:
                selected_company = None
                st.info("파일에서 결제 계정 정보를 찾을 수 없습니다. 전체 데이터를 처리합니다.")

        with col_btn:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            run_button = st.button(
                "▶  정산 시작",
                type="primary",
                use_container_width=True,
            )

        # ── 정산 실행 ─────────────────────────────────────────────────────────
        if run_button:
            sku_rows = _df_to_sku_rows(st.session_state.master_df)
            if not sku_rows:
                st.error("SKU 마스터 데이터가 없습니다. [⚙️ SKU 단가 관리] 탭에서 단가표를 확인하세요.")
            else:
                prog = st.progress(0, text="⚙️ 파일 전처리 중...")
                try:
                    raw_rows = preprocess_usage_file(
                        str(tmp_input_path), billing_month,
                        company_filter=selected_company,
                    )
                    prog.progress(28, text="📦 SKU 마스터 로드 중...")

                    sku_master = load_sku_master(sku_rows)
                    usage_rows = load_usage_rows(raw_rows)

                    prog.progress(52, text="🧮 Waterfall 과금 계산 중...")
                    _ex = Decimal(str(exchange_rate))
                    _mr = Decimal(str(margin_rate))
                    line_items  = calculate_billing(usage_rows, sku_master, _ex, _mr)
                    proj_results = calculate_billing_by_project(usage_rows, sku_master, _ex, _mr)

                    prog.progress(78, text="📊 발송용 엑셀 리포트 생성 중...")
                    excel_bytes = create_report_excel(
                        line_items      = line_items,
                        company_name    = selected_company or "전체",
                        billing_month   = billing_month,
                        exchange_rate   = exchange_rate,
                        margin_rate     = margin_rate,
                        sku_master_rows = sku_rows,
                        proj_results    = proj_results,
                    )

                    prog.progress(100, text="✅ 완료!")
                    prog.empty()

                    st.session_state._last_result = {
                        "line_items":   line_items,
                        "excel_bytes":  excel_bytes,
                        "company":      selected_company,
                    }

                except Exception as exc:
                    prog.empty()
                    st.error(f"정산 중 오류가 발생했습니다:\n\n```\n{exc}\n```")

        # ── 결과 출력 ─────────────────────────────────────────────────────────
        result = st.session_state.get("_last_result")
        if result:
            line_items  = result["line_items"]
            excel_bytes = result["excel_bytes"]
            company_out = result["company"]

            if not line_items:
                st.warning(
                    "처리된 청구 항목이 없습니다. "
                    "SKU 마스터의 SKU ID가 파일의 데이터와 일치하는지 확인하세요."
                )
            else:
                # 성공 배너 + balloons
                if run_button:  # 방금 계산된 경우만 balloons
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
                )
                df_proj["소계(USD)"] = df_proj["소계(USD)"].map("$ {:,.4f}".format)
                df_proj["최종(KRW)"] = df_proj["최종(KRW)"].map("₩ {:,.0f}".format)
                st.dataframe(df_proj, use_container_width=True, hide_index=True)

                st.divider()

                # SKU별 세부 내역
                st.markdown("#### 📋 SKU별 세부 내역")
                df_disp = df_result.copy()
                df_disp["소계(USD)"] = df_disp["소계(USD)"].map("$ {:,.4f}".format)
                df_disp["최종(KRW)"] = df_disp["최종(KRW)"].map("₩ {:,.0f}".format)
                for col in ("총사용량", "무료차감", "청구대상"):
                    df_disp[col] = df_disp[col].map("{:,}".format)
                st.dataframe(df_disp, use_container_width=True, hide_index=True)

                # 다운로드 버튼
                st.divider()
                safe  = (company_out or "전체").replace("/", "_").replace("\\", "_")
                fname = f"정산리포트_{safe}_{billing_month}.xlsx"
                st.download_button(
                    label=f"⬇  {fname}  다운로드",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


# ═══════════════════════════════════════════════════════════════════════════════
# Tab 2 ─ SKU 단가 관리
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("#### ⚙️ SKU 마스터 단가표 관리")
    st.caption(
        "테이블에서 직접 단가를 수정하거나, 표준 양식 CSV를 다운로드해 수정 후 업로드하세요. "
        "저장 후 [통합 정산 실행] 탭의 계산에 즉시 반영됩니다."
    )

    col_cfg = {
        "sku_id":        st.column_config.TextColumn("SKU ID",          width="medium"),
        "sku_name":      st.column_config.TextColumn("SKU명",           width="medium"),
        "is_billable":   st.column_config.CheckboxColumn("과금 여부",   width="small"),
        "category":      st.column_config.SelectboxColumn(
                             "카테고리", width="small",
                             options=["Maps", "Places", "Routes", "Environment", "Tax", "Other"],
                         ),
        "free_usage_cap": st.column_config.NumberColumn(
                              "무료 제공량 (건/월)", min_value=0, format="%d", width="medium",
                          ),
        "tier_number":   st.column_config.NumberColumn(
                             "구간 번호", min_value=1, format="%d", width="small",
                         ),
        "tier_limit":    st.column_config.NumberColumn(
                             "구간 상한 (건)", min_value=0, format="%d", width="medium",
                             help="마지막 구간(무제한)은 비워두세요",
                         ),
        "tier_cpm":      st.column_config.NumberColumn(
                             "단가 (USD / 1,000건)", min_value=0.0,
                             format="$%.4f", width="medium",
                         ),
    }

    edited_df = st.data_editor(
        st.session_state.master_df,
        column_config=col_cfg,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="sku_editor",
    )

    col_save, col_dl, col_ul_label = st.columns([1, 1, 1])

    with col_save:
        if st.button("💾  변경사항 저장", type="primary", use_container_width=True):
            st.session_state.master_df = edited_df.copy()
            _save_master_df(edited_df)
            st.success("✅ 저장 완료 — 정산 계산에 즉시 반영됩니다.")

    with col_dl:
        csv_bytes = edited_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="⬇  현재 단가표 CSV 다운로드",
            data=csv_bytes,
            file_name="master_data_template.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_ul_label:
        uploaded_master = st.file_uploader(
            "📤  단가표 CSV 업로드",
            type=["csv"],
            key="master_uploader",
            label_visibility="collapsed",
            help="표준 양식(위 다운로드)에 맞춰 작성한 CSV를 업로드하세요.",
        )

    if uploaded_master is not None:
        try:
            new_df = pd.read_csv(uploaded_master, dtype=str)
            new_df["is_billable"] = new_df["is_billable"].map(
                {"True": True, "False": False, "true": True, "false": False}
            ).fillna(False)
            new_df["free_usage_cap"] = (
                pd.to_numeric(new_df["free_usage_cap"], errors="coerce").fillna(0).astype(int)
            )
            for col in ("tier_number", "tier_limit", "tier_cpm"):
                new_df[col] = pd.to_numeric(new_df[col], errors="coerce")
            st.session_state.master_df = new_df
            _save_master_df(new_df)
            st.success(f"✅ 단가표 업로드 완료 — {len(new_df)}행 적용됨")
            st.rerun()
        except Exception as exc:
            st.error(f"파일 처리 오류: {exc}")

    st.divider()
    st.markdown("##### 📌 컬럼 설명")
    st.markdown("""
| 컬럼 | 설명 |
|------|------|
| **SKU ID** | 구글 API SKU의 고유 해시 ID (예: `FAF4-3B2D-51B2`) |
| **SKU명** | SKU 표시 이름 (예: `Dynamic Maps`) |
| **과금 여부** | 체크 = 과금 대상, 미체크 = 세금 등 비과금 항목 |
| **카테고리** | Maps / Places / Routes / Tax 등 |
| **무료 제공량** | 월별 무료 제공 건수 (Waterfall 계산 전 선차감) |
| **구간 번호** | Waterfall 구간 순번 (1부터 오름차순, SKU별로 구성) |
| **구간 상한** | 해당 구간의 누적 사용량 상한 — 마지막 구간은 비워두면 무제한 처리 |
| **단가 ($/1K)** | 해당 구간의 1,000건당 USD 단가 |
""")

    # ── GMP 공식 요금표 Excel 업로드 ──────────────────────────────────────
    st.divider()
    with st.expander("📥  GMP 공식 요금표 Excel로 단가표 자동 생성", expanded=False):
        st.caption(
            "구글 Maps Platform 공식 요금표 Excel(.xlsx)을 업로드하면 "
            "단가표를 자동으로 파싱합니다. "
            "파싱 후 SKU ID를 실제 청구 ID로 수정하고 저장하세요."
        )
        uploaded_price_xl = st.file_uploader(
            "GMP 요금표 Excel 업로드",
            type=["xlsx"],
            key="price_excel_uploader",
            label_visibility="collapsed",
        )

        if uploaded_price_xl is not None:
            try:
                with st.spinner("요금표 파싱 중..."):
                    parsed_rows = parse_gmp_price_excel(uploaded_price_xl.read())

                if not parsed_rows:
                    st.warning(
                        "파싱된 데이터가 없습니다. "
                        "파일 형식을 확인하세요 (A열에 'Monthly volume range' 헤더가 있어야 합니다)."
                    )
                else:
                    parsed_df = pd.DataFrame(parsed_rows)

                    # 컬럼 타입 정규화
                    parsed_df["is_billable"] = parsed_df["is_billable"].astype(bool)
                    parsed_df["free_usage_cap"] = (
                        pd.to_numeric(parsed_df["free_usage_cap"], errors="coerce")
                        .fillna(0).astype(int)
                    )
                    for col in ("tier_number", "tier_limit", "tier_cpm"):
                        parsed_df[col] = pd.to_numeric(parsed_df[col], errors="coerce")

                    n_skus = parsed_df["sku_id"].nunique()
                    n_rows_parsed = len(parsed_df)

                    st.success(
                        f"✅ 파싱 완료 — SKU {n_skus}종 / 구간 행 {n_rows_parsed}개"
                    )
                    st.info(
                        "⚠️ SKU ID는 API명 기반 합성 ID입니다. "
                        "실제 구글 청구 데이터의 SKU ID(예: `FAF4-3B2D-51B2`)와 "
                        "일치하도록 적용 후 직접 수정하세요."
                    )

                    st.dataframe(
                        parsed_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "free_usage_cap": st.column_config.NumberColumn(
                                "무료 제공량", format="%d"
                            ),
                            "tier_cpm": st.column_config.NumberColumn(
                                "단가 ($/1K)", format="$%.4f"
                            ),
                        },
                    )

                    if st.button(
                        "⬆  위 파싱 결과를 단가표에 적용 (기존 데이터 덮어쓰기)",
                        type="primary",
                        key="apply_parsed_excel",
                    ):
                        st.session_state.master_df = parsed_df.copy()
                        _save_master_df(parsed_df)
                        st.success("✅ 단가표 적용 완료 — 위 편집 테이블을 확인하고 SKU ID를 수정하세요.")
                        st.rerun()

            except Exception as exc:
                st.error(f"요금표 파싱 오류:\n\n```\n{exc}\n```")
