import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats
import json
import os
from pathlib import Path
from datetime import datetime, timedelta, date

st.set_page_config(page_title="루저스클럽 ERP", page_icon="🏷️", layout="wide", initial_sidebar_state="expanded")

# 비밀번호 인증
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    st.title("🔐 루저스클럽 ERP")
    pw = st.text_input("비밀번호를 입력하세요", type="password")
    if st.button("로그인"):
        if pw == st.secrets.get("password", "teddy"):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False

if not check_password():
    st.stop()

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR.parent / "매출자료"
MGMT_FILE = BASE_DIR / "management.json"
MKTG_FILE = BASE_DIR / "루저스클럽 마케팅 효율.xlsx"

CHANNELS = ['무신사', '자사몰', '무신사글로벌', '29CM', '팝업(더현대)']
CH_COLORS = {'무신사':'#1E88E5','자사몰':'#43A047','무신사글로벌':'#FB8C00','29CM':'#E53935','팝업(더현대)':'#8E24AA'}
CAT_MAP = {'PT':'팬츠','TS':'반팔티','BG':'가방','HZ':'후리스집업','SW':'스웻/후드','JK':'자켓',
           'JW':'주얼리','CP':'캡/비니','SP':'스웻팬츠','DP':'데님팬츠','KR':'크루넥','SL':'슬리브리스','AC':'악세서리'}

st.markdown("""
<style>
/* ── 전체 여백 축소 ── */
.block-container{padding-top:1rem!important;padding-bottom:1rem!important;}
section[data-testid="stSidebar"] .block-container{padding-top:1rem!important;}

/* ── KPI 카드 ── */
.kpi{background:#fff;border-radius:10px;padding:12px 16px;border:1px solid #e8e8e8;
     border-top:3px solid #1E88E5;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:6px;}
.kpi-label{font-size:11px;color:#888;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px;}
.kpi-value{font-size:22px;font-weight:800;color:#111;line-height:1.1;}
.kpi-delta{font-size:12px;margin-top:3px;font-weight:600;}
.up{color:#2e7d32;}.dn{color:#c62828;}

/* ── 랭킹 행 ── */
.rk{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:8px;
    margin-bottom:3px;background:#fafafa;border:1px solid #f0f0f0;transition:background .15s;}
.rk:hover{background:#f0f4ff;}
.rk-medal{font-size:15px;width:28px;text-align:center;flex-shrink:0;}
.rk-name{flex:1;font-size:13px;color:#333;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.rk-val{font-size:13px;font-weight:700;color:#1E88E5;white-space:nowrap;}
.rk-delta{font-size:11px;font-weight:700;white-space:nowrap;min-width:44px;text-align:right;}

/* ── 채널 행 ── */
.ch-row{display:flex;align-items:center;gap:8px;padding:8px 12px;border-radius:8px;
        margin-bottom:4px;background:#fff;border:1px solid #eee;}
.ch-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0;}
.ch-name{flex:1;font-size:13px;color:#444;}
.ch-val{font-size:14px;font-weight:700;color:#111;}

/* ── 정보 카드 ── */
.warn-card{background:#fff8f0;border-radius:8px;padding:9px 13px;border-left:3px solid #FB8C00;margin:3px 0;font-size:13px;}
.danger-card{background:#fff5f5;border-radius:8px;padding:9px 13px;border-left:3px solid #E53935;margin:3px 0;font-size:13px;}
.good-card{background:#f1f8f1;border-radius:8px;padding:9px 13px;border-left:3px solid #43A047;margin:3px 0;font-size:13px;}
.insight-card{background:#f0f6ff;border-radius:8px;padding:10px 14px;border-left:3px solid #1E88E5;margin:4px 0;font-size:13px;}

/* ── 섹션 제목 ── */
.sec-title{font-size:14px;font-weight:700;color:#444;margin:0 0 8px 0;padding-bottom:5px;
           border-bottom:2px solid #f0f0f0;letter-spacing:.3px;}

/* ── metric-card (기존 호환) ── */
.metric-card{background:#fff;border-radius:10px;padding:12px 16px;border-left:4px solid #1E88E5;
             border:1px solid #eee;border-left:4px solid #1E88E5;margin-bottom:6px;font-size:13px;}
.metric-label{font-size:11px;color:#888;font-weight:600;}
.metric-value{font-size:20px;font-weight:800;color:#111;}
.metric-delta{font-size:12px;margin-top:2px;font-weight:600;}

/* ── Streamlit 상단 툴바 숨김 ── */
header[data-testid="stHeader"]{display:none!important;}
#MainMenu{display:none!important;}
footer{display:none!important;}

/* ── 탭 바 상단 고정 ── */
[data-baseweb="tab-list"]{
    position:sticky!important;top:0!important;z-index:999!important;
    background:#fff!important;padding:4px 0!important;
    box-shadow:0 2px 6px rgba(0,0,0,.07);}
</style>""", unsafe_allow_html=True)

# ── 데이터 로더 ───────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data(filepath: str):
    df = pd.read_excel(filepath)
    df = df.rename(columns={'판매처':'채널','판매일자':'날짜'})
    df['날짜'] = pd.to_datetime(df['날짜'])
    df['년월'] = df['날짜'].dt.to_period('M').astype(str)
    df['주차'] = df['날짜'].dt.to_period('W').astype(str)
    df['년'] = df['날짜'].dt.year
    df['월'] = df['날짜'].dt.month
    ch_map = {'자사몰':'자사몰','무신사':'무신사','무신사글로벌':'무신사글로벌','29CM':'29CM','더현대':'팝업(더현대)'}
    df['채널'] = df['채널'].map(ch_map).fillna(df['채널'])
    df['상품코드'] = df['상품코드'].astype(str)
    df['스타일코드'] = df['상품코드'].str.rsplit('_', n=1).str[0]

    def get_cat(code):
        if pd.isna(code): return '기타'
        code = str(code)
        for k, v in CAT_MAP.items():
            if k in code: return v
        return '기타'
    df['카테고리'] = df['스타일코드'].apply(get_cat)
    df['판매단가'] = np.where(df['수량'] > 0, df['판매금액'] / df['수량'], np.nan)

    # 실현율 = 판매금액 / (택가 × 수량). 채널 보정 없음 (무신사글로벌은 환율 프리미엄이 이미 판매금액에 반영)
    df['총택가'] = df['택가'] * df['수량']
    df['실현율'] = np.where(
        (df['총택가'] > 0) & df['택가'].notna(),
        (df['판매금액'] / df['총택가']) * 100,
        np.nan
    )
    # 이상값 제거 (30% 미만 or 150% 초과는 오류 데이터)
    df['실현율'] = df['실현율'].where((df['실현율'] >= 30) & (df['실현율'] <= 150))
    df['할인율'] = 100 - df['실현율']

    non29 = df[df['채널'] != '29CM']
    def get_kor_name(x):
        kor = x[~x.str.contains(r'^Lsc |^LSC ', na=False, regex=True)]
        return kor.mode()[0] if len(kor) > 0 else x.mode()[0]
    name_map = non29.groupby('스타일코드')['상품명'].apply(get_kor_name).to_dict()
    df['대표명'] = df['스타일코드'].map(name_map).fillna(df['상품명'])

    style_rev = df.groupby('스타일코드')['판매금액'].sum().sort_values(ascending=False)
    top20 = style_rev.head(20).index.tolist()
    df['상품구분'] = df['스타일코드'].apply(lambda x: '주력' if x in top20 else '기타')
    return df

def load_mgmt():
    if MGMT_FILE.exists():
        with open(MGMT_FILE, 'r', encoding='utf-8') as f: return json.load(f)
    return {"targets":{},"memos":{},"todos":[],"stock_notes":{}}

def save_mgmt(data):
    with open(MGMT_FILE, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False, indent=2)

def load_goals():
    return load_mgmt().get("targets", {})

def save_goals(goals):
    d = load_mgmt(); d["targets"] = goals; save_mgmt(d)

def load_memos():
    return load_mgmt().get("memos", {})

def save_memos(memos):
    d = load_mgmt(); d["memos"] = memos; save_mgmt(d)

def load_todos():
    return load_mgmt().get("todos", [])

def save_todos(todos):
    d = load_mgmt(); d["todos"] = todos; save_mgmt(d)

def kpi_card(col, label, value, delta=None, color="#1E88E5"):
    with col:
        delta_html = ""
        if delta is not None and not (isinstance(delta, float) and np.isnan(delta)):
            arrow = "▲" if delta > 0 else "▼"
            cls = "up" if delta > 0 else "dn"
            delta_html = f'<div class="kpi-delta {cls}">{arrow} {abs(delta):.1f}%</div>'
        st.markdown(
            f'<div class="kpi" style="border-top-color:{color}">'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            f'{delta_html}</div>',
            unsafe_allow_html=True)

# ── 사이드바 ──────────────────────────────────────────────────────────────────
st.sidebar.markdown("## 🏷️ 루저스클럽 ERP")
st.sidebar.markdown("---")

excel_files = sorted(DATA_DIR.glob("*.xlsx"), key=os.path.getmtime, reverse=True)
if not excel_files:
    st.error(f"매출자료 폴더에 엑셀 파일이 없습니다: {DATA_DIR}")
    st.stop()

selected_file = st.sidebar.selectbox("📂 매출 데이터", [f.name for f in excel_files])
filepath = str(DATA_DIR / selected_file)
df_all = load_data(filepath)

months_all = sorted(df_all['년월'].unique())
c1, c2 = st.sidebar.columns(2)
start_m = c1.selectbox("시작", months_all, index=0)
end_m = c2.selectbox("종료", months_all, index=len(months_all)-1)
selected_ch = st.sidebar.multiselect("채널", CHANNELS, default=CHANNELS)

mask = (df_all['년월'] >= start_m) & (df_all['년월'] <= end_m)
if selected_ch: mask = mask & df_all['채널'].isin(selected_ch)
df = df_all[mask].copy()

st.sidebar.markdown("---")
st.sidebar.caption(f"📅 {df_all['날짜'].min().strftime('%Y.%m.%d')} ~ {df_all['날짜'].max().strftime('%Y.%m.%d')}")
st.sidebar.caption(f"총 {len(df_all):,}건 | 필터 {len(df):,}건")

st.sidebar.markdown("---")
if st.sidebar.button("📥 매출 업데이트", use_container_width=True, type="primary"):
    import subprocess, sys
    merge_script = Path(__file__).parent / "merge_raw.py"
    if not merge_script.exists():
        st.sidebar.error("merge_raw.py 없음")
    else:
        with st.spinner("업데이트 중..."):
            result = subprocess.run(
                [sys.executable, str(merge_script)],
                capture_output=True, text=True, encoding='utf-8', errors='replace'
            )
        if result.returncode == 0:
            st.sidebar.success("✅ 완료!")
            st.cache_data.clear()
            st.rerun()
        else:
            st.sidebar.error("❌ 오류")
            st.sidebar.code(result.stderr or result.stdout)

# ── 마케팅 데이터 로더 ───────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_marketing(filepath: str):
    if not Path(filepath).exists():
        return None
    xf = pd.ExcelFile(filepath)
    result = {}
    df = pd.read_excel(xf, sheet_name='무신사 메타 광고', header=1)
    df = df.rename(columns={'Unnamed: 0':'_'})
    df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
    df = df.dropna(subset=['날짜'])
    df['년월'] = df['날짜'].dt.to_period('M').astype(str)
    df['주차'] = df['날짜'].dt.to_period('W').astype(str)
    df['ROAS'] = pd.to_numeric(df['catalog_segment_value_omni_purchase_roas:omni_purchase'], errors='coerce')
    df['지출 금액 (KRW)'] = pd.to_numeric(df['지출 금액 (KRW)'], errors='coerce').fillna(0)
    df['링크 클릭'] = pd.to_numeric(df['링크 클릭'], errors='coerce').fillna(0)
    df['노출'] = pd.to_numeric(df['노출'], errors='coerce').fillna(0)
    df['공유 항목이 포함된 구매'] = pd.to_numeric(df['공유 항목이 포함된 구매'], errors='coerce').fillna(0)
    df['공유 항목의 구매 전환값'] = pd.to_numeric(df['공유 항목의 구매 전환값'], errors='coerce').fillna(0)
    result['meta_ms'] = df

    df2 = pd.read_excel(xf, sheet_name='자사몰 메타 광고', header=1)
    df2['날짜'] = pd.to_datetime(df2['날짜'], errors='coerce')
    df2 = df2.dropna(subset=['날짜'])
    df2['년월'] = df2['날짜'].dt.to_period('M').astype(str)
    df2['주차'] = df2['날짜'].dt.to_period('W').astype(str)
    for col in ['지출 금액 (KRW)','링크 클릭','노출','구매','구매 전환값']:
        df2[col] = pd.to_numeric(df2[col], errors='coerce').fillna(0)
    result['meta_js'] = df2

    df3 = pd.read_excel(xf, sheet_name='무신사 상품광고', header=4)
    df3 = df3.dropna(subset=['날짜'])
    df3['날짜'] = pd.to_datetime(df3['날짜'], errors='coerce')
    df3['년월'] = df3['날짜'].dt.to_period('M').astype(str)
    df3['주차'] = df3['날짜'].dt.to_period('W').astype(str)
    for col in ['집행 광고비','매출','광고 수익률(ROAS)','클릭률','전환율','노출 수','클릭 수']:
        df3[col] = pd.to_numeric(df3[col], errors='coerce').fillna(0)
    result['prod_ms'] = df3

    df4 = pd.read_excel(xf, sheet_name='29CM 상품광고', header=4)
    df4 = df4.dropna(subset=['날짜'])
    df4['날짜'] = pd.to_datetime(df4['날짜'], errors='coerce')
    df4['년월'] = df4['날짜'].dt.to_period('M').astype(str)
    df4['주차'] = df4['날짜'].dt.to_period('W').astype(str)
    for col in ['집행 광고비','매출','광고 수익률(ROAS)','클릭률','전환율','노출 수','클릭 수']:
        df4[col] = pd.to_numeric(df4[col], errors='coerce').fillna(0)
    result['prod_29'] = df4
    return result

# ── 탭 (7탭) ──────────────────────────────────────────────────────────────────
tabs = st.tabs(["📊 홈", "🏷️ 카테고리", "📺 채널", "🏆 상품랭킹",
                "🔬 상품분석", "🏪 팝업", "📣 마케팅", "💡 인사이트", "⚙️ 관리"])
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = tabs

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — 홈
# ════════════════════════════════════════════════════════════════════════════
with tab1:
    mgmt_home = load_mgmt()
    targets = mgmt_home.get("targets", {})

    # 월별 KPI 계산
    monthly = df.groupby('년월').agg(매출=('판매금액','sum'),수량=('수량','sum'),SKU=('상품명','nunique')).reset_index()
    monthly['ASP'] = monthly['매출'] / monthly['수량']
    # 가중 실현율 = 월 전체 판매금액 / 월 전체 총택가
    _real_w = df.groupby('년월').apply(
        lambda x: x['판매금액'].sum() / x['총택가'].sum() * 100 if x['총택가'].sum() > 0 else np.nan
    ).reset_index(name='실현율_w')
    monthly = monthly.merge(_real_w, on='년월', how='left')
    monthly['실현율'] = monthly['실현율_w']
    monthly.drop(columns=['실현율_w'], inplace=True)
    monthly['매출_만'] = (monthly['매출']/10000).round()
    monthly = monthly.sort_values('년월').reset_index(drop=True)
    monthly['MoM'] = monthly['매출'].pct_change()*100
    ym_map = dict(zip(monthly['년월'], monthly['매출']))
    def get_yoy(ym):
        prev = str(pd.Period(ym,'M')-12)
        return (ym_map[ym]-ym_map[prev])/ym_map[prev]*100 if prev in ym_map else None
    monthly['YoY'] = monthly['년월'].apply(get_yoy)

    last = monthly.iloc[-1]
    prev_row = monthly.iloc[-2] if len(monthly) > 1 else None
    last_ym = last['년월']
    actual_rev = last['매출']
    target_val = targets.get(last_ym, 0) * 10000  # 만원 → 원

    # 예상 달성율: 오늘 기준 월 진행률로 추정
    today = date.today()
    try:
        days_in_month = pd.Period(last_ym, 'M').days_in_month
        days_passed = min(today.day, days_in_month)
        progress_rate = days_passed / days_in_month
    except:
        progress_rate = 1.0
    expected_rev = actual_rev / progress_rate if progress_rate > 0 else actual_rev
    ach_rate = actual_rev / target_val * 100 if target_val > 0 else None
    expected_ach = expected_rev / target_val * 100 if target_val > 0 else None
    gap_man = (actual_rev - target_val) / 10000 if target_val > 0 else None

    # 전년 동월
    prev_ym = str(pd.Period(last_ym, 'M') - 12)
    all_ym_map = dict(zip(df_all.groupby('년월')['판매금액'].sum().index,
                          df_all.groupby('년월')['판매금액'].sum().values))
    prev_rev = all_ym_map.get(prev_ym, 0)
    yoy_val = (actual_rev - prev_rev) / prev_rev * 100 if prev_rev > 0 else None

    # ── 목표 달성 현황 ──────────────────────────────────────────────────────
    st.markdown(f'<div class="sec-title">📊 {last_ym} 목표 달성 현황</div>', unsafe_allow_html=True)
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    kpi_card(c1, "이번달 매출", f"{actual_rev/10000:,.0f}만", last['MoM'])
    if target_val > 0:
        kpi_card(c2, "목표 매출", f"{target_val/10000:,.0f}만", color="#757575")
        ach_color = "#43A047" if (ach_rate or 0) >= 100 else ("#FB8C00" if (ach_rate or 0) >= 80 else "#E53935")
        kpi_card(c3, "달성율", f"{ach_rate:.1f}%" if ach_rate else "-", color=ach_color)
        gap_color = "#43A047" if (gap_man or 0) >= 0 else "#E53935"
        kpi_card(c4, "목표 대비 갭", f"{gap_man:+,.0f}만" if gap_man is not None else "-", color=gap_color)
    else:
        kpi_card(c2, "목표 미설정", "관리탭에서 입력", color="#757575")
        kpi_card(c3, "달성율", "-", color="#757575")
        kpi_card(c4, "갭", "-", color="#757575")
    kpi_card(c5, f"전년 대비 ({prev_ym})", f"{yoy_val:+.1f}%" if yoy_val is not None else "데이터 없음",
             color="#1E88E5" if (yoy_val or 0) >= 0 else "#E53935")
    kpi_card(c6, "예상 달성율", f"{expected_ach:.1f}%" if expected_ach else "-",
             color="#43A047" if (expected_ach or 0) >= 100 else "#FB8C00")

    # ── 이번주 베스트 10 & 채널별 전주대비 ─────────────────────────────────
    last_date = df_all['날짜'].max()
    week_start = last_date - timedelta(days=6)
    prev_week_start = week_start - timedelta(days=7)

    this_week_df = df_all[df_all['날짜'] >= week_start]
    prev_week_df = df_all[(df_all['날짜'] >= prev_week_start) & (df_all['날짜'] < week_start)]

    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.markdown(f'<div class="sec-title">🏆 이번주 베스트 10 &nbsp;<span style="font-weight:400;color:#aaa;font-size:11px">{week_start.strftime("%m/%d")} ~ {last_date.strftime("%m/%d")}</span></div>', unsafe_allow_html=True)

        this_week_style = this_week_df.groupby(['스타일코드','대표명']).agg(
            이번주=('판매금액','sum'), 수량=('수량','sum')).reset_index()
        prev_week_style = prev_week_df.groupby('스타일코드')['판매금액'].sum().reset_index()
        prev_week_style.columns = ['스타일코드','전주']

        best10 = this_week_style.sort_values('이번주', ascending=False).head(10)
        best10 = best10.merge(prev_week_style, on='스타일코드', how='left')
        best10['전주'] = best10['전주'].fillna(0)
        best10['WoW'] = np.where(best10['전주'] > 0,
            (best10['이번주'] - best10['전주']) / best10['전주'] * 100, np.nan)
        best10['이번주_만'] = (best10['이번주']/10000).round(1)

        for i, row in best10.iterrows():
            rank = best10.index.get_loc(i) + 1
            wow_html = ""
            if pd.notna(row['WoW']):
                cls = "up" if row['WoW'] > 0 else "dn"
                arrow = "▲" if row['WoW'] > 0 else "▼"
                wow_html = f'<span class="rk-delta {cls}">{arrow}{abs(row["WoW"]):.0f}%</span>'
            else:
                wow_html = '<span class="rk-delta" style="color:#ccc">신규</span>'
            medal = {1:"🥇",2:"🥈",3:"🥉"}.get(rank, f'<span style="font-size:11px;color:#999">{rank}</span>')
            name_short = row['대표명'][:24]+'…' if len(row['대표명']) > 24 else row['대표명']
            st.markdown(
                f'<div class="rk">'
                f'<span class="rk-medal">{medal}</span>'
                f'<span class="rk-name">{name_short}</span>'
                f'<span class="rk-val">{row["이번주_만"]}만</span>'
                f'{wow_html}</div>',
                unsafe_allow_html=True)

    with col_right:
        st.markdown('<div class="sec-title">📺 채널별 전주 대비</div>', unsafe_allow_html=True)

        ch_this = this_week_df.groupby('채널')['판매금액'].sum().reset_index().rename(columns={'판매금액':'이번주'})
        ch_prev = prev_week_df.groupby('채널')['판매금액'].sum().reset_index().rename(columns={'판매금액':'전주'})
        ch_comp = ch_this.merge(ch_prev, on='채널', how='outer').fillna(0)
        ch_comp['WoW'] = np.where(ch_comp['전주'] > 0,
            (ch_comp['이번주'] - ch_comp['전주']) / ch_comp['전주'] * 100, np.nan)
        ch_comp['이번주_만'] = (ch_comp['이번주']/10000).round(1)
        ch_comp = ch_comp.sort_values('이번주', ascending=False)

        for _, row in ch_comp.iterrows():
            if row['이번주'] == 0 and row['전주'] == 0: continue
            color = CH_COLORS.get(row['채널'], '#999')
            wow_html = ""
            if pd.notna(row['WoW']):
                cls = "up" if row['WoW'] > 0 else "dn"
                arrow = "▲" if row['WoW'] > 0 else "▼"
                wow_html = f'<span class="rk-delta {cls}">{arrow}{abs(row["WoW"]):.0f}%</span>'
            st.markdown(
                f'<div class="ch-row">'
                f'<span class="ch-dot" style="background:{color}"></span>'
                f'<span class="ch-name">{row["채널"]}</span>'
                f'<span class="ch-val">{row["이번주_만"]}만</span>'
                f'{wow_html}</div>',
                unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 채널별 이번주 바 차트
        if not ch_comp.empty:
            fig_ch = go.Figure()
            fig_ch.add_bar(
                x=ch_comp['채널'], y=ch_comp['이번주_만'],
                marker_color=[CH_COLORS.get(c,'#999') for c in ch_comp['채널']],
                name='이번주', text=ch_comp['이번주_만'].apply(lambda x: f"{x:.1f}만"),
                textposition='outside')
            fig_ch.update_layout(height=220, plot_bgcolor='white', showlegend=False,
                                  yaxis_title='매출(만)', margin=dict(t=10,b=10))
            st.plotly_chart(fig_ch, use_container_width=True)

    st.markdown("---")

    # ── 월별 매출 추이 ───────────────────────────────────────────────────────
    st.markdown("### 📈 월별 추이")
    col_a, col_b = st.columns([2,1])
    with col_a:
        fig = go.Figure()
        fig.add_bar(x=monthly['년월'], y=monthly['매출_만'], name='매출(만)', marker_color='#1E88E5', opacity=0.8)
        fig.add_scatter(x=monthly['년월'], y=monthly['실현율'], name='실현율(%)', yaxis='y2',
                        line=dict(color='#E53935',width=2), mode='lines+markers')
        if targets:
            tgt_line = [targets.get(ym, None) for ym in monthly['년월']]
            fig.add_scatter(x=monthly['년월'], y=tgt_line, name='목표(만)', yaxis='y',
                            line=dict(color='#FB8C00',width=2,dash='dash'), mode='lines+markers')
        fig.update_layout(title="월별 매출 & 실현율", yaxis=dict(title="매출(만)"),
                          yaxis2=dict(title="실현율(%)",overlaying='y',side='right',range=[50,100]),
                          legend=dict(orientation='h',y=1.1), height=320,
                          plot_bgcolor='white', xaxis=dict(tickangle=-45))
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        colors = ['#43A047' if v>=0 else '#E53935' for v in monthly['MoM'].fillna(0)]
        fig2 = go.Figure()
        fig2.add_bar(x=monthly['년월'], y=monthly['MoM'].round(1), marker_color=colors, name='MoM(%)')
        if monthly['YoY'].notna().any():
            fig2.add_scatter(x=monthly['년월'], y=monthly['YoY'].round(1), name='YoY(%)',
                             mode='lines+markers', line=dict(color='#FB8C00',width=2,dash='dot'))
        fig2.update_layout(title="MoM / YoY", height=320, plot_bgcolor='white',
                           legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45))
        st.plotly_chart(fig2, use_container_width=True)

    with st.expander("📋 월별 상세"):
        disp = monthly[['년월','매출_만','MoM','YoY','수량','ASP','실현율','SKU']].copy()
        disp['MoM'] = disp['MoM'].apply(lambda x: f"{x:+.1f}%" if pd.notna(x) else "-")
        disp['YoY'] = disp['YoY'].apply(lambda x: f"{x:+.1f}%" if pd.notna(x) else "-")
        disp['ASP'] = disp['ASP'].apply(lambda x: f"{x/1000:.1f}천")
        disp['실현율'] = disp['실현율'].apply(lambda x: f"{x:.1f}%")
        st.dataframe(disp, use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 2 — 카테고리
# ════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### 🏷️ 카테고리별 매출")
    cat_m = df.groupby(['년월','카테고리'])['판매금액'].sum().reset_index()
    cat_m['매출_만'] = (cat_m['판매금액']/10000).round()
    cat_t = df.groupby('카테고리')['판매금액'].sum().sort_values(ascending=False).reset_index()
    cat_t['매출_만'] = (cat_t['판매금액']/10000).round()
    cat_t['비중'] = (cat_t['판매금액']/cat_t['판매금액'].sum()*100).round(1)

    col_a,col_b = st.columns([1,2])
    with col_a:
        fig = px.pie(cat_t, values='매출_만', names='카테고리', hole=0.4,
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(height=340, showlegend=False, title='카테고리 비중')
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        fig = px.bar(cat_m, x='년월', y='매출_만', color='카테고리', title='카테고리별 월별 추이',
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_layout(height=340, plot_bgcolor='white', xaxis=dict(tickangle=-45),
                          legend=dict(orientation='h',y=1.1))
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(cat_t[['카테고리','매출_만','비중']].rename(columns={'매출_만':'매출(만)','비중':'비중(%)'}),
                 use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 3 — 채널
# ════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### 📺 채널별 분석")

    # ── 기초 집계 (가중 실현율) ────────────────────────────────────────────
    ch_m = df.groupby(['년월','채널'])['판매금액'].sum().reset_index()
    ch_m['매출_만'] = (ch_m['판매금액']/10000).round()

    # 채널별 가중 실현율 = 총판매금액 / 총택가합
    def _ch_real(g):
        tk = g['총택가'].sum()
        return g['판매금액'].sum() / tk * 100 if tk > 0 else np.nan

    ch_t = df.groupby('채널').agg(
        매출=('판매금액','sum'), 수량=('수량','sum')
    ).reset_index()
    ch_t['가중실현율'] = df.groupby('채널').apply(_ch_real).values
    ch_t['할인율'] = 100 - ch_t['가중실현율']
    ch_t['매출_만'] = (ch_t['매출']/10000).round()
    ch_t['비중'] = (ch_t['매출']/ch_t['매출'].sum()*100).round(1)
    ch_t['ASP'] = (ch_t['매출']/ch_t['수량']/1000).round(1)
    ch_t = ch_t.sort_values('매출', ascending=False).reset_index(drop=True)

    # ── 상단 KPI ───────────────────────────────────────────────────────────
    kpi_cols = st.columns(len(ch_t))
    for i, (_, row) in enumerate(ch_t.iterrows()):
        color = CH_COLORS.get(row['채널'], '#999')
        disc = row['할인율']
        disc_txt = f"할인 {disc:.1f}%" if disc > 0 else f"프리미엄 {abs(disc):.1f}%"
        disc_color = "#43A047" if disc <= 0 else ("#FB8C00" if disc <= 15 else "#E53935")
        kpi_cols[i].markdown(
            f'<div class="kpi" style="border-top-color:{color}">'
            f'<div class="kpi-label">{row["채널"]}</div>'
            f'<div class="kpi-value">{row["매출_만"]:,}만</div>'
            f'<div class="kpi-delta" style="color:{disc_color}">{disc_txt} | {row["비중"]}%</div>'
            f'</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 채널별 월매출 + 실현율 추이 ────────────────────────────────────────
    col_a, col_b = st.columns([3, 2])
    with col_a:
        fig = px.bar(ch_m, x='년월', y='매출_만', color='채널',
                     title='채널별 월별 매출 (만원)', color_discrete_map=CH_COLORS,
                     barmode='stack')
        fig.update_layout(height=340, plot_bgcolor='white', xaxis=dict(tickangle=-45),
                          legend=dict(orientation='h', y=1.12))
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        # 채널별 월별 가중 실현율
        ch_real_m = df.groupby(['년월','채널']).apply(
            lambda x: x['판매금액'].sum() / x['총택가'].sum() * 100
            if x['총택가'].sum() > 0 else np.nan
        ).reset_index(name='실현율')
        fig2 = px.line(ch_real_m, x='년월', y='실현율', color='채널',
                       title='채널별 월별 실현율(%)', color_discrete_map=CH_COLORS,
                       markers=True)
        fig2.add_hline(y=100, line_dash='dot', line_color='#888', annotation_text='택가 100%')
        fig2.update_layout(height=340, plot_bgcolor='white', xaxis=dict(tickangle=-45),
                           legend=dict(orientation='h', y=1.12),
                           yaxis=dict(range=[60, 115]))
        st.plotly_chart(fig2, use_container_width=True)

    # ── 채널별 할인율 상세 ─────────────────────────────────────────────────
    st.markdown("### 💸 채널별 할인율 분석")
    col_c, col_d = st.columns([1, 2])

    with col_c:
        # 채널별 실현율 바
        fig3 = go.Figure()
        for _, row in ch_t.sort_values('가중실현율').iterrows():
            color = CH_COLORS.get(row['채널'], '#999')
            fig3.add_bar(
                x=[row['가중실현율']], y=[row['채널']],
                orientation='h',
                marker_color=color,
                text=[f"{row['가중실현율']:.1f}%"],
                textposition='outside',
                name=row['채널']
            )
        fig3.add_vline(x=100, line_dash='dot', line_color='#888')
        fig3.update_layout(
            title='채널별 가중 실현율', height=280,
            plot_bgcolor='white', showlegend=False,
            xaxis=dict(range=[60, 115], title='실현율(%)'),
            margin=dict(l=10, r=60)
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col_d:
        # 할인율 구간 분포 (채널별)
        bins   = [-0.1, 5, 10, 15, 20, 25, 30, 40, 100]
        labels = ['0~5%','6~10%','11~15%','16~20%','21~25%','26~30%','31~40%','41%+']
        df_real = df[df['실현율'].notna()].copy()
        df_real['할인구간'] = pd.cut(df_real['할인율'], bins=bins, labels=labels)
        dist = (df_real.groupby(['채널','할인구간'])['판매금액']
                .sum().unstack(fill_value=0))
        dist_pct = dist.div(dist.sum(axis=1), axis=0) * 100
        fig4 = px.bar(
            dist_pct.reset_index().melt(id_vars='채널', var_name='할인구간', value_name='비중'),
            x='채널', y='비중', color='할인구간',
            title='채널별 할인율 구간 비중(%)',
            color_discrete_sequence=px.colors.sequential.RdYlGn_r,
            barmode='stack'
        )
        fig4.update_layout(height=280, plot_bgcolor='white',
                           legend=dict(orientation='h', y=1.12))
        st.plotly_chart(fig4, use_container_width=True)

    # ── 채널 요약 테이블 ────────────────────────────────────────────────────
    with st.expander("📋 채널 상세 요약"):
        disp_ch = ch_t[['채널','매출_만','비중','가중실현율','할인율','ASP','수량']].copy()
        disp_ch['가중실현율'] = disp_ch['가중실현율'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
        disp_ch['할인율'] = disp_ch['할인율'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
        disp_ch['ASP'] = disp_ch['ASP'].apply(lambda x: f"{x:.1f}천")
        disp_ch = disp_ch.rename(columns={'매출_만':'매출(만)','비중':'비중(%)','가중실현율':'실현율(가중)','할인율':'평균할인율','ASP':'평균단가'})
        st.dataframe(disp_ch, use_container_width=True, hide_index=True)

    # ── 이벤트 감지 (일별 실현율 급락) ────────────────────────────────────
    st.markdown("### 🔍 이벤트 / 할인 구간 감지")
    event_ch = st.selectbox("채널 선택", ['무신사','자사몰','29CM','무신사글로벌','팝업(더현대)'], key='event_ch')
    df_event = df[df['채널'] == event_ch].copy()
    if len(df_event) > 0:
        daily_real = df_event.groupby('날짜').apply(
            lambda x: x['판매금액'].sum() / x['총택가'].sum() * 100
            if x['총택가'].sum() > 0 else np.nan
        ).reset_index(name='실현율')
        daily_rev = df_event.groupby('날짜')['판매금액'].sum().reset_index(name='매출')
        daily = daily_real.merge(daily_rev, on='날짜')
        daily['매출_만'] = daily['매출'] / 10000

        fig5 = go.Figure()
        # 이벤트 기간 (실현율 < 80%) 하이라이트
        event_days = daily[daily['실현율'] < 80]['날짜']
        for ed in event_days:
            fig5.add_vrect(x0=str(ed), x1=str(ed + pd.Timedelta(days=1)),
                           fillcolor='rgba(229,57,53,0.12)', line_width=0)
        fig5.add_bar(x=daily['날짜'], y=daily['매출_만'], name='일매출(만)',
                     marker_color='#90CAF9', opacity=0.7, yaxis='y2')
        fig5.add_scatter(x=daily['날짜'], y=daily['실현율'], name='실현율(%)',
                         line=dict(color='#E53935', width=2), mode='lines+markers',
                         marker=dict(size=5))
        fig5.add_hline(y=80, line_dash='dash', line_color='#FB8C00',
                       annotation_text='80% 기준선')
        fig5.add_hline(y=100, line_dash='dot', line_color='#888',
                       annotation_text='택가 100%')
        fig5.update_layout(
            title=f'{event_ch} 일별 실현율 & 매출 (빨간 음영 = 이벤트 구간)',
            height=340, plot_bgcolor='white',
            yaxis=dict(title='실현율(%)', range=[50, 115]),
            yaxis2=dict(title='일매출(만)', overlaying='y', side='right'),
            legend=dict(orientation='h', y=1.1),
            xaxis=dict(tickangle=-45)
        )
        st.plotly_chart(fig5, use_container_width=True)

        # 이벤트 구간 요약
        if len(event_days) > 0:
            ev_df = daily[daily['실현율'] < 80]
            st.markdown(
                f'<div class="warn-card">⚠️ <b>이벤트 감지</b>: 실현율 80% 미만 구간 '
                f'<b>{len(ev_df)}일</b> — '
                f'{ev_df["날짜"].min().strftime("%m/%d")}~{ev_df["날짜"].max().strftime("%m/%d")} | '
                f'평균 실현율 {ev_df["실현율"].mean():.1f}% | '
                f'이벤트 매출 {ev_df["매출_만"].sum():,.0f}만</div>',
                unsafe_allow_html=True)
        else:
            st.markdown('<div class="good-card">✅ 이벤트 구간(실현율 80% 미만) 없음</div>',
                        unsafe_allow_html=True)
    else:
        st.info("해당 채널 데이터 없음")

    # ── 월별 채널 할인율 히트맵 ────────────────────────────────────────────
    with st.expander("🗓️ 월별 × 채널별 실현율 히트맵"):
        heatmap_data = df.groupby(['년월','채널']).apply(
            lambda x: round(x['판매금액'].sum() / x['총택가'].sum() * 100, 1)
            if x['총택가'].sum() > 0 else np.nan
        ).unstack()
        fig6 = px.imshow(
            heatmap_data,
            color_continuous_scale='RdYlGn',
            zmin=70, zmax=110,
            text_auto='.1f',
            title='월별 × 채널별 실현율(%) 히트맵',
            aspect='auto'
        )
        fig6.update_layout(height=350)
        st.plotly_chart(fig6, use_container_width=True)
        st.caption("※ 100% 초과 = 환율 프리미엄(무신사글로벌). 80% 미만 = 이벤트/기획전 구간.")

    # ── 채널 간 상관 ───────────────────────────────────────────────────────
    with st.expander("🔀 채널 매출 상관 분석"):
        ch_monthly_wide = df_all.groupby(['년월','채널'])['판매금액'].sum().unstack(fill_value=0).reset_index()
        ch_for_corr = [c for c in CHANNELS if c in ch_monthly_wide.columns and ch_monthly_wide[c].sum() > 0]
        if len(ch_for_corr) >= 2:
            corr_matrix = ch_monthly_wide[ch_for_corr].corr().round(2)
            fig7 = px.imshow(corr_matrix, color_continuous_scale='RdBu_r', zmin=-1, zmax=1,
                             text_auto=True, title='채널 간 매출 상관계수')
            fig7.update_layout(height=300)
            st.plotly_chart(fig7, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 4 — 상품랭킹
# ════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown("### 🏆 상품별 순위")
    style_agg = df.groupby(['스타일코드','대표명','카테고리','상품구분']).agg(
        매출=('판매금액','sum'), 수량=('수량','sum')).reset_index()
    style_agg = style_agg.sort_values('매출', ascending=False).reset_index(drop=True)
    style_agg['순위'] = style_agg.index+1
    style_agg['매출_만'] = (style_agg['매출']/10000).round()
    style_agg['ASP'] = (style_agg['매출']/style_agg['수량']/1000).round(1)
    style_agg['누적비중'] = (style_agg['매출'].cumsum()/style_agg['매출'].sum()*100).round(1)

    col_a,col_b = st.columns([2,1])
    with col_a:
        top20 = style_agg.head(20)
        fig = go.Figure(go.Bar(
            x=top20['매출_만'], orientation='h',
            y=top20['대표명'].apply(lambda x: x[:20]+'...' if len(x)>20 else x),
            marker_color=['#1E88E5' if g=='주력' else '#90CAF9' for g in top20['상품구분']],
            text=top20['매출_만'].apply(lambda x: f"{x:,.0f}만"), textposition='outside'))
        fig.update_layout(title='TOP20 상품', height=580, plot_bgcolor='white',
                          yaxis=dict(autorange='reversed'), xaxis_title='매출(만)')
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        grp = style_agg.groupby('상품구분')['매출'].sum().reset_index()
        fig2 = px.pie(grp, values='매출', names='상품구분', hole=0.4,
                      color_discrete_map={'주력':'#1E88E5','기타':'#90CAF9'}, title='주력/기타')
        fig2.update_layout(height=260)
        st.plotly_chart(fig2, use_container_width=True)
        cat_rank = style_agg.groupby('카테고리')['매출_만'].sum().sort_values(ascending=False).head(8)
        st.markdown("**카테고리 순위**")
        for i,(cat,rev) in enumerate(cat_rank.items(),1):
            st.markdown(f"`{i}위` **{cat}** — {rev:,.0f}만")

    with st.expander("📋 전체 TOP50"):
        disp = style_agg.head(50)[['순위','대표명','카테고리','상품구분','매출_만','수량','ASP','누적비중']].copy()
        disp['누적비중'] = disp['누적비중'].apply(lambda x: f"{x}%")
        st.dataframe(disp, use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 5 — 상품 분석 (수명주기 + 재고소진 + 이상치)
# ════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown("### 🔬 상품 분석")
    sub_a, sub_b, sub_c = st.tabs(["🔬 수명주기 (PLC)", "⏱️ 재고소진 예측", "🚨 이상치 감지"])

    # ── PLC ──────────────────────────────────────────────────────────────────
    with sub_a:
        st.caption("최근 판매 추세 기울기와 누적 판매기간으로 도입→성장→성숙→쇠퇴 자동 태깅")
        style_monthly = df_all.groupby(['스타일코드','대표명','년월'])['판매금액'].sum().reset_index()
        months_sorted = sorted(df_all['년월'].unique())
        plc_rows = []
        for (sc, nm), grp in style_monthly.groupby(['스타일코드','대표명']):
            grp = grp.sort_values('년월')
            first_sale = grp[grp['판매금액']>0]['년월'].min()
            if pd.isna(first_sale): continue
            months_active = len(months_sorted) - months_sorted.index(first_sale)
            recent_3 = grp[grp['년월'] >= months_sorted[-3]]['판매금액'].sum() if len(months_sorted)>=3 else 0
            prev_3_start = max(0, len(months_sorted)-6)
            prev_3 = grp[(grp['년월'] >= months_sorted[prev_3_start]) & (grp['년월'] < months_sorted[-3])]['판매금액'].sum()
            all_m = months_sorted
            y = [grp[grp['년월']==m]['판매금액'].sum() if m in grp['년월'].values else 0 for m in all_m]
            slope, _, _, _, _ = stats.linregress(range(len(y)), y) if len(y) >= 3 and sum(y) > 0 else (0,0,0,0,0)
            trend_pct = (recent_3-prev_3)/prev_3*100 if prev_3>0 else 0
            total_rev = grp['판매금액'].sum()
            if months_active <= 3: plc = '🌱 도입'
            elif slope > 50000 and trend_pct > 10: plc = '🚀 성장'
            elif abs(slope) <= 50000 and total_rev > 5000000: plc = '💪 성숙'
            elif slope < -50000 or trend_pct < -30: plc = '📉 쇠퇴'
            else: plc = '➡️ 유지'
            plc_rows.append({'스타일코드':sc,'상품명':nm,'PLC':plc,'총매출_만':round(total_rev/10000),
                             '판매기간(월)':months_active,'최근3개월_만':round(recent_3/10000),
                             '추세변화(%)':round(trend_pct,1),'기울기':round(slope)})
        plc_df = pd.DataFrame(plc_rows).sort_values('총매출_만', ascending=False)
        plc_counts = plc_df['PLC'].value_counts()
        cols = st.columns(5)
        for i, (label, color) in enumerate([('🌱 도입','#43A047'),('🚀 성장','#1E88E5'),
                                             ('💪 성숙','#FB8C00'),('📉 쇠퇴','#E53935'),('➡️ 유지','#757575')]):
            kpi_card(cols[i], label, f"{plc_counts.get(label,0)}개", color=color)
        fig = px.scatter(plc_df, x='판매기간(월)', y='최근3개월_만', size='총매출_만',
                         color='PLC', hover_data=['상품명','추세변화(%)'],
                         title='PLC 포지션 맵 (크기=누적 매출)',
                         color_discrete_map={'🌱 도입':'#43A047','🚀 성장':'#1E88E5',
                                             '💪 성숙':'#FB8C00','📉 쇠퇴':'#E53935','➡️ 유지':'#9E9E9E'})
        fig.update_layout(height=380, plot_bgcolor='white')
        st.plotly_chart(fig, use_container_width=True)
        plc_filter = st.multiselect("PLC 단계 필터", ['🌱 도입','🚀 성장','💪 성숙','📉 쇠퇴','➡️ 유지'],
                                     default=['🚀 성장','📉 쇠퇴'])
        filtered_plc = plc_df[plc_df['PLC'].isin(plc_filter)] if plc_filter else plc_df
        st.dataframe(filtered_plc[['상품명','PLC','총매출_만','판매기간(월)','최근3개월_만','추세변화(%)']],
                     use_container_width=True, hide_index=True)

    # ── 재고소진 예측 ─────────────────────────────────────────────────────────
    with sub_b:
        st.caption("최근 4주 판매 속도 vs 이전 4주 비교 — 급감 = 재고 소진 신호")
        last_date2 = df_all['날짜'].max()
        w4_start = last_date2 - timedelta(weeks=4)
        w8_start = last_date2 - timedelta(weeks=8)
        recent4 = df_all[df_all['날짜'] > w4_start].groupby(['스타일코드','대표명']).agg(
            최근4주_수량=('수량','sum'), 최근4주_매출=('판매금액','sum'),
            마지막판매=('날짜','max')).reset_index()
        prev4 = df_all[(df_all['날짜'] > w8_start) & (df_all['날짜'] <= w4_start)].groupby('스타일코드').agg(
            이전4주_수량=('수량','sum')).reset_index()
        inv_df = recent4.merge(prev4, on='스타일코드', how='left')
        inv_df['이전4주_수량'] = inv_df['이전4주_수량'].fillna(0)
        inv_df['수량변화(%)'] = np.where(inv_df['이전4주_수량']>0,
            (inv_df['최근4주_수량']-inv_df['이전4주_수량'])/inv_df['이전4주_수량']*100, 0)
        inv_df['경과일'] = (last_date2 - inv_df['마지막판매']).dt.days
        def risk_level(row):
            if row['경과일'] > 14 and row['최근4주_수량'] == 0: return '🔴 소진 의심'
            if row['수량변화(%)'] < -60 or (row['경과일'] > 7 and row['최근4주_수량'] < 3): return '🟡 소진 임박'
            if row['수량변화(%)'] < -30: return '🟠 감소 주의'
            return '🟢 정상'
        inv_df['리스크'] = inv_df.apply(risk_level, axis=1)
        inv_df['최근4주_매출_만'] = (inv_df['최근4주_매출']/10000).round(1)
        inv_df['마지막판매'] = inv_df['마지막판매'].dt.strftime('%Y-%m-%d')
        inv_df = inv_df.sort_values('수량변화(%)')
        rc = inv_df['리스크'].value_counts()
        c1,c2,c3,c4 = st.columns(4)
        kpi_card(c1, "🔴 소진 의심", f"{rc.get('🔴 소진 의심',0)}개", color="#E53935")
        kpi_card(c2, "🟡 소진 임박", f"{rc.get('🟡 소진 임박',0)}개", color="#FB8C00")
        kpi_card(c3, "🟠 감소 주의", f"{rc.get('🟠 감소 주의',0)}개", color="#FF6F00")
        kpi_card(c4, "🟢 정상", f"{rc.get('🟢 정상',0)}개", color="#43A047")
        danger = inv_df[inv_df['리스크'] != '🟢 정상']
        for _, row in danger.iterrows():
            cls = 'danger-card' if '🔴' in row['리스크'] else 'warn-card'
            st.markdown(f'<div class="{cls}"><b>{row["리스크"]} {row["대표명"]}</b> — 최근4주 {int(row["최근4주_수량"])}개 (전주 대비 {row["수량변화(%)"]:.0f}%) | 마지막판매: {row["마지막판매"]} ({int(row["경과일"])}일 전)</div>', unsafe_allow_html=True)
        top_risk = inv_df[inv_df['이전4주_수량']>0].head(20)
        fig = go.Figure(go.Bar(
            x=top_risk['수량변화(%)'].round(0),
            y=top_risk['대표명'].apply(lambda x: x[:18]+'...' if len(x)>18 else x),
            orientation='h',
            marker_color=['#E53935' if v<-60 else '#FB8C00' if v<-30 else '#43A047' for v in top_risk['수량변화(%)']],
            text=top_risk['수량변화(%)'].apply(lambda x: f"{x:.0f}%"), textposition='outside'))
        fig.update_layout(title='최근 4주 판매량 변화 TOP20', height=500, plot_bgcolor='white',
                          yaxis=dict(autorange='reversed'), xaxis_title='변화(%)')
        st.plotly_chart(fig, use_container_width=True)

    # ── 이상치 감지 ───────────────────────────────────────────────────────────
    with sub_c:
        st.caption("전주 vs 4주 평균 대비 ±30% 이상 급변 상품 자동 감지")
        weekly = df_all.groupby(['주차','스타일코드','대표명'])['판매금액'].sum().reset_index()
        weeks_sorted = sorted(df_all['주차'].unique())
        if len(weeks_sorted) < 2:
            st.warning("주차 데이터가 부족합니다.")
        else:
            last_w = weeks_sorted[-1]
            prev_4w = weeks_sorted[max(0,len(weeks_sorted)-5):-1]
            last_week_s = weekly[weekly['주차']==last_w].set_index('스타일코드')['판매금액']
            avg_4w = weekly[weekly['주차'].isin(prev_4w)].groupby('스타일코드')['판매금액'].mean()
            anomaly_rows = []
            for sc in set(last_week_s.index) | set(avg_4w.index):
                lw = last_week_s.get(sc, 0)
                a4 = avg_4w.get(sc, 0)
                if a4 == 0 and lw == 0: continue
                chg = (lw-a4)/a4*100 if a4>0 else (100 if lw>0 else 0)
                nm = df_all[df_all['스타일코드']==sc]['대표명'].mode()
                nm = nm.iloc[0] if len(nm)>0 else sc
                if abs(chg) >= 30 or (a4>0 and lw==0):
                    anomaly_rows.append({'스타일코드':sc,'상품명':nm,'이번주_만':round(lw/10000,1),
                                         '4주평균_만':round(a4/10000,1),'변화(%)':round(chg,1)})
            anom_df = pd.DataFrame(anomaly_rows).sort_values('변화(%)')
            if anom_df.empty:
                st.success("이번 주 이상치 없음")
            else:
                surge = anom_df[anom_df['변화(%)'] >= 30].sort_values('변화(%)', ascending=False)
                drop = anom_df[anom_df['변화(%)'] < -30].sort_values('변화(%)')
                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown(f"#### 🔺 급등 ({len(surge)}개)")
                    for _, r in surge.iterrows():
                        st.markdown(f'<div class="good-card"><b>{r["상품명"]}</b><br>이번주 {r["이번주_만"]}만 | 4주평균 {r["4주평균_만"]}만 | <b>+{r["변화(%)"]:.0f}%</b></div>', unsafe_allow_html=True)
                with col_b:
                    st.markdown(f"#### 🔻 급락 ({len(drop)}개)")
                    for _, r in drop.iterrows():
                        st.markdown(f'<div class="danger-card"><b>{r["상품명"]}</b><br>이번주 {r["이번주_만"]}만 | 4주평균 {r["4주평균_만"]}만 | <b>{r["변화(%)"]:.0f}%</b></div>', unsafe_allow_html=True)
            weekly_total = df_all.groupby('주차')['판매금액'].sum().reset_index()
            weekly_total['매출_만'] = (weekly_total['판매금액']/10000).round()
            weekly_total['4주이동평균'] = weekly_total['매출_만'].rolling(4, min_periods=1).mean().round()
            fig = go.Figure()
            fig.add_bar(x=weekly_total['주차'], y=weekly_total['매출_만'], name='주간매출', marker_color='#90CAF9')
            fig.add_scatter(x=weekly_total['주차'], y=weekly_total['4주이동평균'], name='4주 이동평균',
                            line=dict(color='#E53935',width=2))
            fig.update_layout(title='주간 매출 추이', height=300, plot_bgcolor='white',
                              legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45, nticks=10))
            st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 6 — 프로모션
# ════════════════════════════════════════════════════════════════════════════
PROMO_TYPES = ["팝업", "기획전", "할인행사", "콜라보", "시즌오프", "기타"]
PROMO_CHANNELS = ["전채널", "무신사", "자사몰", "29CM", "팝업(더현대)", "무신사글로벌"]
TYPE_COLORS = {"팝업":"#8E24AA","기획전":"#1E88E5","할인행사":"#E53935","콜라보":"#FB8C00","시즌오프":"#43A047","기타":"#607D8B"}

with tab6:
    st.markdown("### 🎪 프로모션")
    mgmt_p = load_mgmt()
    promotions = mgmt_p.get("promotions", [])
    pr_sub1, pr_sub2, pr_sub3 = st.tabs(["📋 전체 목록", "🔍 프로모션 분석", "➕ 신규 등록"])

    # ── 전체 목록 ─────────────────────────────────────────────────────────────
    with pr_sub1:
        if not promotions:
            st.info("등록된 프로모션이 없습니다. '신규 등록' 탭에서 추가하세요.")
        else:
            st.markdown(f"**총 {len(promotions)}개 프로모션**")
            for i, p in enumerate(sorted(promotions, key=lambda x: x.get('start_date',''), reverse=True)):
                tc = TYPE_COLORS.get(p.get('type','기타'), '#607D8B')
                days = 0
                try:
                    s = datetime.strptime(p['start_date'], '%Y-%m-%d')
                    e = datetime.strptime(p['end_date'], '%Y-%m-%d')
                    days = (e - s).days + 1
                except: pass
                col_a, col_b = st.columns([8, 1])
                with col_a:
                    st.markdown(
                        f'<div class="metric-card" style="border-left-color:{tc};padding:10px 14px">'
                        f'<span style="background:{tc};color:white;border-radius:4px;padding:2px 8px;font-size:11px">{p.get("type","기타")}</span>'
                        f'<span style="margin-left:8px;font-weight:600;font-size:15px">{p.get("name","")}</span>'
                        f'<span style="margin-left:10px;color:#888;font-size:12px">{p.get("channel","")}</span><br>'
                        f'<span style="color:#555;font-size:12px">📅 {p.get("start_date","")} ~ {p.get("end_date","")} ({days}일)'
                        f'{"  |  🎯 목표 "+str(p.get("target_rev",""))+"만" if p.get("target_rev") else ""}'
                        f'{"  |  "+p.get("memo","") if p.get("memo") else ""}</span>'
                        f'</div>', unsafe_allow_html=True)
                with col_b:
                    if st.button("🗑️", key=f"del_promo_{i}"):
                        promotions_del = [x for j,x in enumerate(promotions) if x.get('name') != p.get('name') or x.get('start_date') != p.get('start_date')]
                        mgmt_p['promotions'] = promotions_del
                        save_mgmt(mgmt_p); st.rerun()

    # ── 프로모션 분석 ─────────────────────────────────────────────────────────
    with pr_sub2:
        if not promotions:
            st.info("등록된 프로모션이 없습니다.")
        else:
            promo_labels = [f"{p.get('name','')}  ({p.get('start_date','')} ~ {p.get('end_date','')})" for p in promotions]
            sel_idx = st.selectbox("분석할 프로모션 선택", range(len(promo_labels)), format_func=lambda i: promo_labels[i])
            sel_p = promotions[sel_idx]

            try:
                p_start = pd.to_datetime(sel_p['start_date'])
                p_end   = pd.to_datetime(sel_p['end_date'])
                p_days  = (p_end - p_start).days + 1
                prev_end   = p_start - pd.Timedelta(days=1)
                prev_start = prev_end - pd.Timedelta(days=p_days - 1)
            except:
                st.error("날짜 형식 오류"); st.stop()

            ch = sel_p.get('channel', '전채널')
            if ch == '전채널':
                p_df   = df_all[(df_all['날짜'] >= p_start) & (df_all['날짜'] <= p_end)]
                prev_df = df_all[(df_all['날짜'] >= prev_start) & (df_all['날짜'] <= prev_end)]
            else:
                p_df   = df_all[(df_all['날짜'] >= p_start) & (df_all['날짜'] <= p_end) & (df_all['채널'] == ch)]
                prev_df = df_all[(df_all['날짜'] >= prev_start) & (df_all['날짜'] <= prev_end) & (df_all['채널'] == ch)]

            tc = TYPE_COLORS.get(sel_p.get('type','기타'), '#607D8B')
            st.markdown(f'<div class="metric-card" style="border-left-color:{tc}">'
                        f'<b>{sel_p.get("name","")}</b> | {sel_p.get("type","")} | {ch} | {sel_p.get("start_date","")} ~ {sel_p.get("end_date","")}'
                        f'</div>', unsafe_allow_html=True)

            # KPI
            rev   = p_df['판매금액'].sum()
            qty   = p_df['수량'].sum()
            prev_rev = prev_df['판매금액'].sum()
            target_val = (sel_p.get('target_rev') or 0) * 10000
            wow = (rev - prev_rev) / prev_rev * 100 if prev_rev > 0 else None
            ach = rev / target_val * 100 if target_val > 0 else None

            c1,c2,c3,c4 = st.columns(4)
            kpi_card(c1, "프로모션 매출", f"{rev/10000:,.1f}만", color=tc)
            kpi_card(c2, "판매 수량", f"{int(qty):,}개", color=tc)
            kpi_card(c3, "전기간 대비", f"{wow:+.1f}%" if wow is not None else "-", color="#43A047" if (wow or 0) >= 0 else "#E53935")
            kpi_card(c4, "목표 달성율", f"{ach:.1f}%" if ach is not None else "목표 미설정", color="#43A047" if (ach or 0) >= 100 else "#FB8C00")

            if p_df.empty:
                st.warning("해당 기간/채널 데이터가 없습니다.")
            else:
                col1, col2 = st.columns([3,2])
                with col1:
                    # 일별 매출 트렌드
                    daily = p_df.groupby('날짜')['판매금액'].sum().reset_index()
                    daily_prev = prev_df.groupby('날짜')['판매금액'].sum().reset_index() if not prev_df.empty else pd.DataFrame()
                    trend_fig = go.Figure()
                    trend_fig.add_scatter(x=daily['날짜'], y=(daily['판매금액']/10000).round(1),
                                         name='프로모션 기간', mode='lines+markers',
                                         line=dict(color=tc, width=2), fill='tozeroy', fillcolor=tc+'22')
                    if not daily_prev.empty:
                        trend_fig.add_scatter(x=daily_prev['날짜'], y=(daily_prev['판매금액']/10000).round(1),
                                             name='전기간', mode='lines', line=dict(color='#aaa', width=1, dash='dot'))
                    trend_fig.update_layout(title='일별 매출 (만원)', height=280, plot_bgcolor='white',
                                           legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-30))
                    st.plotly_chart(trend_fig, use_container_width=True)
                with col2:
                    # 상품별 TOP 10
                    top_items = p_df.groupby(['스타일코드','대표명']).agg(
                        매출=('판매금액','sum'), 수량=('수량','sum')).reset_index()
                    top_items = top_items.sort_values('매출', ascending=False).head(10)
                    top_items['매출_만'] = (top_items['매출']/10000).round(1)
                    item_fig = go.Figure(go.Bar(
                        x=top_items['매출_만'],
                        y=top_items['대표명'].apply(lambda x: x[:18]+'...' if len(x)>18 else x),
                        orientation='h', marker_color=tc))
                    item_fig.update_layout(title='상품별 매출 TOP10', height=280,
                                          plot_bgcolor='white', xaxis_title='만원',
                                          yaxis=dict(autorange='reversed'))
                    st.plotly_chart(item_fig, use_container_width=True)

                # 채널별 분포 (전채널인 경우)
                if ch == '전채널' and p_df['채널'].nunique() > 1:
                    ch_breakdown = p_df.groupby('채널')['판매금액'].sum().reset_index().sort_values('판매금액', ascending=False)
                    ch_breakdown['매출_만'] = (ch_breakdown['판매금액']/10000).round(1)
                    ch_fig = go.Figure(go.Bar(x=ch_breakdown['채널'], y=ch_breakdown['매출_만'],
                                             marker_color=[CH_COLORS.get(c,'#888') for c in ch_breakdown['채널']]))
                    ch_fig.update_layout(title='채널별 매출 분포', height=220, plot_bgcolor='white')
                    st.plotly_chart(ch_fig, use_container_width=True)

                # 인사이트 자동 생성
                st.markdown("---")
                st.markdown("### 💡 인사이트")
                insights = []
                if wow is not None:
                    if wow >= 20:
                        insights.append(f"✅ **전기간 대비 +{wow:.1f}%** — 프로모션 효과 뚜렷. 동일 유형 반복 운영 권장.")
                    elif wow >= 0:
                        insights.append(f"🟡 **전기간 대비 +{wow:.1f}%** — 소폭 증가. 프로모션 조건 강화 시 효과 개선 가능.")
                    else:
                        insights.append(f"🔴 **전기간 대비 {wow:.1f}%** — 매출 하락. 프로모션 유형/채널/타이밍 재검토 필요.")
                if ach is not None:
                    if ach >= 100:
                        insights.append(f"🎯 **목표 달성율 {ach:.1f}%** — 목표 초과 달성.")
                    else:
                        insights.append(f"⚠️ **목표 달성율 {ach:.1f}%** — 목표 미달 {target_val/10000 - rev/10000:,.1f}만원 부족.")
                if not top_items.empty:
                    top1 = top_items.iloc[0]
                    insights.append(f"🏆 **최고 매출 상품:** {top1['대표명']} ({top1['매출_만']}만 / {int(top1['수량'])}개)")
                    top3_share = top_items.head(3)['매출'].sum() / rev * 100
                    insights.append(f"📊 **TOP3 집중도:** 상위 3개 상품이 프로모션 매출의 {top3_share:.1f}% 차지.")
                avg_daily = rev / p_days / 10000
                insights.append(f"📅 **일평균 매출:** {avg_daily:.1f}만원 ({p_days}일 기준)")
                for txt in insights:
                    st.markdown(f'<div class="insight-card">{txt}</div>', unsafe_allow_html=True)

                with st.expander("📋 전체 상품 목록"):
                    full_items = p_df.groupby(['스타일코드','대표명','카테고리']).agg(
                        매출=('판매금액','sum'), 수량=('수량','sum')).reset_index()
                    full_items = full_items.sort_values('매출', ascending=False)
                    full_items['매출(만)'] = (full_items['매출']/10000).round(1)
                    st.dataframe(full_items[['대표명','카테고리','매출(만)','수량']], use_container_width=True, hide_index=True)

    # ── 신규 등록 ─────────────────────────────────────────────────────────────
    with pr_sub3:
        st.markdown("**새 프로모션 등록**")
        r1c1, r1c2, r1c3 = st.columns(3)
        new_name    = r1c1.text_input("프로모션명", placeholder="예: 무신사 여름 기획전")
        new_type    = r1c2.selectbox("유형", PROMO_TYPES)
        new_channel = r1c3.selectbox("채널", PROMO_CHANNELS)

        r2c1, r2c2, r2c3 = st.columns(3)
        new_start  = r2c1.date_input("시작일")
        new_end    = r2c2.date_input("종료일")
        new_target = r2c3.number_input("목표 매출 (만원, 0=미설정)", min_value=0, step=100)

        new_memo = st.text_area("메모 (선택)", placeholder="예: 여름 시즌오프, 티셔츠 20% 할인", height=80)

        if st.button("💾 등록", type="primary"):
            if not new_name:
                st.error("프로모션명을 입력하세요.")
            elif new_end < new_start:
                st.error("종료일이 시작일보다 빠릅니다.")
            else:
                new_promo = {
                    "name": new_name,
                    "type": new_type,
                    "channel": new_channel,
                    "start_date": str(new_start),
                    "end_date": str(new_end),
                    "target_rev": int(new_target) if new_target > 0 else None,
                    "memo": new_memo,
                    "registered": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }
                promotions.append(new_promo)
                mgmt_p['promotions'] = promotions
                save_mgmt(mgmt_p)
                st.success(f"✅ '{new_name}' 등록 완료!")
                st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# TAB 7 — 마케팅
# ════════════════════════════════════════════════════════════════════════════
with tab7:
    st.markdown("### 📣 마케팅")
    mkt_sub = st.tabs(["📊 마케팅 효율", "💡 이번주 인사이트"])

    # ── 마케팅 효율 ───────────────────────────────────────────────────────────
    with mkt_sub[0]:
        mktg = load_marketing(str(MKTG_FILE))
        if mktg is None:
            st.warning(f"마케팅 파일을 찾을 수 없습니다: {MKTG_FILE}")
        else:
            meta_ms = mktg['meta_ms']; meta_js = mktg['meta_js']
            prod_ms = mktg['prod_ms']; prod_29 = mktg['prod_29']
            total_adspend = (meta_ms['지출 금액 (KRW)'].sum() + meta_js['지출 금액 (KRW)'].sum() +
                             prod_ms['집행 광고비'].sum() + prod_29['집행 광고비'].sum())
            total_conv = (meta_ms['공유 항목의 구매 전환값'].sum() + meta_js['구매 전환값'].sum() +
                          prod_ms['매출'].sum() + prod_29['매출'].sum())
            total_roas = total_conv / total_adspend if total_adspend > 0 else 0
            total_purchase = meta_ms['공유 항목이 포함된 구매'].sum() + meta_js['구매'].sum()
            c1,c2,c3,c4 = st.columns(4)
            kpi_card(c1, "총 광고비", f"{total_adspend/10000:,.0f}만원", color="#E53935")
            kpi_card(c2, "총 구매전환값", f"{total_conv/10000:,.0f}만원", color="#1E88E5")
            kpi_card(c3, "전체 ROAS", f"{total_roas:.1f}x", color="#43A047")
            kpi_card(c4, "총 구매 건수", f"{int(total_purchase):,}건", color="#FB8C00")
            st.markdown("---")

            # 채널별 월별 집계
            MKTG_COLORS = {'무신사 메타':'#1E88E5','자사몰 메타':'#43A047',
                           '무신사 상품광고':'#FB8C00','29CM 상품광고':'#E53935'}
            def make_ch_monthly(d, adcol, convcol, clickcol, impcol, chname):
                m = d.groupby('년월').agg(광고비=(adcol,'sum'), 구매전환값=(convcol,'sum'),
                                          클릭=(clickcol,'sum'), 노출=(impcol,'sum')).reset_index()
                m['ROAS'] = (m['구매전환값']/m['광고비'].replace(0,np.nan)).fillna(0).round(2)
                m['CTR'] = (m['클릭']/m['노출'].replace(0,np.nan)*100).fillna(0).round(3)
                m['CPC'] = (m['광고비']/m['클릭'].replace(0,np.nan)).fillna(0).round(0)
                m['채널'] = chname
                return m
            ms_m = make_ch_monthly(meta_ms,'지출 금액 (KRW)','공유 항목의 구매 전환값','링크 클릭','노출','무신사 메타')
            js_m = make_ch_monthly(meta_js,'지출 금액 (KRW)','구매 전환값','링크 클릭','노출','자사몰 메타')
            pm_m = make_ch_monthly(prod_ms,'집행 광고비','매출','클릭 수','노출 수','무신사 상품광고')
            p29_m = make_ch_monthly(prod_29,'집행 광고비','매출','클릭 수','노출 수','29CM 상품광고')
            all_ch = pd.concat([ms_m, js_m, pm_m, p29_m], ignore_index=True)

            col_a, col_b = st.columns(2)
            with col_a:
                fig = px.line(all_ch, x='년월', y='ROAS', color='채널', markers=True,
                              title='채널별 월별 ROAS', color_discrete_map=MKTG_COLORS)
                fig.add_hline(y=3, line_dash='dash', line_color='gray', annotation_text='ROAS 3x')
                fig.update_layout(height=300, plot_bgcolor='white', legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-30))
                st.plotly_chart(fig, use_container_width=True)
            with col_b:
                fig2 = px.bar(all_ch, x='년월', y='광고비', color='채널', barmode='stack',
                              title='채널별 월별 광고비', color_discrete_map=MKTG_COLORS)
                fig2.update_layout(height=300, plot_bgcolor='white', legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-30))
                st.plotly_chart(fig2, use_container_width=True)

            summary_rows = []
            for d, ch in [(ms_m,'무신사 메타'),(js_m,'자사몰 메타'),(pm_m,'무신사 상품광고'),(p29_m,'29CM 상품광고')]:
                summary_rows.append({'채널':ch,
                    '총 광고비':f"{d['광고비'].sum()/10000:,.0f}만",
                    '총 전환값':f"{d['구매전환값'].sum()/10000:,.0f}만",
                    '평균 ROAS':f"{(d['구매전환값'].sum()/d['광고비'].sum()):.1f}x" if d['광고비'].sum()>0 else "-",
                    '평균 CTR':f"{d['CTR'].mean():.3f}%",
                    '평균 CPC':f"{d['CPC'].mean():,.0f}원"})
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    # ── 이번주 인사이트 ───────────────────────────────────────────────────────
    with mkt_sub[1]:
        st.markdown("### 💡 이번주 채널별 마케팅 인사이트")
        mktg2 = load_marketing(str(MKTG_FILE))
        if mktg2 is None:
            st.warning("마케팅 파일이 없습니다.")
        else:
            meta_ms2 = mktg2['meta_ms']; meta_js2 = mktg2['meta_js']
            prod_ms2 = mktg2['prod_ms']; prod_29_2 = mktg2['prod_29']

            # 이번주/전주 광고 데이터
            last_w2 = sorted(meta_ms2['주차'].unique())[-1] if '주차' in meta_ms2.columns and len(meta_ms2) > 0 else None

            def get_week_summary(d, adcol, convcol, clickcol):
                if '주차' not in d.columns or last_w2 is None: return None, None
                weeks = sorted(d['주차'].unique())
                if len(weeks) < 1: return None, None
                last_w_data = d[d['주차'] == weeks[-1]]
                prev_w_data = d[d['주차'] == weeks[-2]] if len(weeks) >= 2 else pd.DataFrame()
                this_spend = last_w_data[adcol].sum()
                this_conv = last_w_data[convcol].sum()
                this_roas = this_conv / this_spend if this_spend > 0 else 0
                prev_spend = prev_w_data[adcol].sum() if not prev_w_data.empty else 0
                prev_conv = prev_w_data[convcol].sum() if not prev_w_data.empty else 0
                prev_roas = prev_conv / prev_spend if prev_spend > 0 else 0
                return {'광고비':this_spend, '전환값':this_conv, 'ROAS':this_roas,
                        '전주광고비':prev_spend, '전주전환값':prev_conv, '전주ROAS':prev_roas}, None

            ch_insights = {
                '무신사 메타': get_week_summary(meta_ms2,'지출 금액 (KRW)','공유 항목의 구매 전환값','링크 클릭')[0],
                '자사몰 메타': get_week_summary(meta_js2,'지출 금액 (KRW)','구매 전환값','링크 클릭')[0],
                '무신사 상품광고': get_week_summary(prod_ms2,'집행 광고비','매출','클릭 수')[0],
                '29CM 상품광고': get_week_summary(prod_29_2,'집행 광고비','매출','클릭 수')[0],
            }

            # 채널별 카드
            col1, col2 = st.columns(2)
            insight_texts = []
            for idx, (ch_name, data) in enumerate(ch_insights.items()):
                col = col1 if idx % 2 == 0 else col2
                with col:
                    if data:
                        roas_chg = data['ROAS'] - data['전주ROAS']
                        spend_chg = (data['광고비'] - data['전주광고비']) / data['전주광고비'] * 100 if data['전주광고비'] > 0 else 0
                        roas_color = "#43A047" if data['ROAS'] >= 3 else ("#FB8C00" if data['ROAS'] >= 1.5 else "#E53935")
                        roas_arrow = "▲" if roas_chg > 0 else "▼"
                        roas_cls = "up" if roas_chg > 0 else "dn"
                        st.markdown(
                            f'<div class="metric-card" style="border-left-color:{roas_color}">'
                            f'<div class="metric-label">{ch_name}</div>'
                            f'<div class="metric-value">ROAS {data["ROAS"]:.1f}x '
                            f'<span class="{roas_cls}" style="font-size:14px">{roas_arrow}{abs(roas_chg):.1f}</span></div>'
                            f'<div style="font-size:12px;color:#666;margin-top:4px">'
                            f'광고비 {data["광고비"]/10000:.1f}만 | 전환값 {data["전환값"]/10000:.1f}만 | '
                            f'전주대비 광고비 {spend_chg:+.0f}%</div></div>',
                            unsafe_allow_html=True)
                        # 인사이트 자동 생성
                        if data['ROAS'] >= 4:
                            insight_texts.append(f"✅ **{ch_name}**: ROAS {data['ROAS']:.1f}x — 고효율 채널. 예산 증액 검토 필요.")
                        elif data['ROAS'] >= 2:
                            insight_texts.append(f"🟡 **{ch_name}**: ROAS {data['ROAS']:.1f}x — 준수한 성과. 크리에이티브 A/B테스트로 개선 여지 있음.")
                        else:
                            insight_texts.append(f"🔴 **{ch_name}**: ROAS {data['ROAS']:.1f}x — 효율 낮음. 예산 축소 또는 소재 교체 필요.")
                        if roas_chg > 0.5:
                            insight_texts.append(f"   → 전주 대비 ROAS +{roas_chg:.1f} 상승 — 현재 캠페인 유지 권장.")
                        elif roas_chg < -0.5:
                            insight_texts.append(f"   → 전주 대비 ROAS {roas_chg:.1f} 하락 — 소재 피로도 확인 필요.")

            st.markdown("---")
            st.markdown("### 📋 인사이트 요약")
            if insight_texts:
                for txt in insight_texts:
                    st.markdown(f'<div class="insight-card">{txt}</div>', unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 🧭 추후 방향성 제안")

            # 채널별 ROAS 순위로 방향성 자동 도출
            valid = {k:v for k,v in ch_insights.items() if v and v['광고비'] > 0}
            if valid:
                sorted_ch = sorted(valid.items(), key=lambda x: x[1]['ROAS'], reverse=True)
                best_ch = sorted_ch[0][0]
                worst_ch = sorted_ch[-1][0]
                best_roas = sorted_ch[0][1]['ROAS']
                worst_roas = sorted_ch[-1][1]['ROAS']
                total_spend = sum(v['광고비'] for v in valid.values())
                total_conv2 = sum(v['전환값'] for v in valid.values())
                overall_roas = total_conv2 / total_spend if total_spend > 0 else 0

                directions = [
                    f"**1. 예산 집중:** {best_ch} (ROAS {best_roas:.1f}x)이 가장 효율적입니다. 이번주 예산의 40~50%를 이 채널에 집중하세요.",
                    f"**2. 효율 개선 or 축소:** {worst_ch} (ROAS {worst_roas:.1f}x)는 효율이 가장 낮습니다. 소재를 교체하거나 예산을 줄이고 상위 채널로 이전하세요.",
                    f"**3. 전체 ROAS:** 이번주 통합 ROAS는 **{overall_roas:.1f}x**입니다. " +
                    ("목표치(3x) 달성 — 현재 운영 방향 유지하되 스케일업 시점입니다." if overall_roas >= 3
                     else "목표치(3x) 미달 — 채널 믹스 재조정이 필요합니다."),
                    "**4. 시즌 대응:** 현재 SS 시즌 후반 — 무신사 기획전 및 시즌오프 할인과 광고를 연계하면 ROAS 상승 가능성이 높습니다.",
                    "**5. 26FW 준비:** 8월부터 FW 신상품 티저 광고를 준비하세요. 특히 후리스/스웻팬츠 카테고리 타겟팅을 선점하면 유리합니다.",
                ]
                for d in directions:
                    st.markdown(f'<div class="insight-card" style="margin-bottom:6px">{d}</div>', unsafe_allow_html=True)
            else:
                st.info("이번주 광고 데이터가 없습니다.")

# ════════════════════════════════════════════════════════════════════════════
# TAB 8 — 인사이트
# ════════════════════════════════════════════════════════════════════════════
with tab8:
    st.markdown("### 💡 이번 주 매출 인사이트")
    st.caption("데이터 기반 자동 분석 · 20년 실무자가 봐야 할 시그널 중심")

    # ── 인사이트 카드 렌더러 ─────────────────────────────────────────────────
    def ins_card(tag, emoji, title, body_text, action=None):
        tag_colors = {
            "기회":  ("#e8f5e9", "#2e7d32", "🟢"),
            "리스크": ("#fff5f5", "#c62828", "🔴"),
            "주의":  ("#fff8e1", "#e65100", "⚠️"),
            "액션":  ("#e8f0fe", "#1a237e", "📌"),
            "정보":  ("#f0f6ff", "#1565c0", "ℹ️"),
        }
        bg, border, dot = tag_colors.get(tag, ("#f9f9f9","#888","•"))
        action_html = f'<div style="margin-top:6px;font-size:12px;color:#555;background:rgba(0,0,0,.04);padding:5px 8px;border-radius:5px;">→ {action}</div>' if action else ""
        st.markdown(
            f'<div style="background:{bg};border-left:4px solid {border};border-radius:8px;'
            f'padding:12px 16px;margin:6px 0;">'
            f'<div style="font-size:11px;font-weight:700;color:{border};letter-spacing:.5px;margin-bottom:4px;">'
            f'{dot} {tag.upper()}</div>'
            f'<div style="font-size:14px;font-weight:700;color:#111;margin-bottom:4px;">{title}</div>'
            f'<div style="font-size:13px;color:#444;line-height:1.6;">{body_text}</div>'
            f'{action_html}</div>',
            unsafe_allow_html=True
        )

    # ── 데이터 계산 ──────────────────────────────────────────────────────────
    # 주별 집계
    weekly_all = df.groupby('주차').agg(
        매출=('판매금액','sum'), 수량=('수량','sum')
    ).reset_index().sort_values('주차').reset_index(drop=True)

    # 일별 집계 (최근 28일)
    cutoff_28 = pd.Timestamp.now() - pd.Timedelta(days=28)
    df_28 = df[df['날짜'] >= cutoff_28]
    daily_28 = df_28.groupby('날짜')['판매금액'].sum().reset_index()
    daily_avg_28 = daily_28['판매금액'].mean() if len(daily_28) > 0 else 0

    # 최근 2주 vs 그 이전 2주 비교
    if len(weekly_all) >= 4:
        last2w_rev  = weekly_all.tail(2)['매출'].sum()
        prev2w_rev  = weekly_all.tail(4).head(2)['매출'].sum()
        wow2_pct    = (last2w_rev - prev2w_rev) / prev2w_rev * 100 if prev2w_rev > 0 else 0
    else:
        last2w_rev, prev2w_rev, wow2_pct = 0, 0, 0

    # 이동평균 (4주)
    if len(weekly_all) >= 4:
        ma4  = weekly_all.tail(4)['매출'].mean()
        ma8  = weekly_all.tail(8)['매출'].mean() if len(weekly_all) >= 8 else weekly_all['매출'].mean()
        trend_pct = (ma4 - ma8) / ma8 * 100 if ma8 > 0 else 0
        last_week_rev = weekly_all.iloc[-1]['매출']
        last_week_wow = (last_week_rev - weekly_all.iloc[-2]['매출']) / weekly_all.iloc[-2]['매출'] * 100 if len(weekly_all) >= 2 and weekly_all.iloc[-2]['매출'] > 0 else 0
    else:
        ma4, trend_pct, last_week_rev, last_week_wow = 0, 0, 0, 0

    # 카테고리 쏠림
    cat_rev = df.groupby('카테고리')['판매금액'].sum().sort_values(ascending=False)
    cat_pct = (cat_rev / cat_rev.sum() * 100).round(1)
    top_cat = cat_pct.index[0] if len(cat_pct) > 0 else "-"
    top_cat_pct = cat_pct.iloc[0] if len(cat_pct) > 0 else 0

    # 채널 의존도
    ch_rev = df.groupby('채널')['판매금액'].sum().sort_values(ascending=False)
    ch_pct = (ch_rev / ch_rev.sum() * 100).round(1)
    top_ch = ch_pct.index[0] if len(ch_pct) > 0 else "-"
    top_ch_pct = ch_pct.iloc[0] if len(ch_pct) > 0 else 0

    # 평균 할인율 추이
    if len(weekly_all) >= 2:
        last_wk_str = weekly_all.iloc[-1]['주차']
        prev_wk_str = weekly_all.iloc[-2]['주차']
        disc_last = df[df['주차'] == last_wk_str]['할인율'].mean()
        disc_prev = df[df['주차'] == prev_wk_str]['할인율'].mean()
        disc_delta = disc_last - disc_prev if not (np.isnan(disc_last) or np.isnan(disc_prev)) else 0
    else:
        disc_last, disc_delta = df['할인율'].mean(), 0

    # ASP (평균 판매단가) 추이
    if len(weekly_all) >= 2:
        asp_last = df[df['주차'] == last_wk_str]['판매단가'].mean()
        asp_prev = df[df['주차'] == prev_wk_str]['판매단가'].mean()
        asp_delta_pct = (asp_last - asp_prev) / asp_prev * 100 if asp_prev > 0 and not np.isnan(asp_prev) else 0
    else:
        asp_last, asp_delta_pct = df['판매단가'].mean(), 0

    # TOP3 상품 집중도
    top3_rev = cat_rev.head(3).sum()
    top3_pct = top3_rev / cat_rev.sum() * 100 if cat_rev.sum() > 0 else 0

    # 다음 주 예측 (최근 4주 평균 × 계절 보정 없음)
    forecast_low  = ma4 * 0.85
    forecast_high = ma4 * 1.15

    # ── 대시보드 헤더 수치 ────────────────────────────────────────────────────
    col_a, col_b, col_c, col_d = st.columns(4)
    kpi_card(col_a, "최근 주 매출", f"₩{last_week_rev/10000:,.0f}만",
             delta=last_week_wow)
    kpi_card(col_b, "4주 이동평균", f"₩{ma4/10000:,.0f}만",
             delta=trend_pct, color="#43A047")
    kpi_card(col_c, "일 평균 (28일)", f"₩{daily_avg_28/10000:,.0f}만",
             color="#FB8C00")
    kpi_card(col_d, "평균 할인율", f"{disc_last:.1f}%" if not np.isnan(disc_last) else "-",
             delta=disc_delta, color="#E53935")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # ── 인사이트 카드 ─────────────────────────────────────────────────────────
    col_l, col_r = st.columns(2)

    with col_l:
        st.markdown('<div class="sec-title">📡 신호 분석</div>', unsafe_allow_html=True)

        # ① 신호 vs 노이즈
        if abs(last_week_wow) >= 20:
            tag = "기회" if last_week_wow > 0 else "리스크"
            _cmp = "상회" if last_week_rev > ma4 else "하회"
            ins_card(tag, "", "유의미한 매출 변화 감지",
                f"전주 대비 <b>{last_week_wow:+.1f}%</b>. 단발성 이벤트 여부 확인 필요. "
                f"4주 이동평균(₩{ma4/10000:,.0f}만)과 비교 시 {_cmp}.",
                action="이벤트/행사 여부 확인 → 다음 주 지속 여부 판단")
        elif abs(last_week_wow) >= 10:
            ins_card("주의", "", "노이즈 경계 구간",
                f"전주 대비 <b>{last_week_wow:+.1f}%</b>. 아직 노이즈 범위지만 2주 연속 지속되면 신호.",
                action="다음 주 매출 주시 — 같은 방향이면 원인 분석 필요")
        else:
            _dir = "📈 상승" if trend_pct > 2 else ("📉 하락" if trend_pct < -2 else "→ 보합")
            ins_card("정보", "", "정상 범위 내 변동",
                f"전주 대비 <b>{last_week_wow:+.1f}%</b>. 4주 이동평균 대비 안정적. "
                f"트렌드 방향: {_dir} ({trend_pct:+.1f}%/4주).")

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        # ② 할인율 시그널
        if not np.isnan(disc_last):
            if disc_last > 20:
                ins_card("리스크", "", f"할인율 위험 수위 ({disc_last:.1f}%)",
                    f"평균 할인율 <b>{disc_last:.1f}%</b>. 전주 대비 {disc_delta:+.1f}%p. "
                    "쿠폰/프로모션 없이 팔리는 비중이 낮아지고 있음. 정가 경쟁력 점검 필요.",
                    action="할인 없는 주간의 베이스 매출 별도 추적 시작")
            elif disc_delta > 3:
                ins_card("주의", "", f"할인율 상승 추세 ({disc_last:.1f}%)",
                    f"전주 대비 <b>+{disc_delta:.1f}%p</b> 상승. 프로모션 의존도 증가 신호.",
                    action="행사 없는 주간 할인율과 비교 — 구조적 상승인지 확인")
            else:
                ins_card("정보", "", f"할인율 안정 ({disc_last:.1f}%)",
                    f"전주 대비 {disc_delta:+.1f}%p. 과도한 쿠폰 의존 없이 운영 중.")

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        # ③ ASP 추이
        if not np.isnan(asp_delta_pct):
            if asp_delta_pct < -5:
                ins_card("주의", "", f"평균 판매단가 하락 ({asp_delta_pct:+.1f}%)",
                    f"단가가 <b>{asp_delta_pct:.1f}%</b> 내려갔습니다. 저가 상품 믹스 증가 또는 할인 확대 신호.",
                    action="고단가 상품(PT/팬츠) 판매 비중 확인")
            elif asp_delta_pct > 5:
                ins_card("기회", "", f"평균 판매단가 상승 ({asp_delta_pct:+.1f}%)",
                    f"단가가 <b>+{asp_delta_pct:.1f}%</b> 올랐습니다. 고가 상품 판매 증가 또는 할인 축소 효과.")
            else:
                ins_card("정보", "", f"단가 안정 ({asp_delta_pct:+.1f}%)",
                    f"평균 판매단가 전주 대비 {asp_delta_pct:+.1f}%. 상품 믹스 변화 없음.")

    with col_r:
        st.markdown('<div class="sec-title">⚠️ 리스크 & 액션</div>', unsafe_allow_html=True)

        # ④ 카테고리 쏠림
        if top_cat_pct >= 85:
            ins_card("리스크", "", f"{top_cat} 쏠림 위험 ({top_cat_pct:.0f}%)",
                f"전체 매출의 <b>{top_cat_pct:.0f}%</b>가 {top_cat} 단일 카테고리. "
                "이 카테고리가 꺾이면 브랜드 전체가 꺾입니다.",
                action="2위 카테고리 육성 플랜 — 기획전/시딩 배분 조정 검토")
        elif top_cat_pct >= 70:
            ins_card("주의", "", f"{top_cat} 의존도 주의 ({top_cat_pct:.0f}%)",
                f"매출의 {top_cat_pct:.0f}%가 {top_cat}. 아직 관리 가능하나 추가 쏠림 방지 필요.",
                action="다음 기획전에서 2위 카테고리 노출 비중 의도적으로 높이기")
        else:
            ins_card("기회", "", f"카테고리 분산 양호 ({top_cat_pct:.0f}%)",
                f"1위 {top_cat}이 {top_cat_pct:.0f}%. 분산이 이루어지고 있음.")

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        # ⑤ 채널 의존도
        if top_ch_pct >= 80:
            ins_card("리스크", "", f"{top_ch} 채널 쏠림 ({top_ch_pct:.0f}%)",
                f"매출의 <b>{top_ch_pct:.0f}%</b>가 {top_ch}. "
                f"{top_ch} 알고리즘/정책 변경 시 즉각적 타격.",
                action="29CM / 자사몰 비중 늘리는 기획전 별도 운영")
        elif top_ch_pct >= 65:
            ins_card("주의", "", f"{top_ch} 의존도 높음 ({top_ch_pct:.0f}%)",
                f"{top_ch}이 {top_ch_pct:.0f}%. 단일 채널 의존 리스크 관리 시점.",
                action="이번 달 자사몰 기획전 1회 추가 운영 검토")
        else:
            ins_card("기회", "", f"채널 분산 진행 중 ({top_ch}: {top_ch_pct:.0f}%)",
                f"최상위 채널이 {top_ch_pct:.0f}%. 멀티채널 전략 효과 나오는 중.")

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        # ⑥ 다음 주 예측
        ins_card("액션", "", "다음 주 예상 매출 범위",
            f"최근 4주 이동평균 기준: <b>₩{forecast_low/10000:,.0f}만 ~ ₩{forecast_high/10000:,.0f}만</b>.<br>"
            "이벤트 없는 베이스 예측. 기획전 집행 시 +20~40% 가산 가능.",
            action=f"다음 주 목표를 ₩{(ma4*1.05)/10000:,.0f}만으로 설정하고 미팅에서 확인")

    # ── 주간 매출 트렌드 차트 ─────────────────────────────────────────────────
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="sec-title">📈 주간 매출 트렌드 (최근 12주)</div>', unsafe_allow_html=True)
    if len(weekly_all) > 0:
        plot_w = weekly_all.tail(12).copy()
        plot_w['매출_만'] = (plot_w['매출'] / 10000).round(1)
        plot_w['MA4'] = plot_w['매출'].rolling(4, min_periods=1).mean() / 10000

        fig_w = go.Figure()
        fig_w.add_bar(x=plot_w['주차'], y=plot_w['매출_만'],
                      name='주간 매출', marker_color='#90CAF9', opacity=0.85)
        fig_w.add_scatter(x=plot_w['주차'], y=plot_w['MA4'],
                          name='4주 이동평균', mode='lines+markers',
                          line=dict(color='#1E88E5', width=2.5),
                          marker=dict(size=6))
        fig_w.update_layout(height=260, margin=dict(t=10,b=10,l=0,r=0),
                            yaxis_title="만원", xaxis_tickangle=-30,
                            legend=dict(orientation='h', y=1.15),
                            plot_bgcolor='white', paper_bgcolor='white')
        fig_w.update_yaxes(gridcolor='#f0f0f0')
        st.plotly_chart(fig_w, use_container_width=True)

    # ── 카테고리 & 채널 구성 ─────────────────────────────────────────────────
    col_pie1, col_pie2 = st.columns(2)
    with col_pie1:
        st.markdown('<div class="sec-title">카테고리 구성</div>', unsafe_allow_html=True)
        fig_cat = px.pie(
            values=cat_rev.values, names=cat_rev.index,
            hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_cat.update_layout(height=220, margin=dict(t=0,b=0,l=0,r=0),
                               showlegend=True, legend=dict(font=dict(size=11)))
        fig_cat.update_traces(textposition='inside', textinfo='percent+label',
                               textfont_size=11)
        st.plotly_chart(fig_cat, use_container_width=True)
    with col_pie2:
        st.markdown('<div class="sec-title">채널 구성</div>', unsafe_allow_html=True)
        fig_ch = px.pie(
            values=ch_rev.values, names=ch_rev.index,
            hole=0.5, color_discrete_map=CH_COLORS
        )
        fig_ch.update_layout(height=220, margin=dict(t=0,b=0,l=0,r=0),
                              showlegend=True, legend=dict(font=dict(size=11)))
        fig_ch.update_traces(textposition='inside', textinfo='percent+label',
                              textfont_size=11)
        st.plotly_chart(fig_ch, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 9 — 관리
# ════════════════════════════════════════════════════════════════════════════
with tab9:
    sub1,sub2,sub3,sub4 = st.tabs(["🎯 월별 목표","📦 상품 메모","✅ TODO","📥 매출 업데이트"])

    # ── 서브탭 1: 월별 목표 ──────────────────────────────────────────────────
    with sub1:
        st.markdown("### 🎯 월별 매출 목표 관리")
        goal_data = load_goals()
        months_list = sorted(df_all['년월'].unique(), reverse=True) if '년월' in df_all.columns else []

        if months_list:
            sel_month = st.selectbox("월 선택", months_list, key="goal_month")
            actual = df_all[df_all['년월'] == sel_month]['판매금액'].sum()
            # targets는 만원 단위로 저장됨 (홈탭: * 10000 해서 원으로 변환)
            cur_goal_man = goal_data.get(sel_month, 0)           # 만원
            cur_goal_won = cur_goal_man * 10000                   # 원
            achievement = actual / cur_goal_won * 100 if cur_goal_won > 0 else 0

            c1, c2, c3 = st.columns(3)
            kpi_card(c1, "실적", f"₩{actual/10000:,.0f}만")
            kpi_card(c2, "목표", f"₩{cur_goal_man:,}만", color="#1565C0")
            kpi_card(c3, "달성률", f"{achievement:.1f}%",
                     delta=achievement-100, color="#2E7D32" if achievement>=100 else "#C62828")

            st.markdown("---")
            new_goal_man = st.number_input("목표 금액 (만원)", value=int(cur_goal_man), step=100,
                                           format="%d", key="new_goal_input")
            if st.button("💾 저장", key="save_goal"):
                goal_data[sel_month] = new_goal_man   # 만원 단위로 저장
                save_goals(goal_data)
                st.success(f"{sel_month} 목표 ₩{new_goal_man:,}만 저장 완료!")
                st.rerun()

            # 전체 월 목표 현황표
            st.markdown("#### 전체 월 목표 현황")
            goal_rows = []
            for m in months_list:
                act = df_all[df_all['년월'] == m]['판매금액'].sum()
                g_man = goal_data.get(m, 0)           # 만원
                ach = act / (g_man * 10000) * 100 if g_man > 0 else 0
                goal_rows.append({"월": m, "실적(만)": f"{act/10000:,.0f}", "목표(만)": f"{g_man:,}", "달성률": f"{ach:.1f}%"})
            st.dataframe(pd.DataFrame(goal_rows), use_container_width=True, hide_index=True)
        else:
            st.info("데이터를 먼저 업로드하세요.")

    # ── 서브탭 2: 상품 메모 ──────────────────────────────────────────────────
    with sub2:
        st.markdown("### 📦 상품별 메모")
        memo_data = load_memos()
        products = sorted(df_all['상품명'].dropna().unique()) if '상품명' in df_all.columns else []

        if products:
            sel_prod = st.selectbox("상품 선택", products, key="memo_prod")
            cur_memo = memo_data.get(sel_prod, "")
            new_memo = st.text_area("메모", value=cur_memo, height=150, key="memo_text")
            if st.button("💾 메모 저장", key="save_memo"):
                memo_data[sel_prod] = new_memo
                save_memos(memo_data)
                st.success("저장 완료!")
            if cur_memo:
                st.markdown(f'<div class="insight-card">{cur_memo}</div>', unsafe_allow_html=True)
        else:
            st.info("데이터를 먼저 업로드하세요.")

    # ── 서브탭 3: TODO ────────────────────────────────────────────────────────
    with sub3:
        st.markdown("### ✅ TODO 리스트")
        todo_data = load_todos()
        new_todo = st.text_input("새 TODO 입력 후 엔터", key="new_todo")
        if new_todo:
            todo_data.append({"text": new_todo, "done": False})
            save_todos(todo_data)
            st.rerun()

        # 기존 데이터 형식 정규화 (str → dict)
        todo_data = [
            item if isinstance(item, dict) else {"text": str(item), "done": False}
            for item in todo_data
        ]

        for i, item in enumerate(todo_data):
            col_chk, col_txt, col_del = st.columns([0.1, 0.8, 0.1])
            with col_chk:
                done = st.checkbox("", value=item.get("done", False), key=f"todo_{i}")
                if done != item.get("done", False):
                    todo_data[i]["done"] = done
                    save_todos(todo_data)
                    st.rerun()
            with col_txt:
                style = "text-decoration:line-through;color:#aaa;" if item.get("done") else ""
                st.markdown(f'<span style="{style}">{item.get("text","")}</span>', unsafe_allow_html=True)
            with col_del:
                if st.button("🗑", key=f"del_todo_{i}"):
                    todo_data.pop(i)
                    save_todos(todo_data)
                    st.rerun()

    # ── 서브탭 4: 매출 업데이트 ──────────────────────────────────────────────
    with sub4:
        st.markdown("### 📥 매출 데이터 업데이트")
        st.info("새 매출 파일(Excel)을 업로드하면 기존 데이터에 병합됩니다.")
        uploaded = st.file_uploader("Excel 파일 업로드", type=["xlsx","xls"], key="upload_new")
        if uploaded:
            try:
                df_new = pd.read_excel(uploaded)
                st.write(f"업로드된 행 수: {len(df_new):,}")
                st.dataframe(df_new.head(), use_container_width=True)
                if st.button("병합 & 저장", key="merge_btn"):
                    df_merged = pd.concat([df_all, df_new], ignore_index=True)
                    df_merged.drop_duplicates(inplace=True)
                    df_merged.to_excel(filepath, index=False)
                    st.success(f"병합 완료! 총 {len(df_merged):,}행 저장됨.")
                    st.rerun()
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

