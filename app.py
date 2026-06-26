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

    def calc_real(row):
        if pd.isna(row['택가']) or row['택가'] <= 0: return np.nan
        if pd.isna(row['판매단가']): return np.nan
        base = row['택가'] * 1.2 if row['채널'] == '무신사글로벌' else row['택가']
        return (row['판매단가'] / base) * 100
    df['실현율'] = df.apply(calc_real, axis=1)
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
                "🔬 상품분석", "🏪 팝업", "📣 마케팅", "⚙️ 관리"])
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = tabs

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — 홈
# ════════════════════════════════════════════════════════════════════════════
with tab1:
    mgmt_home = load_mgmt()
    targets = mgmt_home.get("targets", {})

    # 월별 KPI 계산
    monthly = df.groupby('년월').agg(매출=('판매금액','sum'),수량=('수량','sum'),SKU=('상품명','nunique')).reset_index()
    monthly['ASP'] = monthly['매출'] / monthly['수량']
    monthly['실현율'] = df.groupby('년월')['실현율'].mean().reindex(monthly['년월']).values
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
    ch_m = df.groupby(['년월','채널'])['판매금액'].sum().reset_index()
    ch_m['매출_만'] = (ch_m['판매금액']/10000).round()
    ch_t = df.groupby('채널').agg(매출=('판매금액','sum'),수량=('수량','sum'),실현율=('실현율','mean')).reset_index()
    ch_t['매출_만'] = (ch_t['매출']/10000).round()
    ch_t['비중'] = (ch_t['매출']/ch_t['매출'].sum()*100).round(1)
    ch_t['ASP'] = (ch_t['매출']/ch_t['수량']/1000).round(1)

    col_a,col_b = st.columns([2,1])
    with col_a:
        fig = px.bar(ch_m, x='년월', y='매출_만', color='채널', title='채널별 월별 매출',
                     color_discrete_map=CH_COLORS)
        fig.update_layout(height=360, plot_bgcolor='white', xaxis=dict(tickangle=-45),
                          legend=dict(orientation='h',y=1.1))
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        fig = px.pie(ch_t, values='매출_만', names='채널', hole=0.4,
                     color='채널', color_discrete_map=CH_COLORS, title='채널 비중')
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(height=360, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    fig = go.Figure()
    for _, row in ch_t.iterrows():
        if pd.notna(row['실현율']):
            fig.add_bar(x=[row['채널']], y=[round(row['실현율'],1)],
                        marker_color=CH_COLORS.get(row['채널'],'#999'),
                        text=[f"{row['실현율']:.1f}%"], textposition='outside')
    fig.update_layout(title='채널별 실현율', yaxis=dict(range=[50,105]),
                      plot_bgcolor='white', showlegend=False, height=260)
    st.plotly_chart(fig, use_container_width=True)

    # 채널 간 카니발리제이션
    with st.expander("🔀 채널 카니발리제이션 분석"):
        ch_monthly_wide = df_all.groupby(['년월','채널'])['판매금액'].sum().unstack(fill_value=0).reset_index()
        for ch in CHANNELS:
            if ch not in ch_monthly_wide.columns: ch_monthly_wide[ch] = 0
        ch_monthly_wide = ch_monthly_wide.sort_values('년월')
        musinsa_real = df_all[df_all['채널']=='무신사'].groupby('년월')['실현율'].mean().reset_index()
        musinsa_real.columns = ['년월','무신사_실현율']
        ch_monthly_wide = ch_monthly_wide.merge(musinsa_real, on='년월', how='left')

        merged = ch_monthly_wide.dropna(subset=['무신사_실현율'])
        if len(merged) > 3 and '자사몰' in merged.columns:
            corr = merged['무신사_실현율'].corr(merged['자사몰'])
            col_x, col_y = st.columns(2)
            with col_x:
                st.metric("무신사 실현율 ↔ 자사몰 상관계수", f"{corr:.3f}")
                if corr < -0.3:
                    st.markdown('<div class="danger-card">🔴 카니발리제이션 감지</div>', unsafe_allow_html=True)
                elif corr > 0.3:
                    st.markdown('<div class="good-card">🟢 헤일로 효과</div>', unsafe_allow_html=True)
                else:
                    st.info("채널 간 독립적 움직임")
            with col_y:
                ch_for_corr = [ch for ch in CHANNELS if ch in ch_monthly_wide.columns and ch_monthly_wide[ch].sum()>0]
                if len(ch_for_corr) >= 2:
                    corr_matrix = ch_monthly_wide[ch_for_corr].corr().round(2)
                    fig2 = px.imshow(corr_matrix, color_continuous_scale='RdBu_r', zmin=-1, zmax=1, text_auto=True)
                    fig2.update_layout(height=280)
                    st.plotly_chart(fig2, use_container_width=True)

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
# TAB 8 — 관리
# ════════════════════════════════════════════════════════════════════════════
with tab8:
    st.markdown("### ⚙️ 관리")
    mgmt = load_mgmt()
    sub1,sub2,sub3,sub4 = st.tabs(["🎯 월별 목표","📦 상품 메모","✅ TODO","📥 매출 업데이트"])

    with sub1:
        st.markdown("**월별 매출 목표 설정 (만원)**")
        targets2 = mgmt.get("targets",{})
        cols = st.columns(3)
        new_targets = {}
        for i,ym in enumerate(months_all):
            with cols[i%3]:
                actual = df_all[df_all['년월']==ym]['판매금액'].sum()/10000
                val = st.number_input(f"{ym}", value=int(targets2.get(ym, round(actual*1.1))), step=100, key=f"t_{ym}")
                new_targets[ym] = val
                ach = actual/val*100 if val>0 else 0
                color = "green" if ach>=100 else ("orange" if ach>=80 else "red")
                st.markdown(f"<span style='color:{color};font-size:12px'>달성 {ach:.1f}% (실적 {actual:,.0f}만)</span>", unsafe_allow_html=True)
        if st.button("💾 목표 저장", type="primary"):
            mgmt['targets'] = {k:int(v) for k,v in new_targets.items()}
            save_mgmt(mgmt); st.success("저장됐어요!")
        if new_targets:
            tgt_df = pd.DataFrame([{'년월':ym,'목표':new_targets.get(ym,0),'실적':df_all[df_all['년월']==ym]['판매금액'].sum()/10000} for ym in months_all])
            fig = go.Figure()
            fig.add_bar(x=tgt_df['년월'], y=tgt_df['실적'], name='실적', marker_color='#1E88E5')
            fig.add_scatter(x=tgt_df['년월'], y=tgt_df['목표'], name='목표', mode='lines+markers', line=dict(color='#E53935',dash='dash',width=2))
            fig.update_layout(title='목표 vs 실적', height=280, plot_bgcolor='white',
                              legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45))
            st.plotly_chart(fig, use_container_width=True)

    with sub2:
        st.markdown("**상품별 재고 현황 / 메모**")
        stock_notes = mgmt.get("stock_notes",{})
        style_list = df_all.groupby(['스타일코드','대표명'])['판매금액'].sum().sort_values(ascending=False).reset_index()['대표명'].head(60).tolist()
        sel = st.selectbox("상품 선택", style_list)
        existing = stock_notes.get(sel,{"재고":"","상태":"정상","메모":""})
        c1,c2 = st.columns(2)
        new_stock = c1.text_input("재고 현황", value=existing.get("재고",""), placeholder="예: 무신사 30개 / 자사몰 10개")
        new_status = c2.selectbox("상태", ["정상","재고 부족 주의","재고 소진 임박","단종","재출시 검토"],
                                   index=["정상","재고 부족 주의","재고 소진 임박","단종","재출시 검토"].index(existing.get("상태","정상")))
        new_memo = st.text_area("메모", value=existing.get("메모",""), height=100)
        if st.button("💾 저장", type="primary"):
            stock_notes[sel] = {"재고":new_stock,"상태":new_status,"메모":new_memo,"수정일":datetime.now().strftime("%Y-%m-%d %H:%M")}
            mgmt['stock_notes'] = stock_notes; save_mgmt(mgmt); st.success("저장됐어요!")
        if stock_notes:
            st.markdown("---")
            memo_df = pd.DataFrame([{"상품":k,"재고":v.get("재고",""),"상태":v.get("상태",""),"메모":v.get("메모",""),"수정일":v.get("수정일","")} for k,v in stock_notes.items()])
            st.dataframe(memo_df, use_container_width=True, hide_index=True)

    with sub3:
        todos = mgmt.get("todos",[])
        c1,c2,c3 = st.columns([3,1,1])
        new_task = c1.text_input("새 할 일", placeholder="예: MA-1 자켓 재고 확인")
        new_priority = c2.selectbox("우선순위", ["🔴 즉시","🟡 이번달","🟢 26FW"])
        c3.markdown("<br>", unsafe_allow_html=True)
        if c3.button("➕ 추가") and new_task:
            todos.append({"task":new_task,"priority":new_priority,"done":False,"added":datetime.now().strftime("%Y-%m-%d")})
            mgmt['todos'] = todos; save_mgmt(mgmt); st.rerun()
        for priority in ["🔴 즉시","🟡 이번달","🟢 26FW"]:
            items = [t for t in todos if t.get('priority')==priority]
            if items:
                st.markdown(f"**{priority}**")
                for i,todo in enumerate(todos):
                    if todo.get('priority')!=priority: continue
                    cl1,cl2 = st.columns([10,1])
                    done = cl1.checkbox(todo['task'], value=todo.get('done',False), key=f"todo_{i}")
                    if done != todo.get('done',False):
                        todos[i]['done'] = done; mgmt['todos'] = todos; save_mgmt(mgmt)
                    if cl2.button("🗑️", key=f"del_{i}"):
                        todos.pop(i); mgmt['todos'] = todos; save_mgmt(mgmt); st.rerun()
        done_items = [t for t in todos if t.get('done')]
        if done_items:
            with st.expander(f"✅ 완료 ({len(done_items)}개)"):
                for t in done_items: st.markdown(f"~~{t['task']}~~")

    with sub4:
        import subprocess, sys
        st.markdown("**📥 채널별 로우파일 → 통합 매출 자동 업데이트**")
        st.markdown("""
        `ERP/raw/` 폴더에 아래 파일을 넣고 버튼을 누르세요.

        | 채널 | 파일명 |
        |------|--------|
        | 무신사 | `무신사.xls` |
        | 무신사글로벌 | `무신사글로벌.xls` |
        | 자사몰 | `자사몰.csv` |
        | 29CM | `29cm.xlsx` |
        """)
        raw_dir = Path(__file__).parent / "raw"
        raw_files = list(raw_dir.glob("*")) if raw_dir.exists() else []
        if raw_files:
            st.success(f"raw 폴더 내 파일: {', '.join([f.name for f in raw_files])}")
        else:
            st.warning("raw 폴더가 비어 있거나 없습니다.")

        if st.button("🔄 지금 업데이트 실행", type="primary"):
            merge_script = Path(__file__).parent / "merge_raw.py"
            if not merge_script.exists():
                st.error("merge_raw.py 파일이 없습니다.")
            else:
                with st.spinner("실행 중..."):
                    result = subprocess.run(
                        [sys.executable, str(merge_script)],
                        capture_output=True, text=True, encoding='utf-8', errors='replace'
                    )
                if result.returncode == 0:
                    st.success("✅ 업데이트 완료!")
                    st.code(result.stdout or "(출력 없음)")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("❌ 오류 발생")
                    st.code(result.stderr or result.stdout)
