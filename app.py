import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats
import json
import os
from pathlib import Path
from datetime import datetime, timedelta

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
        if pw == st.secrets.get("password", ""):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False

if not check_password():
    st.stop()

BASE_DIR = Path(__file__).parent
# 로컬: ERP 폴더 기준 / 클라우드: app.py 옆 data 폴더
DATA_DIR = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR.parent / "매출자료"
MGMT_FILE = BASE_DIR / "management.json"
MKTG_FILE = BASE_DIR / "루저스클럽 마케팅 효율.xlsx"

CHANNELS = ['무신사', '자사몰', '무신사글로벌', '29CM', '팝업(더현대)']
CH_COLORS = {'무신사':'#1E88E5','자사몰':'#43A047','무신사글로벌':'#FB8C00','29CM':'#E53935','팝업(더현대)':'#8E24AA'}
CAT_MAP = {'PT':'팬츠','TS':'반팔티','BG':'가방','HZ':'후리스집업','SW':'스웻/후드','JK':'자켓',
           'JW':'주얼리','CP':'캡/비니','SP':'스웻팬츠','DP':'데님팬츠','KR':'크루넥','SL':'슬리브리스','AC':'악세서리'}

st.markdown("""
<style>
.metric-card{background:#f8f9fa;border-radius:12px;padding:16px 20px;border-left:4px solid #1E88E5;margin-bottom:8px;}
.metric-label{font-size:12px;color:#666;font-weight:500;}
.metric-value{font-size:26px;font-weight:700;color:#1a1a1a;line-height:1.2;}
.metric-delta{font-size:13px;margin-top:2px;}
.up{color:#43A047;}.dn{color:#E53935;}
.warn-card{background:#FFF3E0;border-radius:8px;padding:10px 14px;border-left:4px solid #FB8C00;margin:4px 0;}
.danger-card{background:#FFEBEE;border-radius:8px;padding:10px 14px;border-left:4px solid #E53935;margin:4px 0;}
.good-card{background:#E8F5E9;border-radius:8px;padding:10px 14px;border-left:4px solid #43A047;margin:4px 0;}
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
            delta_html = f'<div class="metric-delta {cls}">{arrow} {abs(delta):.1f}%</div>'
        st.markdown(f'<div class="metric-card" style="border-left-color:{color}"><div class="metric-label">{label}</div><div class="metric-value">{value}</div>{delta_html}</div>', unsafe_allow_html=True)

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

    # 무신사 메타 광고
    df = pd.read_excel(xf, sheet_name='무신사 메타 광고', header=1)
    df = df.rename(columns={'Unnamed: 0':'_'})
    df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
    df = df.dropna(subset=['날짜'])
    df['년월'] = df['날짜'].dt.to_period('M').astype(str)
    df['ROAS'] = pd.to_numeric(df['catalog_segment_value_omni_purchase_roas:omni_purchase'], errors='coerce')
    df['지출 금액 (KRW)'] = pd.to_numeric(df['지출 금액 (KRW)'], errors='coerce').fillna(0)
    df['링크 클릭'] = pd.to_numeric(df['링크 클릭'], errors='coerce').fillna(0)
    df['노출'] = pd.to_numeric(df['노출'], errors='coerce').fillna(0)
    df['공유 항목이 포함된 구매'] = pd.to_numeric(df['공유 항목이 포함된 구매'], errors='coerce').fillna(0)
    df['공유 항목의 구매 전환값'] = pd.to_numeric(df['공유 항목의 구매 전환값'], errors='coerce').fillna(0)
    result['meta_ms'] = df

    # 자사몰 메타 광고
    df2 = pd.read_excel(xf, sheet_name='자사몰 메타 광고', header=1)
    df2['날짜'] = pd.to_datetime(df2['날짜'], errors='coerce')
    df2 = df2.dropna(subset=['날짜'])
    df2['년월'] = df2['날짜'].dt.to_period('M').astype(str)
    for col in ['지출 금액 (KRW)','링크 클릭','노출','구매','구매 전환값']:
        df2[col] = pd.to_numeric(df2[col], errors='coerce').fillna(0)
    result['meta_js'] = df2

    # 무신사 상품광고
    df3 = pd.read_excel(xf, sheet_name='무신사 상품광고', header=4)
    df3 = df3.dropna(subset=['날짜'])
    df3['날짜'] = pd.to_datetime(df3['날짜'], errors='coerce')
    df3['년월'] = df3['날짜'].dt.to_period('M').astype(str)
    for col in ['집행 광고비','매출','광고 수익률(ROAS)','클릭률','전환율','노출 수','클릭 수']:
        df3[col] = pd.to_numeric(df3[col], errors='coerce').fillna(0)
    result['prod_ms'] = df3

    # 29CM 상품광고
    df4 = pd.read_excel(xf, sheet_name='29CM 상품광고', header=4)
    df4 = df4.dropna(subset=['날짜'])
    df4['날짜'] = pd.to_datetime(df4['날짜'], errors='coerce')
    df4['년월'] = df4['날짜'].dt.to_period('M').astype(str)
    for col in ['집행 광고비','매출','광고 수익률(ROAS)','클릭률','전환율','노출 수','클릭 수']:
        df4[col] = pd.to_numeric(df4[col], errors='coerce').fillna(0)
    result['prod_29'] = df4

    return result

# ── 탭 ───────────────────────────────────────────────────────────────────────
tabs = st.tabs(["📊 KPI", "🏷️ 카테고리", "📺 채널", "🏆 상품랭킹",
                "🔬 수명주기", "⏱️ 재고소진예측", "🚨 이상치감지",
                "📈 시즌보정", "🏪 팝업분석", "🔀 카니발리제이션",
                "📣 마케팅", "⚙️ 관리"])
tab1,tab2,tab3,tab4,tab5,tab6,tab7,tab8,tab9,tab10,tab11,tab12 = tabs

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — KPI
# ════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### 📊 월별 KPI")
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

    last = monthly.iloc[-1]; prev = monthly.iloc[-2] if len(monthly)>1 else None
    c1,c2,c3,c4,c5 = st.columns(5)
    kpi_card(c1, f"총매출 ({last['년월']})", f"{last['매출_만']:,.0f}만", last['MoM'])
    kpi_card(c2, "총수량", f"{int(last['수량']):,}개", color="#43A047")
    kpi_card(c3, "ASP", f"{last['ASP']/1000:.1f}천원", color="#FB8C00")
    kpi_card(c4, "실현율", f"{last['실현율']:.1f}%", delta=last['실현율']-(prev['실현율'] if prev is not None else last['실현율']), color="#8E24AA")
    kpi_card(c5, "활성 SKU", f"{int(last['SKU'])}개", color="#E53935")

    col_a, col_b = st.columns([2,1])
    with col_a:
        fig = go.Figure()
        fig.add_bar(x=monthly['년월'], y=monthly['매출_만'], name='매출(만)', marker_color='#1E88E5', opacity=0.8)
        fig.add_scatter(x=monthly['년월'], y=monthly['실현율'], name='실현율(%)', yaxis='y2',
                        line=dict(color='#E53935',width=2), mode='lines+markers')
        fig.update_layout(title="월별 매출 & 실현율", yaxis=dict(title="매출(만)"),
                          yaxis2=dict(title="실현율(%)",overlaying='y',side='right',range=[50,100]),
                          legend=dict(orientation='h',y=1.1), height=340,
                          plot_bgcolor='white', xaxis=dict(tickangle=-45))
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        colors = ['#43A047' if v>=0 else '#E53935' for v in monthly['MoM'].fillna(0)]
        fig2 = go.Figure()
        fig2.add_bar(x=monthly['년월'], y=monthly['MoM'].round(1), marker_color=colors, name='MoM(%)')
        if monthly['YoY'].notna().any():
            fig2.add_scatter(x=monthly['년월'], y=monthly['YoY'].round(1), name='YoY(%)',
                             mode='lines+markers', line=dict(color='#FB8C00',width=2,dash='dot'))
        fig2.update_layout(title="MoM / YoY", height=340, plot_bgcolor='white',
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
# TAB 5 — 수명주기 (PLC)
# ════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown("### 🔬 상품 수명주기 (PLC) 자동 분류")
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

        # 선형 기울기
        all_m = months_sorted
        y = [grp[grp['년월']==m]['판매금액'].sum() if m in grp['년월'].values else 0 for m in all_m]
        if len(y) >= 3 and sum(y) > 0:
            slope, _, _, _, _ = stats.linregress(range(len(y)), y)
        else:
            slope = 0

        trend_pct = (recent_3-prev_3)/prev_3*100 if prev_3>0 else 0
        total_rev = grp['판매금액'].sum()

        # PLC 분류
        if months_active <= 3:
            plc = '🌱 도입'
        elif slope > 50000 and trend_pct > 10:
            plc = '🚀 성장'
        elif abs(slope) <= 50000 and total_rev > 5000000:
            plc = '💪 성숙'
        elif slope < -50000 or trend_pct < -30:
            plc = '📉 쇠퇴'
        else:
            plc = '➡️ 유지'

        plc_rows.append({'스타일코드':sc,'상품명':nm,'PLC':plc,'총매출_만':round(total_rev/10000),
                         '판매기간(월)':months_active,'최근3개월_만':round(recent_3/10000),
                         '추세변화(%)':round(trend_pct,1),'기울기':round(slope)})

    plc_df = pd.DataFrame(plc_rows).sort_values('총매출_만', ascending=False)

    # PLC 요약
    plc_counts = plc_df['PLC'].value_counts()
    cols = st.columns(5)
    for i, (label, color) in enumerate([('🌱 도입','#43A047'),('🚀 성장','#1E88E5'),
                                         ('💪 성숙','#FB8C00'),('📉 쇠퇴','#E53935'),('➡️ 유지','#757575')]):
        cnt = plc_counts.get(label, 0)
        kpi_card(cols[i], label, f"{cnt}개", color=color)

    # 산점도
    fig = px.scatter(plc_df, x='판매기간(월)', y='최근3개월_만', size='총매출_만',
                     color='PLC', hover_data=['상품명','추세변화(%)'],
                     title='PLC 포지션 맵 (크기=누적 매출)',
                     color_discrete_map={'🌱 도입':'#43A047','🚀 성장':'#1E88E5',
                                         '💪 성숙':'#FB8C00','📉 쇠퇴':'#E53935','➡️ 유지':'#9E9E9E'})
    fig.update_layout(height=400, plot_bgcolor='white')
    st.plotly_chart(fig, use_container_width=True)

    # 필터 테이블
    plc_filter = st.multiselect("PLC 단계 필터", ['🌱 도입','🚀 성장','💪 성숙','📉 쇠퇴','➡️ 유지'],
                                 default=['🚀 성장','📉 쇠퇴'])
    filtered_plc = plc_df[plc_df['PLC'].isin(plc_filter)] if plc_filter else plc_df
    st.dataframe(filtered_plc[['상품명','PLC','총매출_만','판매기간(월)','최근3개월_만','추세변화(%)']],
                 use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 6 — 재고 소진 예측
# ════════════════════════════════════════════════════════════════════════════
with tab6:
    st.markdown("### ⏱️ 재고 소진 예측")
    st.caption("최근 4주 판매 속도 vs 이전 4주 비교 — 급감 = 재고 소진 신호")

    last_date = df_all['날짜'].max()
    w4_start = last_date - timedelta(weeks=4)
    w8_start = last_date - timedelta(weeks=8)

    recent4 = df_all[df_all['날짜'] > w4_start].groupby(['스타일코드','대표명']).agg(
        최근4주_수량=('수량','sum'), 최근4주_매출=('판매금액','sum'),
        마지막판매=('날짜','max')).reset_index()
    prev4 = df_all[(df_all['날짜'] > w8_start) & (df_all['날짜'] <= w4_start)].groupby('스타일코드').agg(
        이전4주_수량=('수량','sum')).reset_index()

    inv_df = recent4.merge(prev4, on='스타일코드', how='left')
    inv_df['이전4주_수량'] = inv_df['이전4주_수량'].fillna(0)
    inv_df['수량변화(%)'] = np.where(inv_df['이전4주_수량']>0,
        (inv_df['최근4주_수량']-inv_df['이전4주_수량'])/inv_df['이전4주_수량']*100, 0)
    inv_df['경과일'] = (last_date - inv_df['마지막판매']).dt.days

    def risk_level(row):
        if row['경과일'] > 14 and row['최근4주_수량'] == 0: return '🔴 소진 의심'
        if row['수량변화(%)'] < -60 or (row['경과일'] > 7 and row['최근4주_수량'] < 3): return '🟡 소진 임박'
        if row['수량변화(%)'] < -30: return '🟠 감소 주의'
        return '🟢 정상'
    inv_df['리스크'] = inv_df.apply(risk_level, axis=1)
    inv_df['최근4주_매출_만'] = (inv_df['최근4주_매출']/10000).round(1)
    inv_df['마지막판매'] = inv_df['마지막판매'].dt.strftime('%Y-%m-%d')
    inv_df = inv_df.sort_values('수량변화(%)')

    # 요약
    rc = inv_df['리스크'].value_counts()
    c1,c2,c3,c4 = st.columns(4)
    kpi_card(c1, "🔴 소진 의심", f"{rc.get('🔴 소진 의심',0)}개", color="#E53935")
    kpi_card(c2, "🟡 소진 임박", f"{rc.get('🟡 소진 임박',0)}개", color="#FB8C00")
    kpi_card(c3, "🟠 감소 주의", f"{rc.get('🟠 감소 주의',0)}개", color="#FF6F00")
    kpi_card(c4, "🟢 정상", f"{rc.get('🟢 정상',0)}개", color="#43A047")

    st.markdown("#### 위험 상품 목록")
    danger = inv_df[inv_df['리스크'] != '🟢 정상']
    for _, row in danger.iterrows():
        cls = 'danger-card' if '🔴' in row['리스크'] else 'warn-card'
        st.markdown(f'<div class="{cls}"><b>{row["리스크"]} {row["대표명"]}</b> — 최근4주 {int(row["최근4주_수량"])}개 (이전 대비 {row["수량변화(%)"]:.0f}%) | 마지막판매: {row["마지막판매"]} ({int(row["경과일"])}일 전)</div>', unsafe_allow_html=True)

    # 전체 차트
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

# ════════════════════════════════════════════════════════════════════════════
# TAB 7 — 이상치 감지
# ════════════════════════════════════════════════════════════════════════════
with tab7:
    st.markdown("### 🚨 이상치 자동 감지")
    st.caption("전주 vs 4주 평균 대비 ±30% 이상 급변 상품 자동 감지")

    weekly = df_all.groupby(['주차','스타일코드','대표명'])['판매금액'].sum().reset_index()
    weeks_sorted = sorted(df_all['주차'].unique())
    if len(weeks_sorted) < 2:
        st.warning("주차 데이터가 부족합니다.")
    else:
        last_w = weeks_sorted[-1]
        prev_4w = weeks_sorted[max(0,len(weeks_sorted)-5):-1]

        last_week = weekly[weekly['주차']==last_w].set_index('스타일코드')['판매금액']
        avg_4w = weekly[weekly['주차'].isin(prev_4w)].groupby('스타일코드')['판매금액'].mean()

        anomaly_rows = []
        for sc in set(last_week.index) | set(avg_4w.index):
            lw = last_week.get(sc, 0)
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

        # 주간 전체 매출 추이
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
# TAB 8 — 시즌 보정 성과
# ════════════════════════════════════════════════════════════════════════════
with tab8:
    st.markdown("### 📈 시즌 보정 실질 성과")
    st.caption("SS/FW 시즌성을 제거한 보정 성장률 — 진짜 성장인지 시즌 효과인지 구분")

    monthly_all = df_all.groupby(['년월','월'])['판매금액'].sum().reset_index()
    monthly_all['매출_만'] = (monthly_all['판매금액']/10000).round()

    # 월별 시즌 인덱스 (전체 평균 대비)
    month_avg = monthly_all.groupby('월')['판매금액'].mean()
    overall_avg = monthly_all['판매금액'].mean()
    season_idx = (month_avg / overall_avg).round(3)

    monthly_all['시즌인덱스'] = monthly_all['월'].map(season_idx)
    monthly_all['보정매출_만'] = (monthly_all['매출_만'] / monthly_all['시즌인덱스']).round()

    fig = go.Figure()
    fig.add_scatter(x=monthly_all['년월'], y=monthly_all['매출_만'], name='실제 매출',
                    line=dict(color='#90CAF9',width=2), mode='lines+markers')
    fig.add_scatter(x=monthly_all['년월'], y=monthly_all['보정매출_만'], name='시즌 보정 매출',
                    line=dict(color='#1E88E5',width=2,dash='dash'), mode='lines+markers')
    fig.update_layout(title='실제 매출 vs 시즌 보정 매출', height=350,
                      plot_bgcolor='white', yaxis_title='매출(만)',
                      legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45))
    st.plotly_chart(fig, use_container_width=True)

    col_a,col_b = st.columns(2)
    with col_a:
        st.markdown("#### 월별 시즌 인덱스")
        st.caption("1.0 = 평균, 1.5 = 평균보다 50% 높은 시즌")
        si_df = season_idx.reset_index()
        si_df.columns = ['월','시즌인덱스']
        si_df['해석'] = si_df['시즌인덱스'].apply(lambda x: '성수기🔥' if x>=1.2 else ('비수기❄️' if x<=0.8 else '보통'))
        month_names = {1:'1월',2:'2월',3:'3월',4:'4월',5:'5월',6:'6월',
                       7:'7월',8:'8월',9:'9월',10:'10월',11:'11월',12:'12월'}
        si_df['월명'] = si_df['월'].map(month_names)
        fig2 = go.Figure(go.Bar(x=si_df['월명'], y=si_df['시즌인덱스'],
                                marker_color=['#E53935' if x>=1.2 else '#1E88E5' if x<=0.8 else '#90CAF9' for x in si_df['시즌인덱스']],
                                text=si_df['시즌인덱스'].apply(lambda x: f"{x:.2f}"), textposition='outside'))
        fig2.add_hline(y=1.0, line_dash='dash', line_color='gray')
        fig2.update_layout(height=300, plot_bgcolor='white', yaxis_title='인덱스')
        st.plotly_chart(fig2, use_container_width=True)
    with col_b:
        st.markdown("#### 보정 성장률 (MoM)")
        monthly_all['보정_MoM'] = monthly_all['보정매출_만'].pct_change()*100
        monthly_all['실제_MoM'] = monthly_all['매출_만'].pct_change()*100
        comp = monthly_all[['년월','실제_MoM','보정_MoM']].dropna()
        fig3 = go.Figure()
        fig3.add_bar(x=comp['년월'], y=comp['실제_MoM'].round(1), name='실제 MoM',
                     marker_color='#90CAF9', opacity=0.7)
        fig3.add_scatter(x=comp['년월'], y=comp['보정_MoM'].round(1), name='보정 MoM',
                         mode='lines+markers', line=dict(color='#1E88E5',width=2))
        fig3.update_layout(height=300, plot_bgcolor='white', yaxis_title='%',
                           legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45))
        st.plotly_chart(fig3, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 9 — 팝업 분석
# ════════════════════════════════════════════════════════════════════════════
with tab9:
    st.markdown("### 🏪 팝업(더현대) 분석")

    popup_df = df_all[df_all['채널']=='팝업(더현대)'].copy()
    online_df = df_all[df_all['채널']!='팝업(더현대)'].copy()

    if popup_df.empty:
        st.info("팝업 데이터가 없습니다.")
    else:
        popup_periods = sorted(popup_df['년월'].unique())
        st.markdown(f"**팝업 운영 기간:** {popup_periods[0]} ~ {popup_periods[-1]} ({len(popup_periods)}개월)")

        c1,c2,c3 = st.columns(3)
        kpi_card(c1, "팝업 총 매출", f"{popup_df['판매금액'].sum()/10000:,.0f}만", color="#8E24AA")
        kpi_card(c2, "팝업 총 수량", f"{int(popup_df['수량'].sum()):,}개", color="#8E24AA")
        kpi_card(c3, "팝업 SKU 수", f"{popup_df['스타일코드'].nunique()}개", color="#8E24AA")

        # 팝업 상품 순위
        popup_style = popup_df.groupby(['스타일코드','대표명','카테고리']).agg(
            팝업매출=('판매금액','sum'), 팝업수량=('수량','sum')).reset_index()
        popup_style = popup_style.sort_values('팝업매출', ascending=False).reset_index(drop=True)
        popup_style['팝업매출_만'] = (popup_style['팝업매출']/10000).round(1)

        # 온라인 동일 상품 비교
        online_style = online_df.groupby('스타일코드').agg(
            온라인매출=('판매금액','sum'), 온라인수량=('수량','sum')).reset_index()
        compare = popup_style.merge(online_style, on='스타일코드', how='left')
        compare['온라인매출_만'] = (compare['온라인매출'].fillna(0)/10000).round(1)
        compare['온라인비율'] = (compare['팝업매출']/(compare['팝업매출']+compare['온라인매출'].fillna(0))*100).round(1)

        col_a,col_b = st.columns([3,2])
        with col_a:
            st.markdown("#### 팝업 TOP10 상품 & 온라인 비교")
            top10 = compare.head(10)
            fig = go.Figure()
            fig.add_bar(name='팝업 매출', x=top10['대표명'].apply(lambda x: x[:15]+'...' if len(x)>15 else x),
                        y=top10['팝업매출_만'], marker_color='#8E24AA')
            fig.add_bar(name='온라인 매출', x=top10['대표명'].apply(lambda x: x[:15]+'...' if len(x)>15 else x),
                        y=top10['온라인매출_만'], marker_color='#CE93D8')
            fig.update_layout(barmode='group', title='팝업 vs 온라인 매출 비교 (TOP10)',
                              height=380, plot_bgcolor='white',
                              legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-30))
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            st.markdown("#### 팝업 전용 상품 (온라인 없음)")
            popup_only = compare[compare['온라인매출'].isna() | (compare['온라인매출']==0)]
            if not popup_only.empty:
                for _, r in popup_only.head(8).iterrows():
                    st.markdown(f'<div class="warn-card"><b>{r["대표명"]}</b><br>{r["카테고리"]} | 팝업 {r["팝업매출_만"]}만 | {int(r["팝업수량"])}개</div>', unsafe_allow_html=True)
            else:
                st.info("팝업 전용 상품 없음 (모든 상품이 온라인에도 있음)")

        # 팝업 기간 온라인 매출 변화 (헤일로 or 카니발리제이션)
        st.markdown("#### 팝업 기간 온라인 채널 영향")
        st.caption("팝업 운영 월의 온라인 매출이 전월 대비 증가 = 헤일로 효과 / 감소 = 카니발리제이션")

        popup_month_rev = online_df[online_df['년월'].isin(popup_periods)].groupby('년월')['판매금액'].sum()
        all_monthly_online = online_df.groupby('년월')['판매금액'].sum().sort_index()

        halo_fig = go.Figure()
        halo_fig.add_scatter(x=all_monthly_online.index, y=(all_monthly_online/10000).round(),
                             name='온라인 전체', mode='lines+markers', line=dict(color='#1E88E5',width=2))
        if not popup_month_rev.empty:
            halo_fig.add_scatter(x=popup_month_rev.index, y=(popup_month_rev/10000).round(),
                                 name='팝업 운영월', mode='markers',
                                 marker=dict(color='#8E24AA',size=12,symbol='star'))
        halo_fig.update_layout(title='온라인 매출 추이 + 팝업 운영월 표시',
                               height=300, plot_bgcolor='white',
                               legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45))
        st.plotly_chart(halo_fig, use_container_width=True)

        # 상세 테이블
        with st.expander("📋 팝업 전체 상품 목록"):
            disp = compare[['대표명','카테고리','팝업매출_만','팝업수량','온라인매출_만','온라인비율']].copy()
            disp['온라인비율'] = disp['온라인비율'].apply(lambda x: f"{x}%")
            disp.columns = ['상품명','카테고리','팝업매출(만)','팝업수량','온라인매출(만)','팝업비율(%)']
            st.dataframe(disp, use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 10 — 채널 카니발리제이션
# ════════════════════════════════════════════════════════════════════════════
with tab10:
    st.markdown("### 🔀 채널 카니발리제이션 분석")
    st.caption("무신사 할인(실현율↓) 기간에 자사몰 매출이 감소하는지 상관관계 분석")

    ch_monthly_wide = df_all.groupby(['년월','채널'])['판매금액'].sum().unstack(fill_value=0)
    ch_monthly_wide.columns.name = None
    ch_monthly_wide = ch_monthly_wide.reset_index()
    for ch in CHANNELS:
        if ch not in ch_monthly_wide.columns:
            ch_monthly_wide[ch] = 0
    ch_monthly_wide = ch_monthly_wide.sort_values('년월')

    # 무신사 실현율 월별
    musinsa_real = df_all[df_all['채널']=='무신사'].groupby('년월')['실현율'].mean().reset_index()
    musinsa_real.columns = ['년월','무신사_실현율']
    ch_monthly_wide = ch_monthly_wide.merge(musinsa_real, on='년월', how='left')

    col_a,col_b = st.columns(2)
    with col_a:
        # 무신사 실현율 vs 자사몰 매출
        merged = ch_monthly_wide.dropna(subset=['무신사_실현율'])
        if len(merged) > 3 and '자사몰' in merged.columns:
            corr = merged['무신사_실현율'].corr(merged['자사몰'])
            st.metric("무신사 실현율 ↔ 자사몰 매출 상관계수", f"{corr:.3f}",
                      help="양수=같이 움직임, 음수=반대로 움직임(카니발리제이션)")
            if corr < -0.3:
                st.markdown('<div class="danger-card">🔴 카니발리제이션 감지: 무신사 할인 시 자사몰 매출 감소 경향</div>', unsafe_allow_html=True)
            elif corr > 0.3:
                st.markdown('<div class="good-card">🟢 헤일로 효과: 무신사 할인 시 자사몰도 함께 상승</div>', unsafe_allow_html=True)
            else:
                st.info("📊 채널 간 독립적 움직임 — 카니발리제이션 없음")

            fig = go.Figure()
            fig.add_scatter(x=merged['년월'], y=merged['무신사_실현율'],
                            name='무신사 실현율(%)', yaxis='y2',
                            line=dict(color='#1E88E5',width=2,dash='dash'), mode='lines+markers')
            fig.add_bar(x=merged['년월'], y=(merged['자사몰']/10000).round(),
                        name='자사몰 매출(만)', marker_color='#43A047', opacity=0.7)
            fig.update_layout(title='무신사 실현율 vs 자사몰 매출',
                              yaxis=dict(title='자사몰 매출(만)'),
                              yaxis2=dict(title='실현율(%)',overlaying='y',side='right',range=[60,100]),
                              height=340, plot_bgcolor='white',
                              legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45))
            st.plotly_chart(fig, use_container_width=True)

    with col_b:
        # 채널 간 상관 히트맵
        ch_for_corr = [ch for ch in CHANNELS if ch in ch_monthly_wide.columns and ch_monthly_wide[ch].sum()>0]
        if len(ch_for_corr) >= 2:
            corr_matrix = ch_monthly_wide[ch_for_corr].corr().round(2)
            fig2 = px.imshow(corr_matrix, title='채널 간 매출 상관관계',
                             color_continuous_scale='RdBu_r', zmin=-1, zmax=1,
                             text_auto=True)
            fig2.update_layout(height=340)
            st.plotly_chart(fig2, use_container_width=True)
            st.caption("1.0 = 완전 동행 | -1.0 = 완전 반대 | 0 = 무관계")

    # 채널별 월별 비중 변화
    ch_pct = ch_monthly_wide.copy()
    total = ch_pct[CHANNELS].sum(axis=1)
    for ch in CHANNELS:
        if ch in ch_pct.columns:
            ch_pct[ch] = (ch_pct[ch]/total*100).round(1)
    ch_pct_melt = ch_pct[['년월']+[ch for ch in CHANNELS if ch in ch_pct.columns]].melt(id_vars='년월', var_name='채널', value_name='비중(%)')
    fig3 = px.area(ch_pct_melt, x='년월', y='비중(%)', color='채널',
                   title='채널 비중 변화 추이', color_discrete_map=CH_COLORS)
    fig3.update_layout(height=300, plot_bgcolor='white',
                       legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-45))
    st.plotly_chart(fig3, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 11 — 관리
# ════════════════════════════════════════════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════════
# TAB 11 — 마케팅
# ════════════════════════════════════════════════════════════════════════════
with tab11:
    st.markdown("### 📣 마케팅 효율 분석")

    mktg = load_marketing(str(MKTG_FILE))
    if mktg is None:
        st.warning(f"마케팅 파일을 찾을 수 없습니다: {MKTG_FILE}")
    else:
        meta_ms = mktg['meta_ms']
        meta_js = mktg['meta_js']
        prod_ms = mktg['prod_ms']
        prod_29 = mktg['prod_29']

        # ── 전체 KPI ─────────────────────────────────────────────────────────
        total_adspend = (meta_ms['지출 금액 (KRW)'].sum() + meta_js['지출 금액 (KRW)'].sum() +
                         prod_ms['집행 광고비'].sum() + prod_29['집행 광고비'].sum())
        total_conv    = (meta_ms['공유 항목의 구매 전환값'].sum() + meta_js['구매 전환값'].sum() +
                         prod_ms['매출'].sum() + prod_29['매출'].sum())
        total_roas    = total_conv / total_adspend if total_adspend > 0 else 0
        total_purchase = meta_ms['공유 항목이 포함된 구매'].sum() + meta_js['구매'].sum()

        c1,c2,c3,c4 = st.columns(4)
        kpi_card(c1, "총 광고비", f"{total_adspend/10000:,.0f}만원", color="#E53935")
        kpi_card(c2, "총 구매전환값", f"{total_conv/10000:,.0f}만원", color="#1E88E5")
        kpi_card(c3, "전체 ROAS", f"{total_roas:.1f}x", color="#43A047")
        kpi_card(c4, "총 구매 건수", f"{int(total_purchase):,}건", color="#FB8C00")

        st.markdown("---")
        mkt_sub = st.tabs(["📊 채널별 ROAS", "📈 일별 추이", "🎯 광고 형식별", "🛍️ 상품광고"])

        # ── 채널별 ROAS ───────────────────────────────────────────────────────
        with mkt_sub[0]:
            # 월별 집계
            ms_m = meta_ms.groupby('년월').agg(
                광고비=('지출 금액 (KRW)','sum'), 클릭=('링크 클릭','sum'),
                노출=('노출','sum'), 구매전환값=('공유 항목의 구매 전환값','sum'),
                구매=('공유 항목이 포함된 구매','sum')).reset_index()
            ms_m['ROAS'] = (ms_m['구매전환값']/ms_m['광고비']).replace([np.inf,-np.inf],0).round(2)
            ms_m['CTR'] = (ms_m['클릭']/ms_m['노출']*100).round(3)
            ms_m['CPC'] = (ms_m['광고비']/ms_m['클릭'].replace(0,np.nan)).round(0)
            ms_m['채널'] = '무신사 메타'

            js_m = meta_js.groupby('년월').agg(
                광고비=('지출 금액 (KRW)','sum'), 클릭=('링크 클릭','sum'),
                노출=('노출','sum'), 구매전환값=('구매 전환값','sum'),
                구매=('구매','sum')).reset_index()
            js_m['ROAS'] = (js_m['구매전환값']/js_m['광고비']).replace([np.inf,-np.inf],0).round(2)
            js_m['CTR'] = (js_m['클릭']/js_m['노출']*100).round(3)
            js_m['CPC'] = (js_m['광고비']/js_m['클릭'].replace(0,np.nan)).round(0)
            js_m['채널'] = '자사몰 메타'

            pm_m = prod_ms.groupby('년월').agg(
                광고비=('집행 광고비','sum'), 구매전환값=('매출','sum'),
                클릭=('클릭 수','sum'), 노출=('노출 수','sum')).reset_index()
            pm_m['ROAS'] = (pm_m['구매전환값']/pm_m['광고비']).replace([np.inf,-np.inf],0).round(2)
            pm_m['CTR'] = (pm_m['클릭']/pm_m['노출']*100).round(3)
            pm_m['CPC'] = (pm_m['광고비']/pm_m['클릭'].replace(0,np.nan)).round(0)
            pm_m['채널'] = '무신사 상품광고'; pm_m['구매'] = 0

            p29_m = prod_29.groupby('년월').agg(
                광고비=('집행 광고비','sum'), 구매전환값=('매출','sum'),
                클릭=('클릭 수','sum'), 노출=('노출 수','sum')).reset_index()
            p29_m['ROAS'] = (p29_m['구매전환값']/p29_m['광고비']).replace([np.inf,-np.inf],0).round(2)
            p29_m['CTR'] = (p29_m['클릭']/p29_m['노출']*100).round(3)
            p29_m['CPC'] = (p29_m['광고비']/p29_m['클릭'].replace(0,np.nan)).round(0)
            p29_m['채널'] = '29CM 상품광고'; p29_m['구매'] = 0

            all_ch = pd.concat([ms_m, js_m, pm_m, p29_m], ignore_index=True)

            MKTG_COLORS = {'무신사 메타':'#1E88E5','자사몰 메타':'#43A047',
                           '무신사 상품광고':'#FB8C00','29CM 상품광고':'#E53935'}

            col_a, col_b = st.columns(2)
            with col_a:
                fig = px.line(all_ch, x='년월', y='ROAS', color='채널', markers=True,
                              title='채널별 월별 ROAS', color_discrete_map=MKTG_COLORS)
                fig.add_hline(y=3, line_dash='dash', line_color='gray', annotation_text='ROAS 3x 기준선')
                fig.update_layout(height=320, plot_bgcolor='white',
                                  legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-30))
                st.plotly_chart(fig, use_container_width=True)
            with col_b:
                fig2 = px.bar(all_ch, x='년월', y='광고비', color='채널', barmode='stack',
                              title='채널별 월별 광고비', color_discrete_map=MKTG_COLORS)
                fig2.update_layout(height=320, plot_bgcolor='white',
                                   legend=dict(orientation='h',y=1.1), xaxis=dict(tickangle=-30))
                st.plotly_chart(fig2, use_container_width=True)

            # 채널 종합 비교표
            summary_rows = []
            for d, ch in [(ms_m,'무신사 메타'),(js_m,'자사몰 메타'),(pm_m,'무신사 상품광고'),(p29_m,'29CM 상품광고')]:
                summary_rows.append({
                    '채널': ch,
                    '총 광고비': f"{d['광고비'].sum()/10000:,.0f}만",
                    '총 구매전환값': f"{d['구매전환값'].sum()/10000:,.0f}만",
                    '평균 ROAS': f"{(d['구매전환값'].sum()/d['광고비'].sum()):.1f}x" if d['광고비'].sum()>0 else "-",
                    '평균 CTR': f"{d['CTR'].mean():.3f}%",
                    '평균 CPC': f"{d['CPC'].mean():,.0f}원"
                })
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        # ── 일별 추이 ─────────────────────────────────────────────────────────
        with mkt_sub[1]:
            st.markdown("#### 무신사 메타 광고 일별 추이")
            ms_daily = meta_ms.groupby('날짜').agg(
                광고비=('지출 금액 (KRW)','sum'),
                구매전환값=('공유 항목의 구매 전환값','sum'),
                클릭=('링크 클릭','sum'),
                노출=('노출','sum'),
                구매=('공유 항목이 포함된 구매','sum')).reset_index()
            ms_daily['ROAS'] = (ms_daily['구매전환값']/ms_daily['광고비'].replace(0,np.nan)).fillna(0).round(2)
            ms_daily['CTR'] = (ms_daily['클릭']/ms_daily['노출'].replace(0,np.nan)*100).fillna(0).round(3)

            fig = go.Figure()
            fig.add_bar(x=ms_daily['날짜'], y=ms_daily['광고비']/10000, name='광고비(만)', marker_color='#BBDEFB', opacity=0.8)
            fig.add_scatter(x=ms_daily['날짜'], y=ms_daily['구매전환값']/10000, name='전환값(만)',
                            yaxis='y', line=dict(color='#1E88E5',width=2), mode='lines')
            fig.add_scatter(x=ms_daily['날짜'], y=ms_daily['ROAS'], name='ROAS',
                            yaxis='y2', line=dict(color='#E53935',width=2,dash='dot'), mode='lines+markers')
            fig.update_layout(title='광고비 vs 전환값 vs ROAS', height=340,
                              yaxis=dict(title='금액(만원)'),
                              yaxis2=dict(title='ROAS',overlaying='y',side='right'),
                              plot_bgcolor='white', legend=dict(orientation='h',y=1.1))
            st.plotly_chart(fig, use_container_width=True)

            col_a, col_b = st.columns(2)
            with col_a:
                fig2 = go.Figure()
                fig2.add_scatter(x=ms_daily['날짜'], y=ms_daily['CTR'], name='CTR(%)',
                                 line=dict(color='#43A047',width=2), mode='lines')
                fig2.update_layout(title='CTR 일별 추이', height=240, plot_bgcolor='white',
                                   yaxis_title='CTR(%)')
                st.plotly_chart(fig2, use_container_width=True)
            with col_b:
                fig3 = go.Figure()
                fig3.add_bar(x=ms_daily['날짜'], y=ms_daily['구매'], name='구매건수',
                             marker_color='#FB8C00')
                fig3.update_layout(title='일별 구매 건수', height=240, plot_bgcolor='white',
                                   yaxis_title='건수')
                st.plotly_chart(fig3, use_container_width=True)

        # ── 광고 형식별 ───────────────────────────────────────────────────────
        with mkt_sub[2]:
            st.markdown("#### 광고 형식 / 소재별 효율")
            form_df = meta_ms.groupby('광고 형식').agg(
                광고비=('지출 금액 (KRW)','sum'),
                클릭=('링크 클릭','sum'),
                노출=('노출','sum'),
                구매전환값=('공유 항목의 구매 전환값','sum'),
                구매=('공유 항목이 포함된 구매','sum')).reset_index()
            form_df['ROAS'] = (form_df['구매전환값']/form_df['광고비'].replace(0,np.nan)).fillna(0).round(2)
            form_df['CTR'] = (form_df['클릭']/form_df['노출']*100).round(3)
            form_df['CPC'] = (form_df['광고비']/form_df['클릭'].replace(0,np.nan)).round(0)
            form_df['광고비_만'] = (form_df['광고비']/10000).round(1)
            form_df['전환값_만'] = (form_df['구매전환값']/10000).round(1)

            col_a, col_b, col_c = st.columns(3)
            with col_a:
                fig = px.bar(form_df, x='광고 형식', y='ROAS', color='광고 형식',
                             title='형식별 ROAS', text='ROAS',
                             color_discrete_sequence=['#1E88E5','#43A047','#FB8C00'])
                fig.update_traces(textposition='outside')
                fig.update_layout(height=300, plot_bgcolor='white', showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            with col_b:
                fig2 = px.bar(form_df, x='광고 형식', y='CTR', color='광고 형식',
                              title='형식별 CTR(%)', text=form_df['CTR'].apply(lambda x: f"{x:.3f}%"),
                              color_discrete_sequence=['#1E88E5','#43A047','#FB8C00'])
                fig2.update_traces(textposition='outside')
                fig2.update_layout(height=300, plot_bgcolor='white', showlegend=False)
                st.plotly_chart(fig2, use_container_width=True)
            with col_c:
                fig3 = px.bar(form_df, x='광고 형식', y='CPC', color='광고 형식',
                              title='형식별 CPC(원)', text=form_df['CPC'].apply(lambda x: f"{x:,.0f}원"),
                              color_discrete_sequence=['#1E88E5','#43A047','#FB8C00'])
                fig3.update_traces(textposition='outside')
                fig3.update_layout(height=300, plot_bgcolor='white', showlegend=False)
                st.plotly_chart(fig3, use_container_width=True)

            # 월별 형식별 ROAS 추이
            form_monthly = meta_ms.groupby(['년월','광고 형식']).agg(
                광고비=('지출 금액 (KRW)','sum'),
                구매전환값=('공유 항목의 구매 전환값','sum')).reset_index()
            form_monthly['ROAS'] = (form_monthly['구매전환값']/form_monthly['광고비'].replace(0,np.nan)).fillna(0).round(2)
            fig4 = px.line(form_monthly, x='년월', y='ROAS', color='광고 형식', markers=True,
                           title='월별 광고 형식별 ROAS 추이')
            fig4.update_layout(height=280, plot_bgcolor='white', legend=dict(orientation='h',y=1.1))
            st.plotly_chart(fig4, use_container_width=True)

            # 상세 테이블
            disp_form = form_df[['광고 형식','광고비_만','전환값_만','ROAS','CTR','CPC','구매']].copy()
            disp_form.columns = ['광고 형식','광고비(만)','전환값(만)','ROAS','CTR(%)','CPC(원)','구매건수']
            st.dataframe(disp_form, use_container_width=True, hide_index=True)

        # ── 상품광고 ─────────────────────────────────────────────────────────
        with mkt_sub[3]:
            st.markdown("#### 상품광고 — 무신사 vs 29CM")

            ms_m2 = prod_ms.groupby('년월').agg(
                광고비=('집행 광고비','sum'), 매출=('매출','sum'),
                클릭=('클릭 수','sum'), 노출=('노출 수','sum')).reset_index()
            ms_m2['ROAS'] = (ms_m2['매출']/ms_m2['광고비'].replace(0,np.nan)).fillna(0).round(2)
            ms_m2['CTR'] = (ms_m2['클릭']/ms_m2['노출'].replace(0,np.nan)*100).fillna(0).round(3)
            ms_m2['채널'] = '무신사'

            p29_m2 = prod_29.groupby('년월').agg(
                광고비=('집행 광고비','sum'), 매출=('매출','sum'),
                클릭=('클릭 수','sum'), 노출=('노출 수','sum')).reset_index()
            p29_m2['ROAS'] = (p29_m2['매출']/p29_m2['광고비'].replace(0,np.nan)).fillna(0).round(2)
            p29_m2['CTR'] = (p29_m2['클릭']/p29_m2['노출'].replace(0,np.nan)*100).fillna(0).round(3)
            p29_m2['채널'] = '29CM'

            prod_all = pd.concat([ms_m2, p29_m2], ignore_index=True)

            col_a, col_b = st.columns(2)
            with col_a:
                fig = px.bar(prod_all, x='년월', y='ROAS', color='채널', barmode='group',
                             title='상품광고 ROAS 비교',
                             color_discrete_map={'무신사':'#1E88E5','29CM':'#E53935'})
                fig.add_hline(y=3, line_dash='dash', line_color='gray')
                fig.update_layout(height=300, plot_bgcolor='white',
                                  legend=dict(orientation='h',y=1.1))
                st.plotly_chart(fig, use_container_width=True)
            with col_b:
                fig2 = go.Figure()
                fig2.add_scatter(x=ms_m2['년월'], y=ms_m2['CTR'], name='무신사 CTR(%)',
                                 mode='lines+markers', line=dict(color='#1E88E5',width=2))
                fig2.add_scatter(x=p29_m2['년월'], y=p29_m2['CTR'], name='29CM CTR(%)',
                                 mode='lines+markers', line=dict(color='#E53935',width=2))
                fig2.update_layout(title='상품광고 CTR 비교', height=300, plot_bgcolor='white',
                                   legend=dict(orientation='h',y=1.1))
                st.plotly_chart(fig2, use_container_width=True)

            # 광고비 vs 매출 효율
            fig3 = go.Figure()
            fig3.add_bar(x=ms_m2['년월'], y=ms_m2['광고비']/10000, name='무신사 광고비(만)', marker_color='#BBDEFB')
            fig3.add_bar(x=p29_m2['년월'], y=p29_m2['광고비']/10000, name='29CM 광고비(만)', marker_color='#FFCDD2')
            fig3.add_scatter(x=ms_m2['년월'], y=ms_m2['매출']/10000, name='무신사 매출(만)',
                             mode='lines+markers', line=dict(color='#1E88E5',width=2))
            fig3.add_scatter(x=p29_m2['년월'], y=p29_m2['매출']/10000, name='29CM 매출(만)',
                             mode='lines+markers', line=dict(color='#E53935',width=2))
            fig3.update_layout(title='광고비 vs 광고 기여 매출', barmode='group',
                               height=300, plot_bgcolor='white',
                               legend=dict(orientation='h',y=1.1), yaxis_title='만원')
            st.plotly_chart(fig3, use_container_width=True)

            # 요약 테이블
            for d, nm in [(ms_m2,'무신사'),(p29_m2,'29CM')]:
                st.markdown(f"**{nm} 상품광고 월별**")
                disp = d[['년월','광고비','매출','ROAS','CTR']].copy()
                disp['광고비'] = disp['광고비'].apply(lambda x: f"{x/10000:,.1f}만")
                disp['매출'] = disp['매출'].apply(lambda x: f"{x/10000:,.1f}만")
                disp['ROAS'] = disp['ROAS'].apply(lambda x: f"{x:.1f}x")
                disp['CTR'] = disp['CTR'].apply(lambda x: f"{x:.3f}%")
                st.dataframe(disp, use_container_width=True, hide_index=True)

with tab12:
    st.markdown("### ⚙️ 관리")
    mgmt = load_mgmt()
    sub1,sub2,sub3 = st.tabs(["🎯 월별 목표","📦 상품 메모","✅ TODO"])

    with sub1:
        st.markdown("**월별 매출 목표 설정 (만원)**")
        targets = mgmt.get("targets",{})
        cols = st.columns(3)
        new_targets = {}
        for i,ym in enumerate(months_all):
            with cols[i%3]:
                actual = df_all[df_all['년월']==ym]['판매금액'].sum()/10000
                val = st.number_input(f"{ym}", value=int(targets.get(ym, round(actual*1.1))), step=100, key=f"t_{ym}")
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
