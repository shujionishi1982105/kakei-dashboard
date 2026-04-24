import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import glob
import re
import os
import base64

# ==========================================
# 1. ログイン認証ロジック
# ==========================================
USER_ID = "admin"
USER_PASS = "kakeiseikeigeka"

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown("### 🔐 かけい整形外科 経営分析ログイン")
    user_input = st.text_input("ユーザーID", key="user_id")
    pw_input = st.text_input("パスワード", type="password", key="user_pw")
    
    if st.button("ログイン"):
        if user_input.strip() == USER_ID and pw_input.strip() == USER_PASS:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("IDまたはパスワードが違います")
    return False

if not check_password():
    st.stop()

# ==========================================
# 2. 基本設定とCSS（スマホ対応）
# ==========================================
st.cache_data.clear()
st.set_page_config(page_title="かけい整形外科 経営分析", layout="wide")

def get_base64_of_bin_file(bin_file):
    if os.path.exists(bin_file):
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    return None

logo_path = "スクリーンショット 2026-03-27 163556.png"
logo_base64 = get_base64_of_bin_file(logo_path)

st.markdown("""
    <style>
    div.stButton > button {
        height: 65px !important; border-radius: 10px !important;
        font-weight: bold !important; font-size: 15px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.2s ease-in-out !important; white-space: normal !important;
    }
    div.stButton > button[kind="secondary"] { background-color: #EBF5FB !important; border: 1px solid #AED6F1 !important; color: #154360 !important; }
    div.stButton > button[kind="primary"] { background-color: #E74C3C !important; border: 1px solid #E74C3C !important; color: #FFFFFF !important; }
    
    .header-container { display: flex; align-items: center; gap: 15px; margin-bottom: 25px; }
    .header-logo { height: 55px; width: auto; object-fit: contain; }
    .header-title { margin: 0 !important; color: #2C3E50; font-size: 32px; font-weight: bold; }

    @media screen and (max-width: 768px) {
        .header-container { flex-direction: column !important; }
        .header-title { font-size: 20px !important; text-align: center !important; }
        div.stButton > button { height: 45px !important; font-size: 13px !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 3. データ読み込み関数群
# ==========================================

# (1) 診療行為別CSV（品目データ）
@st.cache_data
def load_kakei_data():
    search_paths = [".", r"G:\マイドライブ\かけい整形外科ダッシュボード", r"G:\My Drive\かけい整形外科ダッシュボード"]
    all_files = []
    for d_path in search_paths:
        if os.path.exists(d_path): all_files.extend(glob.glob(os.path.join(d_path, "*.csv")))
    
    unique_files = {os.path.basename(f): f for f in all_files}
    target_files = [f for f in unique_files.values() if "品目" in f]
    if not target_files: return pd.DataFrame()

    CATEGORY_MAP = {1: '薬剤（院内）・院外処方', 3: '基本診療料・医学管理料等', 4: '調剤・処方', 5: '注射', 6: '処置', 7: '手術', 8: '検査', 9: '画像診断', 10: 'その他', 11: '自費', 14: '薬剤（院内）・院外処方'}
    combined_list = []
    for file in target_files:
        df = None
        for enc in ['cp932', 'utf-8-sig', 'utf-8']:
            try:
                df = pd.read_csv(file, encoding=enc)
                break
            except: continue
        if df is None: continue
        df.columns = df.columns.str.strip().str.replace('"', '').str.replace("'", "")
        df = df.rename(columns={'品名': '診療行為', '価格１': '単価', '合計　使用量': '回数', '合計 使用量': '回数'})
        if '診療行為' not in df.columns or '品目分類' not in df.columns: continue
        df = df[~df['品目分類'].isin([2, 12, 13])]
        df['カテゴリ名'] = df['品目分類'].map(CATEGORY_MAP)
        df['単価'] = pd.to_numeric(df['単価'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        df['回数'] = pd.to_numeric(df['回数'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        df['総値'] = df['単価'] * df['回数']
        df['金額_円'] = df['総値'] * 10
        df.loc[df['品目分類'] == 11, '金額_円'] = df.loc[df['品目分類'] == 11, '総値']
        df.loc[df['品目分類'] == 14, '金額_円'] = 0  
        match = re.search(r'\d{6}', file)
        month_raw = match.group() if match else "000000"
        year, month = int(month_raw[:4]), int(month_raw[4:])
        df['月単体'] = f"{month}月"
        df['年'] = f"R{year-2018}年"
        combined_list.append(df)
    return pd.concat(combined_list, ignore_index=True) if combined_list else pd.DataFrame()

# (2) スプレッドシート（来院統計）
@st.cache_data
def load_spreadsheet_visit_data():
    sheet_id = "1dxSQfw9S5H_d9xNJ_VxduO4R1XRbscN1cpobCvmlhIw"
    gid = "1419942397"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    try:
        df_raw = pd.read_csv(url, header=None)
        records, current_months = [], []
        for idx, row in df_raw.iterrows():
            metric = str(row[0]).strip()
            if metric == '来院種別':
                current_months = [str(x).strip() for x in row[1:] if pd.notna(x) and str(x).strip() not in ['None', 'nan', '']]
            elif metric not in ['None', 'nan', '', 'NaN'] and current_months:
                for i, val in enumerate(row[1:]):
                    if i < len(current_months):
                        records.append({'年月': current_months[i], '指標': metric, '値': val})
        df_parsed = pd.DataFrame(records)
        df_parsed['値'] = df_parsed['値'].astype(str).str.replace(',', '', regex=False).str.replace('%', '', regex=False)
        df_parsed['値'] = pd.to_numeric(df_parsed['値'], errors='coerce').fillna(0)
        df_v = df_parsed.pivot_table(index='年月', columns='指標', values='値', aggfunc='sum').reset_index()
        df_v['年月_dt'] = pd.to_datetime(df_v['年月'].str.extract(r'(\d{4}年\d{1,2}月)')[0].str.replace('年', '/').str.replace('月', ''), format='%Y/%m', errors='coerce')
        df_v = df_v.dropna(subset=['年月_dt']).sort_values('年月_dt')
        df_v['年'] = df_v['年月_dt'].dt.year.astype(str) + "年"
        df_v['月単体'] = df_v['年月_dt'].dt.month.astype(str) + "月"
        return df_v
    except: return pd.DataFrame()

# (3) 患者属性・エリア別CSV（名簿データ）
@st.cache_data
def load_patient_data():
    search_paths = [".", r"G:\マイドライブ\かけい整形外科ダッシュボード", r"G:\My Drive\かけい整形外科ダッシュボード"]
    all_files = []
    for d_path in search_paths:
        if os.path.exists(d_path): all_files.extend(glob.glob(os.path.join(d_path, "*.csv")))
    unique_files = {os.path.basename(f): f for f in all_files}
    patient_files = [f for f in unique_files.values() if "品目" not in f and re.search(r'\d{4}', f)]
    if not patient_files: return pd.DataFrame()
    combined_list = []
    for file in patient_files:
        df = None
        for enc in ['cp932', 'utf-8-sig', 'utf-8']:
            try:
                df_t = pd.read_csv(file, encoding=enc, nrows=5)
                if '現住所 住所' in df_t.columns and '年齢' in df_t.columns:
                    df = pd.read_csv(file, encoding=enc); break
            except: continue
        if df is None: continue
        match = re.search(r'(\d{4})', file)
        year = match.group(1) if match else "0000"
        df['年'] = f"{year}年"
        combined_list.append(df)
    if not combined_list: return pd.DataFrame()
    df_p = pd.concat(combined_list, ignore_index=True)
    df_p['年齢_数値'] = df_p['年齢'].astype(str).str.extract(r'(\d+)').astype(float)
    bins = [0, 9, 19, 29, 39, 49, 59, 69, 79, 89, 150]
    labels = ['0-9歳', '10-19歳', '20-29歳', '30-39歳', '40-49歳', '50-59歳', '60-69歳', '70-79歳', '80-89歳', '90歳以上']
    df_p['年齢層'] = pd.cut(df_p['年齢_数値'], bins=bins, labels=labels, right=True, include_lowest=True)
    addr = df_p['現住所 住所'].astype(str).str.replace(r'^.{2,3}[都道府県]', '', regex=True)
    area_extracted = addr.str.extract(r'([^市区町村]+[市区町村][^町村]+[町村])')[0].fillna(addr.str.extract(r'^([^\d０-９]+)')[0])
    df_p['エリア'] = area_extracted.str.replace(r'字.*$', '', regex=True).str.strip()
    if '患者番号' in df_p.columns: df_p = df_p.drop_duplicates(subset=['年', '患者番号'])
    return df_p

def calc_diff_ratio(curr_val, prev_val):
    if prev_val == 0: return "前年データなし"
    return f"前年比 {(curr_val / prev_val) * 100:.1f}%"

month_order = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]

# ==========================================
# 4. メイン画面・ナビゲーション
# ==========================================
if logo_base64:
    st.markdown(f'<div class="header-container"><img src="data:image/png;base64,{logo_base64}" class="header-logo"><h1 class="header-title">かけい整形外科 経営分析ダッシュボード</h1></div>', unsafe_allow_html=True)
else:
    st.markdown('<h1 class="header-title">かけい整形外科 経営分析ダッシュボード</h1>', unsafe_allow_html=True)

df_all = load_kakei_data()
if df_all.empty: st.error("CSVデータが見つかりません。"); st.stop()

kpi_pages = ["経営全体・主要KPI", "薬剤（院内）・院外処方", "基本診療料・医学管理料等", "調剤・処方", "注射", "処置", "手術", "検査", "画像診断", "その他", "自費", "AI総合アドバイス"]
patient_pages = ["来院分析", "患者属性推移", "エリア別推移"]

if 'current_page' not in st.session_state: st.session_state.current_page = kpi_pages[0]

st.write("### 🔍 KPI分析メニュー")
for i in range(0, len(kpi_pages), 4):
    cols = st.columns(4)
    for j in range(4):
        if i + j < len(kpi_pages):
            p_name = kpi_pages[i + j]
            with cols[j]:
                if st.button(p_name, use_container_width=True, key=f"kpi_{i+j}", type="primary" if st.session_state.current_page == p_name else "secondary"):
                    st.session_state.current_page = p_name; st.rerun()

st.write("### 👥 患者分析メニュー")
cols_p = st.columns(4)
for i, p_name in enumerate(patient_pages):
    with cols_p[i]:
        if st.button(p_name, use_container_width=True, key=f"pat_{i}", type="primary" if st.session_state.current_page == p_name else "secondary"):
            st.session_state.current_page = p_name; st.rerun()
st.write("---")

current_page = st.session_state.current_page

# ==========================================
# 5. 各ページ表示ロジック
# ==========================================

# --- 来院分析 ---
if current_page == "来院分析":
    st.header("📈 患者来院動向分析")
    df_v = load_spreadsheet_visit_data()
    if df_v.empty: st.warning("データが読み込めませんでした。")
    else:
        def get_col(keywords, default=None):
            for c in df_v.columns:
                if all(k in c for k in keywords): return c
            return default
        col_total = get_col(['合計', '総患者'], '合計（1日の総患者数）')
        col_new_rate = get_col(['新規', '率'], '新規患者率')
        col_rehab = get_col(['リハビリ合計'], 'リハビリ合計')
        col_avg_visit = get_col(['日平均来院数'], '１日平均来院数（人）')
        col_avg_rehab = get_col(['日平均リハビリ'], '1日平均リハビリ人数')
        
        df_long = df_v[df_v['年月_dt'] >= '2022-03-01'].copy()
        if col_total in df_long.columns: df_long = df_long[df_long[col_total] > 0]
        
        st.write("#### ① 延べ来院数・新規患者率の推移")
        fig1 = make_subplots(specs=[[{"secondary_y": True}]])
        fig1.add_trace(go.Bar(x=df_long['年月'], y=df_long[col_total], name='延べ来院数', marker_color='#3498DB'), secondary_y=False)
        fig1.add_trace(go.Scatter(x=df_long['年月'], y=df_long[col_new_rate], name='新規患者率(%)', mode='lines+markers', line=dict(color='#E74C3C', width=3)), secondary_y=True)
        fig1.update_layout(hovermode="x unified", margin=dict(t=30)); st.plotly_chart(fig1, use_container_width=True)

        visit_years = sorted(df_v['年'].unique())
        selected_v_year = st.selectbox("📅 年度選択", visit_years, index=len(visit_years)-1)
        df_curr_v = df_v[df_v['年'] == selected_v_year].copy()
        df_curr_v['月単体'] = pd.Categorical(df_curr_v['月単体'], categories=month_order, ordered=True)
        df_curr_v = df_curr_v.sort_values('月単体')

        st.write(f"#### ② {selected_v_year} 1日平均人数の推移")
        fig4 = go.Figure()
        for col, name, color in [(col_avg_visit, '1日平均来院数', '#2980B9'), (col_avg_rehab, '1日平均リハビリ', '#16A085')]:
            if col in df_curr_v.columns:
                fig4.add_trace(go.Scatter(x=df_curr_v['月単体'], y=df_curr_v[col], name=name, mode='lines+markers+text', text=df_curr_v[col].apply(lambda x: f"{x:.1f}" if x > 0 else ""), textposition='top center', line=dict(color=color, width=3)))
        fig4.update_layout(margin=dict(t=40)); st.plotly_chart(fig4, use_container_width=True)

        st.write("#### 📋 詳細データ一覧")
        target_order = ['診療のみ', '診療＆リハビリ', '診療のみ（新患）', '診療＆リハビリ（新患）', '小計', 'リハビリのみ', '合計（1日の総患者数）', '実稼働（1日:1 半日:0.5）', '１日平均来院数（人）', 'リハビリ合計', '1日平均リハビリ人数', '新規患者率']
        disp_df = df_curr_v.set_index('年月')[[c for c in target_order if c in df_curr_v.columns]].T
        st.dataframe(disp_df, use_container_width=True)

# --- 患者属性推移 ---
elif current_page == "患者属性推移":
    df_p = load_patient_data()
    if df_p.empty: st.warning("患者データが見つかりません。")
    else:
        st.header("👥 患者属性推移（年齢層別）")
        age_labels = ['0-9歳', '10-19歳', '20-29歳', '30-39歳', '40-49歳', '50-59歳', '60-69歳', '70-79歳', '80-89歳', '90歳以上']
        trend_df = df_p.groupby(['年', '年齢層']).size().unstack(fill_value=0)[age_labels]
        
        st.write("#### 📈 年毎の年齢層別 患者数推移")
        fig_t = go.Figure()
        colors = ['#F9EBEA', '#F2D7D5', '#D98880', '#C0392B', '#922B21', '#E8DAEF', '#C39BD3', '#8E44AD', '#5B2C6F', '#212F3D']
        for idx, age in enumerate(age_labels):
            fig_t.add_trace(go.Bar(x=trend_df.index, y=trend_df[age], name=age, marker_color=colors[idx]))
        fig_t.update_layout(barmode='stack', hovermode="x unified"); st.plotly_chart(fig_t, use_container_width=True)

        sel_y = st.selectbox("📅 年度選択", trend_df.index.tolist(), index=len(trend_df)-1)
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"#### 🥧 {sel_y} 年齢層の割合")
            pie_v = df_p[df_p['年']==sel_y]['年齢層'].value_counts().reindex(age_labels).fillna(0)
            fig_p = go.Figure(data=[go.Pie(labels=pie_v.index, values=pie_v.values, sort=False, hole=0.4, marker=dict(colors=colors))])
            st.plotly_chart(fig_p, use_container_width=True)
        with c2:
            st.write("#### 📋 データ詳細一覧")
            st.dataframe(trend_df.T.style.format("{:,.0f} 人"), use_container_width=True)

# --- エリア別推移 ---
elif current_page == "エリア別推移":
    df_p = load_patient_data()
    if df_p.empty: st.warning("患者データが見つかりません。")
    else:
        st.header("🗺️ エリア別（町名単位）推移")
        sel_y = st.selectbox("📅 年度選択", sorted(df_p['年'].unique()), index=len(df_p['年'].unique())-1)
        df_c = df_p[df_p['年'] == sel_y]
        df_prev = df_p[df_p['年'] == f"{int(sel_y[:4])-1}年"]
        
        area_c = df_c['エリア'].value_counts().reset_index(); area_c.columns = ['エリア', '当年']
        total = area_c['当年'].sum()
        area_c['割合(%)'] = area_c['當年'] / total * 100 if total > 0 else 0
        area_p = df_prev['エリア'].value_counts().reset_index(); area_p.columns = ['エリア', '前年']
        
        res = pd.merge(area_c, area_p, on='エリア', how='left').fillna(0)
        res['前年比'] = res.apply(lambda x: (x['当年']/x['前年']*100) if x['前年']>0 else 0, axis=1)
        top50 = res.sort_values('当年', ascending=False).head(50).reset_index(drop=True)
        top50.index += 1
        st.dataframe(top50.style.format({'当年':"{:,.0f}人",'割合(%)':"{:.1f}%",'前年':"{:,.0f}人",'前年比':"{:.1f}%"}), use_container_width=True)

# --- KPI分析ページ群 ---
elif current_page == "経営全体・主要KPI" or current_page == "AI総合アドバイス":
    available_years = sorted(df_all['年'].dropna().unique())
    selected_year = st.selectbox("📅 年度選択", available_years, index=len(available_years)-1)
    df_curr = df_all[df_all['年'] == selected_year]
    df_prev = df_all[df_all['年'] == f"R{int(selected_year[1:-1])-1}年"]
    
    st.subheader(f"📊 {selected_year} 全体業績サマリー")
    c1, c2, c3 = st.columns(3)
    c1.metric("当年度 総収益", f"¥{int(df_curr['金額_円'].sum()):,}", calc_diff_ratio(df_curr['金額_円'].sum(), df_prev[df_prev['月単体'].isin(df_curr['月単体'].unique())]['金額_円'].sum()))
    
    rehab_c = df_curr[df_curr['診療行為'].str.contains('リハビリ|運動器', na=False)]['金額_円'].sum()
    rehab_p = df_prev[df_prev['月単体'].isin(df_curr['月単体'].unique()) & df_prev['診療行為'].str.contains('リハビリ|運動器', na=False)]['金額_円'].sum()
    c2.metric("リハビリ料合計", f"¥{int(rehab_c):,}", calc_diff_ratio(rehab_c, rehab_p))
    
    st.write("#### 📋 詳細数値データ（年間一覧）")
    pivot = df_curr.pivot_table(index='月単体', columns='カテゴリ名', values='金額_円', aggfunc='sum').reindex(month_order).fillna(0)
    pivot.insert(0, '総診療報酬', pivot.sum(axis=1))
    sum_row = pivot.sum().to_frame().T; sum_row.index = ['年間合計']
    final = pd.concat([pivot, sum_row]).rename(columns={'基本診療料・医学管理料等':'基本・管理','薬剤（院内）・院外処方':'薬剤','調剤・処方':'調剤','画像診断':'画像'})
    st.dataframe(final.style.format("{:,.0f}"), use_container_width=True)

else:
    # 各カテゴリ詳細ページ
    target_cat = current_page
    available_years = sorted(df_all['年'].dropna().unique())
    selected_year = st.selectbox("📅 年度選択", available_years, index=len(available_years)-1)
    df_c = df_all[(df_all['年'] == selected_year) & (df_all['カテゴリ名'] == target_cat)]
    st.header(f"📊 {selected_year} 【{target_cat}】 詳細")
    if not df_c.empty:
        rank = df_c.groupby('診療行為').agg({'回数':'sum','総値':'sum'}).sort_values('総値', ascending=False).head(20)
        st.write("#### 🏆 TOP20 ランキング")
        st.dataframe(rank.style.format("{:,.0f}"), use_container_width=True)
