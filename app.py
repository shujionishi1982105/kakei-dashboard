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
    [data-testid="stTable"] th:first-child { min-width: 100px !important; }
    div[data-baseweb="select"], div[data-baseweb="select"] * { cursor: pointer !important; }

    div.stButton > button {
        height: 65px !important; border-radius: 10px !important;
        font-weight: bold !important; font-size: 15px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.2s ease-in-out !important; white-space: normal !important;
    }
    div.stButton > button[kind="secondary"] { background-color: #EBF5FB !important; border: 1px solid #AED6F1 !important; color: #154360 !important; }
    div.stButton > button[kind="primary"] { background-color: #E74C3C !important; border: 1px solid #E74C3C !important; color: #FFFFFF !important; }
    div.stButton > button:hover { transform: translateY(-2px) !important; filter: brightness(0.95) !important; }

    .header-container { display: flex; align-items: center; gap: 15px; margin-bottom: 25px; margin-top: 10px; }
    .header-logo { height: 55px; width: auto; object-fit: contain; }
    .header-title { margin: 0 !important; color: #2C3E50; font-size: 32px; font-weight: bold; }

    div[data-testid="stMetricValue"] { font-size: 2rem !important; white-space: nowrap !important; overflow: visible !important; }
    div[data-testid="stMetricValue"] > div { overflow: visible !important; text-overflow: clip !important; }

    @media screen and (max-width: 768px) {
        .header-container { flex-direction: column !important; gap: 5px !important; margin-bottom: 10px !important; }
        .header-logo { height: 40px !important; }
        .header-title { font-size: 20px !important; text-align: center !important; border-bottom: 2px solid #E74C3C !important; padding-bottom: 5px !important; }
        div.stButton > button { height: 45px !important; font-size: 13px !important; padding: 2px !important; }
        div[data-testid="stMetricValue"] { font-size: 1.5rem !important; }
        [data-testid="stTable"] th:first-child { min-width: 70px !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 3. データ読み込み関数
# ==========================================

# ① KPI用（品目データ）
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
        
        mask = df['診療行為'].str.contains('外来・在宅ベースアップ評価料', na=False)
        df.loc[mask, '診療行為'] = df.loc[mask, '診療行為'].astype(str) + '（' + df.loc[mask, '単価'].astype(int).astype(str) + '点）'
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

# ② 来院分析用（スプレッドシート）
@st.cache_data
def load_spreadsheet_visit_data():
    sheet_id = "1dxSQfw9S5H_d9xNJ_VxduO4R1XRbscN1cpobCvmlhIw"
    gid = "1419942397"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    try:
        df_raw = pd.read_csv(url, header=None)
        records = []
        current_months = []
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
    except:
        return pd.DataFrame()

# ③ 患者属性・エリア用（CSV名簿データ）
@st.cache_data
def load_patient_data():
    search_paths = [".", r"G:\マイドライブ\かけい整形外科ダッシュボード", r"G:\My Drive\かけい整形外科ダッシュボード"]
    all_files = []
    for d_path in search_paths:
        if os.path.exists(d_path):
            all_files.extend(glob.glob(os.path.join(d_path, "*.csv")))
    unique_files = {os.path.basename(f): f for f in all_files}
    patient_files = [f for f in unique_files.values() if "品目" not in f and re.search(r'\d{4}', f)]

    if not patient_files: return pd.DataFrame()

    combined_list = []
    for file in patient_files:
        df = None
        for enc in ['cp932', 'utf-8-sig', 'utf-8']:
            try:
                df_temp = pd.read_csv(file, encoding=enc, nrows=5)
                if '現住所 住所' in df_temp.columns and '年齢' in df_temp.columns:
                    df = pd.read_csv(file, encoding=enc)
                    break
            except: continue
        if df is None: continue

        match = re.search(r'(\d{4})[._]?(\d{2})', file)
        if match:
            df['年'] = f"{match.group(1)}年"
            df['月単体'] = f"{int(match.group(2))}月"
        else:
            match_y = re.search(r'(\d{4})', file)
            df['年'] = f"{match_y.group(1)}年" if match_y else "0000年"
            df['月単体'] = "不明"
            
        combined_list.append(df)
            
    if not combined_list: return pd.DataFrame()
    df_p = pd.concat(combined_list, ignore_index=True)

    df_p['年齢_数値'] = df_p['年齢'].astype(str).str.extract(r'(\d+)').astype(float)
    bins = [0, 9, 19, 29, 39, 49, 59, 69, 79, 89, 150]
    labels = ['0-9歳', '10-19歳', '20-29歳', '30-39歳', '40-49歳', '50-59歳', '60-69歳', '70-79歳', '80-89歳', '90歳以上']
    df_p['年齢層'] = pd.cut(df_p['年齢_数値'], bins=bins, labels=labels, right=True, include_lowest=True)

    addr = df_p['現住所 住所'].astype(str).str.replace(r'^.{2,3}[都道府県]', '', regex=True)
    area_extracted = addr.str.extract(r'([^市区町村]+[市区町村][^町村]+[町村])')[0]
    area_extracted = area_extracted.fillna(addr.str.extract(r'^([^\d０-９]+)')[0])
    df_p['エリア'] = area_extracted.str.replace(r'字.*$', '', regex=True).str.strip()

    if '患者番号' in df_p.columns:
        df_p = df_p.drop_duplicates(subset=['年', '月単体', '患者番号'])

    return df_p

def calc_diff_ratio(curr_val, prev_val):
    if prev_val == 0: return "前年データなし"
    return f"前年比 {(curr_val / prev_val) * 100:.1f}%"

month_order = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]

# ==========================================
# 4. メイン画面とナビゲーション
# ==========================================
if logo_base64:
    header_html = f"""<div class="header-container"><img src="data:image/png;base64,{logo_base64}" class="header-logo"><h1 class="header-title">かけい整形外科 経営分析ダッシュボード</h1></div>"""
    st.markdown(header_html, unsafe_allow_html=True)
else:
    st.markdown('<h1 class="header-title">かけい整形外科 経営分析ダッシュボード</h1>', unsafe_allow_html=True)

df_all = load_kakei_data()
if df_all.empty:
    st.error("KPI用のCSVデータが見つかりません。")
    st.stop()

kpi_pages = ["経営全体・主要KPI", "薬剤（院内）・院外処方", "基本診療料・医学管理料等", "調剤・処方", "注射", "処置", "手術", "検査", "画像診断", "その他", "自費", "AI総合アドバイス"]
patient_pages = ["来院分析", "患者属性推移", "エリア別推移"]
special_pages = ["クーリーフ", "体外衝撃波"]

if 'current_page' not in st.session_state: st.session_state.current_page = kpi_pages[0]

st.write("### 🔍 KPI分析メニュー")
for i in range(0, len(kpi_pages), 4):
    cols = st.columns(4)
    for j in range(4):
        if i + j < len(kpi_pages):
            p_name = kpi_pages[i + j]
            with cols[j]:
                if st.button(p_name, use_container_width=True, key=f"kpi_btn_{i+j}", type="primary" if st.session_state.current_page == p_name else "secondary"):
                    st.session_state.current_page = p_name
                    st.rerun()

st.write("### 👥 患者分析メニュー")
cols_p = st.columns(4)
for i, p_name in enumerate(patient_pages):
    with cols_p[i]:
        if st.button(p_name, use_container_width=True, key=f"patient_btn_{i}", type="primary" if st.session_state.current_page == p_name else "secondary"):
            st.session_state.current_page = p_name
            st.rerun()

st.write("### ⚡ クーリーフ・体外衝撃波分析メニュー")
cols_s = st.columns(4)
for i, p_name in enumerate(special_pages):
    with cols_s[i]:
        if st.button(p_name, use_container_width=True, key=f"special_btn_{i}", type="primary" if st.session_state.current_page == p_name else "secondary"):
            st.session_state.current_page = p_name
            st.rerun()

st.write("---")
current_page = st.session_state.current_page
available_years = sorted(df_all['年'].dropna().unique())

# ==========================================
# 5. 各ページ表示ロジック
# ==========================================

# ------------------------------------------
# 特殊分析メニュー：クーリーフ / 体外衝撃波
# ------------------------------------------
if current_page in special_pages:
    st.header(f"⚡ {current_page} 実績分析")
    
    if current_page == "クーリーフ":
        df_target = df_all[df_all['診療行為'].str.contains('末梢神経ラジオ波焼灼療法', na=False)]
    else: # 体外衝撃波
        df_target = df_all[df_all['診療行為'].str.contains('体外衝撃波疼痛治療術|衝撃波　単回|衝撃波　3回コース|衝撃波　初回限定', na=False)]
        
    if df_target.empty:
        st.warning(f"{current_page} のデータが見つかりませんでした。")
    else:
        st.write(f"#### 📈 これまでの全期間 金額・回数推移（{current_page}）")
        df_target_long = df_target.copy()
        try:
            year_int = df_target_long['年'].str.extract(r'R(\d+)年')[0].astype(int) + 2018
            month_int = df_target_long['月単体'].str.replace('月', '')
            df_target_long['年月_dt'] = pd.to_datetime(year_int.astype(str) + '/' + month_int + '/1')
            df_target_long['年月ラベル'] = df_target_long['年'] + df_target_long['月単体']
            long_agg = df_target_long.groupby(['年月_dt', '年月ラベル'], as_index=False).agg({'回数': 'sum', '金額_円': 'sum'}).sort_values('年月_dt')
            
            fig_long = make_subplots(specs=[[{"secondary_y": True}]])
            
            # ★ 改善1：金額（棒）は白・太字で棒の内側に収納
            fig_long.add_trace(go.Bar(
                x=long_agg['年月ラベル'], y=long_agg['金額_円'], name='金額(円)', marker_color='#3498DB',
                text=long_agg['金額_円'].apply(lambda x: f"<b>{x:,.0f}</b>" if x>0 else ""), 
                textposition='inside', textangle=0, textfont=dict(color='white', size=14)
            ), secondary_y=False)
            
            # ★ 改善2：回数（折れ線）は濃い赤・さらに大きく・「回」を付けて明確に分離
            fig_long.add_trace(go.Scatter(
                x=long_agg['年月ラベル'], y=long_agg['回数'], name='回数', mode='lines+markers+text', 
                line=dict(color='#E74C3C', width=3), text=long_agg['回数'].apply(lambda x: f"<b>{x:,.0f}回</b>" if x>0 else ""), 
                textposition='top center', textfont=dict(size=18, color='#800000')
            ), secondary_y=True)
            
            fig_long.update_traces(cliponaxis=False)
            fig_long.update_layout(hovermode="x unified", barmode='group', margin=dict(t=40))
            fig_long.update_yaxes(title_text="金額 (円)", secondary_y=False)
            fig_long.update_yaxes(title_text="回数", secondary_y=True)
            st.plotly_chart(fig_long, use_container_width=True)
        except Exception as e:
            st.info("全期間推移グラフの作成に失敗しました。")
            
        st.write("---")
        
        col_year, _ = st.columns(2)
        with col_year:
            selected_spec_year = st.selectbox("📅 表示する年を選択してください", available_years, index=len(available_years)-1, key="spec_year")
        
        df_curr_spec = df_target[df_target['年'] == selected_spec_year]
        
        # グラフ作成 (金額と回数)
        st.write(f"#### 📈 {selected_spec_year} 月別 金額・回数推移")
        monthly_agg = df_curr_spec.groupby('月単体').agg({'回数': 'sum', '金額_円': 'sum'}).reindex(month_order).fillna(0)
        
        fig_s = make_subplots(specs=[[{"secondary_y": True}]])
        
        # ★ 改善1：金額（棒）は白・太字で棒の内側に収納
        fig_s.add_trace(go.Bar(
            x=monthly_agg.index, y=monthly_agg['金額_円'], name='金額(円)', marker_color='#3498DB',
            text=monthly_agg['金額_円'].apply(lambda x: f"<b>{x:,.0f}</b>" if x>0 else ""), 
            textposition='inside', textangle=0, textfont=dict(color='white', size=14)
        ), secondary_y=False)
        
        # ★ 改善2：回数（折れ線）は濃い赤・さらに大きく・「回」を付けて明確に分離
        fig_s.add_trace(go.Scatter(
            x=monthly_agg.index, y=monthly_agg['回数'], name='回数', mode='lines+markers+text', 
            line=dict(color='#E74C3C', width=3), text=monthly_agg['回数'].apply(lambda x: f"<b>{x:,.0f}回</b>" if x>0 else ""), 
            textposition='top center', textfont=dict(size=18, color='#800000')
        ), secondary_y=True)
        
        fig_s.update_traces(cliponaxis=False)
        fig_s.update_layout(hovermode="x unified", barmode='group', margin=dict(t=40))
        fig_s.update_yaxes(title_text="金額 (円)", secondary_y=False)
        fig_s.update_yaxes(title_text="回数", secondary_y=True)
        st.plotly_chart(fig_s, use_container_width=True)
        
        # 詳細一覧表
        st.write(f"#### 📋 {selected_spec_year} データ詳細一覧")
        
        count_pivot = df_curr_spec.pivot_table(index='診療行為', columns='月単体', values='回数', aggfunc='sum').reindex(columns=month_order).fillna(0)
        count_pivot['年間合計回数'] = count_pivot.sum(axis=1)
        
        money_pivot = df_curr_spec.pivot_table(index='診療行為', columns='月単体', values='金額_円', aggfunc='sum').reindex(columns=month_order).fillna(0)
        money_pivot['年間合計金額'] = money_pivot.sum(axis=1)
        
        if current_page != "クーリーフ":
            sum_row_c = count_pivot.sum(numeric_only=True)
            sum_row_c.name = '★月別合計'
            count_pivot = pd.concat([count_pivot, pd.DataFrame([sum_row_c])])
            
            sum_row_m = money_pivot.sum(numeric_only=True)
            sum_row_m.name = '★月別合計'
            money_pivot = pd.concat([money_pivot, pd.DataFrame([sum_row_m])])
        
        def style_spec_table(df):
            styler = df.style.format("{:,.0f}")
            def apply_bold_total(row):
                if '★' in str(row.name): return ['font-weight: bold; background-color: #f0f2f6'] * len(row)
                return [''] * len(row)
            return styler.apply(apply_bold_total, axis=1)

        st.write("##### 🔢 実施回数")
        st.dataframe(style_spec_table(count_pivot), use_container_width=True)
        st.write("##### 💰 売上金額 (円)")
        st.dataframe(style_spec_table(money_pivot), use_container_width=True)


# ------------------------------------------
# 患者分析メニュー：来院分析
# ------------------------------------------
elif current_page == "来院分析":
    st.header("📈 患者来院動向分析")
    df_v = load_spreadsheet_visit_data()
    
    if df_v.empty:
        st.warning("スプレッドシートのデータが読み込めませんでした。")
    else:
        def get_col(keywords, default=None):
            for c in df_v.columns:
                if all(k in c for k in keywords): return c
            return default

        col_total_visit = get_col(['合計', '総患者'], '合計（月の総患者数）')
        col_new_rate = get_col(['新規', '率'], '新規患者率')
        col_avg_visit = get_col(['日平均来院数'], '１日平均来院数（人）')
        col_avg_rehab = get_col(['日平均リハビリ'], '1日平均リハビリ人数')
        
        col_subtotal = get_col(['小計'], '小計')
        col_rehab_only = get_col(['リハビリのみ'], 'リハビリのみ')
        
        new_cols = [c for c in df_v.columns if '新患' in c and '率' not in c]

        st.write("#### ① 延べ来院数・新規患者率の長期推移（2022年3月〜）")
        df_long = df_v[df_v['年月_dt'] >= '2022-03-01'].copy()
        
        if col_total_visit in df_long.columns:
            df_long = df_long[df_long[col_total_visit] > 0]
        
        fig1 = make_subplots(specs=[[{"secondary_y": True}]])
        if col_total_visit in df_long.columns:
            fig1.add_trace(go.Bar(x=df_long['年月'], y=df_long[col_total_visit], name='延べ来院数', marker_color='#3498DB'), secondary_y=False)
        if col_new_rate in df_long.columns:
            fig1.add_trace(go.Scatter(x=df_long['年月'], y=df_long[col_new_rate], name='新規患者率(%)', mode='lines+markers', line=dict(color='#E74C3C', width=3)), secondary_y=True)
        
        fig1.update_layout(hovermode="x unified", margin=dict(l=10, r=10, t=30, b=10))
        fig1.update_yaxes(title_text="延べ来院数 (人)", secondary_y=False)
        fig1.update_yaxes(title_text="新規患者率 (%)", secondary_y=True)
        st.plotly_chart(fig1, use_container_width=True)

        st.write("---")
        visit_years = sorted(df_v['年'].unique())
        col_y, _ = st.columns(2)
        with col_y:
            selected_v_year = st.selectbox("📅 分析する年度を選択してください", visit_years, index=len(visit_years)-1, key="v_year")
        
        df_curr_v = df_v[df_v['年'] == selected_v_year].copy()
        df_curr_v['月単体'] = pd.Categorical(df_curr_v['月単体'], categories=month_order, ordered=True)
        df_curr_v = df_curr_v.sort_values('月単体')

        st.write(f"#### ② {selected_v_year} 診察あり患者・リハビリのみ患者数")
        fig2 = go.Figure()
        if col_subtotal in df_curr_v.columns:
            fig2.add_trace(go.Bar(
                x=df_curr_v['月単体'], y=df_curr_v[col_subtotal], name='診察あり患者(小計)', marker_color='#2E86C1',
                text=df_curr_v[col_subtotal].apply(lambda x: f"{x:,.0f}" if x > 0 else "")
            ))
        if col_rehab_only in df_curr_v.columns:
            fig2.add_trace(go.Bar(
                x=df_curr_v['月単体'], y=df_curr_v[col_rehab_only], name='リハビリのみ', marker_color='#27AE60',
                text=df_curr_v[col_rehab_only].apply(lambda x: f"{x:,.0f}" if x > 0 else "")
            ))
        fig2.update_traces(textposition='outside', textangle=0, cliponaxis=False)
        fig2.update_layout(hovermode="x unified", barmode='group', margin=dict(t=40))
        st.plotly_chart(fig2, use_container_width=True)

        st.write(f"#### ③ {selected_v_year} 新患内訳")
        fig3 = go.Figure()
        colors_fig3 = ['#F39C12', '#8E44AD', '#D35400']
        for idx, c in enumerate(new_cols):
            if c in df_curr_v.columns:
                fig3.add_trace(go.Bar(
                    x=df_curr_v['月単体'], y=df_curr_v[c], name=c, marker_color=colors_fig3[idx % len(colors_fig3)],
                    text=df_curr_v[c].apply(lambda x: f"{x:,.0f}" if x > 0 else "")
                ))
        if not new_cols:
            st.info("※ スプレッドシートに「新患」という名称を含むデータが見つかりませんでした。")
        else:
            fig3.update_traces(textposition='outside', textangle=0, cliponaxis=False)
            fig3.update_layout(hovermode="x unified", barmode='group', margin=dict(t=40))
            st.plotly_chart(fig3, use_container_width=True)

        st.write(f"#### ④ {selected_v_year} 1日平均人数の推移")
        fig4 = go.Figure()
        if col_avg_visit in df_curr_v.columns:
            fig4.add_trace(go.Scatter(
                x=df_curr_v['月単体'], y=df_curr_v[col_avg_visit], name='1日平均来院数', 
                mode='lines+markers+text', line=dict(color='#2980B9', width=3),
                text=df_curr_v[col_avg_visit].apply(lambda x: f"{x:.1f}" if x > 0 else ""), textposition='top center'
            ))
        if col_avg_rehab in df_curr_v.columns:
            fig4.add_trace(go.Scatter(
                x=df_curr_v['月単体'], y=df_curr_v[col_avg_rehab], name='1日平均リハビリ', 
                mode='lines+markers+text', line=dict(color='#16A085', dash='dot', width=3),
                text=df_curr_v[col_avg_rehab].apply(lambda x: f"{x:.1f}" if x > 0 else ""), textposition='top center'
            ))
        fig4.update_traces(cliponaxis=False)
        fig4.update_layout(hovermode="x unified", yaxis_title="人数 (人)", margin=dict(t=40))
        st.plotly_chart(fig4, use_container_width=True)

        st.write(f"#### ⑤ {selected_v_year} 詳細データ一覧")
        target_order = [
            '診療のみ', '診療＆リハビリ', '診療のみ（新患）', '診療＆リハビリ（新患）', '小計',
            'リハビリのみ', '合計（月の総患者数）', '実稼働（1日:1 半日:0.5）', 
            '１日平均来院数（人）', '1日平均来院数（人）', 'リハビリ合計', '1日平均リハビリ人数', '新規患者率'
        ]
        exclude_cols = ['年月', '年月_dt', '年', '月単体']
        available_metrics = [c for c in df_curr_v.columns if c not in exclude_cols]
        final_order = [m for m in target_order if m in available_metrics]
        final_order += [m for m in available_metrics if m not in final_order]
        disp_df = df_curr_v.set_index('年月')[final_order].T
        
        def format_cell(val, metric_name):
            try:
                v = float(val)
                if '率' in metric_name: return f"{v:.1f}%"
                if '平均' in metric_name or '稼働' in metric_name: return f"{v:.1f}"
                return f"{v:,.0f}"
            except: return val

        for col in disp_df.columns:
            disp_df[col] = [format_cell(disp_df.at[idx, col], idx) for idx in disp_df.index]
        st.dataframe(disp_df, use_container_width=True)

# ------------------------------------------
# 患者分析メニュー：患者属性推移
# ------------------------------------------
elif current_page == "患者属性推移":
    df_p = load_patient_data()
    
    if df_p.empty:
        st.warning("患者データのCSVが見つかりません。「現住所 住所」や「年齢」を含むCSVファイルを確認してください。")
    else:
        st.header("👥 患者属性推移（年齢層別）")
        available_p_years = sorted(df_p['年'].unique().tolist())
        col_y, _ = st.columns(2)
        with col_y:
            selected_p_year = st.selectbox("📅 分析する年度を選択してください", available_p_years, index=len(available_p_years)-1, key="p_year")
        
        df_curr_p = df_p[df_p['年'] == selected_p_year].copy()
        df_curr_p['月単体'] = pd.Categorical(df_curr_p['月単体'], categories=month_order, ordered=True)
        age_labels = ['0-9歳', '10-19歳', '20-29歳', '30-39歳', '40-49歳', '50-59歳', '60-69歳', '70-79歳', '80-89歳', '90歳以上']

        st.write(f"#### 📈 {selected_p_year} 月別の年齢層別 患者数推移")
        trend_df = df_curr_p.groupby(['月単体', '年齢層']).size().unstack(fill_value=0)[age_labels]
        
        fig_trend = go.Figure()
        colors = ['#F9EBEA', '#F2D7D5', '#D98880', '#C0392B', '#922B21', '#E8DAEF', '#C39BD3', '#8E44AD', '#5B2C6F', '#212F3D']
        for idx, age in enumerate(age_labels):
            if age in trend_df.columns:
                fig_trend.add_trace(go.Bar(
                    x=trend_df.index, y=trend_df[age], name=age, marker_color=colors[idx % len(colors)]
                ))
        fig_trend.update_layout(barmode='stack', hovermode="x unified", yaxis_title="実患者数 (人)")
        st.plotly_chart(fig_trend, use_container_width=True)

        c1, c2 = st.columns([1, 1])
        with c1:
            st.write(f"#### 🥧 {selected_p_year} 年齢層の割合 (年間ユニーク)")
            df_curr_p_yearly = df_curr_p.drop_duplicates(subset=['患者番号']) if '患者番号' in df_curr_p.columns else df_curr_p
            pie_data = df_curr_p_yearly['年齢層'].value_counts().reindex(age_labels).fillna(0)
            pie_data = pie_data[pie_data > 0] 
            
            fig_pie = go.Figure(data=[go.Pie(
                labels=pie_data.index, values=pie_data.values, sort=False,
                hole=0.4, textinfo='label+percent', marker=dict(colors=colors)
            )])
            fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20))
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c2:
            st.write("#### 📋 月別・年齢層別 データ詳細")
            disp_table = trend_df.T
            disp_table['年間合計'] = disp_table.sum(axis=1)
            for col in disp_table.columns:
                disp_table[col] = disp_table[col].apply(lambda x: f"{x:,.0f} 人")
            st.dataframe(disp_table, use_container_width=True)

# ------------------------------------------
# 患者分析メニュー：エリア別推移
# ------------------------------------------
elif current_page == "エリア別推移":
    df_p = load_patient_data()
    
    if df_p.empty:
        st.warning("患者データのCSVが見つかりません。")
    else:
        st.header("🗺️ エリア別（町名単位）来院患者推移")
        available_p_years = sorted(df_p['年'].unique().tolist())
        col_y, _ = st.columns(2)
        with col_y:
            selected_p_year = st.selectbox("📅 分析する年度を選択してください", available_p_years, index=len(available_p_years)-1, key="a_year")
        
        df_curr_p = df_p[df_p['年'] == selected_p_year]
        prev_year_str = f"{int(selected_p_year.replace('年', '')) - 1}年"
        df_prev_p = df_p[df_p['年'] == prev_year_str]

        active_months_p = df_curr_p['月単体'].unique().tolist()
        df_prev_p_ytd = df_prev_p[df_prev_p['月単体'].isin(active_months_p)]
        
        if len(active_months_p) < 12 and len(active_months_p) > 0:
            st.caption(f"※表の「前年比」および「前年患者数」は、データが存在する期間（{active_months_p[0]}〜{active_months_p[-1]}）の同期間で比較しています。")

        df_curr_p_yearly = df_curr_p.drop_duplicates(subset=['患者番号']) if '患者番号' in df_curr_p.columns else df_curr_p
        df_prev_p_ytd_yearly = df_prev_p_ytd.drop_duplicates(subset=['患者番号']) if '患者番号' in df_prev_p_ytd.columns else df_prev_p_ytd

        curr_area = df_curr_p_yearly['エリア'].value_counts().reset_index()
        curr_area.columns = ['エリア', '当年患者数']
        total_curr_patients = curr_area['当年患者数'].sum()
        curr_area['割合(%)'] = (curr_area['当年患者数'] / total_curr_patients * 100) if total_curr_patients > 0 else 0
        
        prev_area = df_prev_p_ytd_yearly['エリア'].value_counts().reset_index()
        prev_area.columns = ['エリア', '前年患者数']
        
        area_df = pd.merge(curr_area, prev_area, on='エリア', how='left').fillna(0)
        area_df['前年比'] = area_df.apply(lambda x: (x['当年患者数'] / x['前年患者数'] * 100) if x['前年患者数'] > 0 else 0, axis=1)
        
        st.write(f"#### 🏆 {selected_p_year} エリア別 来院患者数 TOP30")
        top30_df = area_df.sort_values('当年患者数', ascending=False).head(30).reset_index(drop=True)
        top30_df.index += 1
        
        def style_area_table(df):
            styler = df.style.format({
                '当年患者数': "{:,.0f} 人", 
                '割合(%)': "{:.1f} %",
                '前年患者数': "{:,.0f} 人", 
                '前年比': "{:.1f} %"
            })
            def color_yoy(val):
                try:
                    v = float(val.replace('%', '').replace(' ', ''))
                    if v >= 100: return 'color: #2E86C1; font-weight: bold'
                    elif 0 < v < 100: return 'color: #E74C3C; font-weight: bold'
                    return ''
                except: return ''
            return styler.map(color_yoy, subset=['前年比'])

        st.dataframe(style_area_table(top30_df), use_container_width=True)

        st.write("---")
        st.write(f"#### 🚀 {selected_p_year} エリア別 前年比成長率 TOP10")
        st.caption("※全体の「1%以上」の患者を占める主要エリアを対象に算出しています。")
        
        valid_area_df = area_df[area_df['割合(%)'] >= 1.0]
        if valid_area_df.empty: valid_area_df = area_df 
        
        top10_yoy = valid_area_df.sort_values('前年比', ascending=False).head(10).reset_index(drop=True)
        top10_yoy.index += 1
        st.dataframe(style_area_table(top10_yoy), use_container_width=True)

# ------------------------------------------
# KPI分析メニュー：サマリーページ
# ------------------------------------------
elif current_page == "経営全体・主要KPI" or current_page == "AI総合アドバイス":
    col_year, col_dummy = st.columns(2)
    with col_year:
        selected_year = st.selectbox("📅 表示する年を選択してください", available_years, index=len(available_years)-1, key="main_year")
    
    df_curr_year = df_all[df_all['年'] == selected_year]
    prev_year_int = int(re.search(r'\d+', selected_year).group()) - 1 if re.search(r'\d+', selected_year) else 0
    df_prev_year = df_all[df_all['年'] == f"R{prev_year_int}年"]

    active_months = df_curr_year['月単体'].unique().tolist()
    df_prev_year_ytd = df_prev_year[df_prev_year['月単体'].isin(active_months)]

    st.subheader(f"📊 {selected_year} 整形外科 全体業績サマリー")
    if len(active_months) < 12 and len(active_months) > 0:
        st.caption(f"※表の「年間合計」の前年比は、データが存在する期間（{active_months[0]}〜{active_months[-1]}）の累計で比較しています。")

    curr_revenue_excl_drug = df_curr_year[df_curr_year['カテゴリ名'] != '薬剤（院内）・院外処方']['金額_円'].sum()
    prev_revenue_excl_drug_ytd = df_prev_year_ytd[df_prev_year_ytd['カテゴリ名'] != '薬剤（院内）・院外処方']['金額_円'].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("当年度 総収益 (※薬剤除く)", f"¥{int(curr_revenue_excl_drug):,}", calc_diff_ratio(curr_revenue_excl_drug, prev_revenue_excl_drug_ytd))
    
    st.write("#### 📈 月別 全体収益トレンド (※薬剤除く)")
    trend_list = []
    for mon in month_order:
        m_val = df_curr_year[(df_curr_year['月単体'] == mon) & (df_curr_year['カテゴリ名'] != '薬剤（院内）・院外処方')]['金額_円'].sum()
        p_val = df_prev_year[(df_prev_year['月単体'] == mon) & (df_prev_year['カテゴリ名'] != '薬剤（院内）・院外処方')]['金額_円'].sum()
        trend_list.append({'月': mon, '当年': m_val, '前年': p_val})
    df_t = pd.DataFrame(trend_list)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_t['月'], y=df_t['当年'], name=f"総収益({selected_year})", line=dict(color='#2E86C1', width=5)))
    if df_t['前年'].sum() > 0:
        fig.add_trace(go.Scatter(x=df_t['月'], y=df_t['前年'], name=f"前年実績", line=dict(color='#ABB2B9', width=2, dash='dot')))
    
    fig.update_layout(hovermode="x unified", xaxis={'categoryorder':'array', 'categoryarray':month_order}, yaxis_title="総収益 (円)", margin=dict(l=10, r=10, t=30, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.write("---")
    st.write("#### 📋 詳細数値データ（年間一覧）")

    if not df_curr_year.empty:
        pivot_df = df_curr_year.pivot_table(index='月単体', columns='カテゴリ名', values='金額_円', aggfunc='sum').reindex(month_order).fillna(0)
        pivot_prev = df_prev_year.pivot_table(index='月単体', columns='カテゴリ名', values='金額_円', aggfunc='sum').reindex(month_order).fillna(0) if not df_prev_year.empty else pd.DataFrame(0, index=month_order, columns=pivot_df.columns)

        total_curr_m = pivot_df.sum(axis=1)
        total_prev_m = pivot_prev.sum(axis=1)
        
        yoy_col = []
        for m in month_order:
            c, p = total_curr_m.get(m, 0), total_prev_m.get(m, 0)
            yoy_col.append(f"{(c/p*100):.1f}%" if p > 0 else ("100.0%" if c > 0 else "-"))

        sum_row_values = pivot_df.sum()
        total_revenue_sum = total_curr_m.sum()
        
        ytd_prev_total = total_prev_m[total_prev_m.index.isin(active_months)].sum()
        if ytd_prev_total > 0:
            yoy_total = f"{(total_revenue_sum / ytd_prev_total * 100):.1f}%"
        else:
            yoy_total = "-"

        pivot_df.insert(0, '総診療報酬', total_curr_m)
        pivot_df['前年比'] = yoy_col

        sum_df = pd.DataFrame([sum_row_values], columns=sum_row_values.index)
        sum_df.insert(0, '総診療報酬', total_revenue_sum)
        sum_df['前年比'] = yoy_total
        sum_df.index = ['年間合計']

        final_df = pd.concat([pivot_df, sum_df])
        rename_cols = {'基本診療料・医学管理料等': '基本・管理', '薬剤（院内）・院外処方': '薬剤', '調剤・処方': '調剤', '画像診断': '画像'}
        final_df = final_df.rename(columns=rename_cols)

        format_dict = {col: "{:,.0f}" for col in final_df.columns if col != '前年比'}
        
        def style_total_row(styler):
            return styler.apply(lambda x: ['font-weight: bold; background-color: #f8f9fa' if x.name == '年間合計' else '' for _ in x], axis=1)

        st.dataframe(style_total_row(final_df.style.format(format_dict)), use_container_width=True)
    else:
        st.info("データがありません。")

# ------------------------------------------
# KPI分析メニュー：各カテゴリページ
# ------------------------------------------
else:
    target_cat = current_page
    col_year, col_dummy = st.columns(2)
    with col_year:
        selected_year = st.selectbox("📅 表示する年を選択してください", available_years, index=len(available_years)-1)
    
    df_curr_year = df_all[df_all['年'] == selected_year]
    prev_year_int = int(re.search(r'\d+', selected_year).group()) - 1 if re.search(r'\d+', selected_year) else 0
    df_prev_year = df_all[df_all['年'] == f"R{prev_year_int}年"]

    active_months = df_curr_year['月単体'].unique().tolist()
    is_yen = ("自費" in current_page)

    def display_category_section(title_label, d_curr, d_prev):
        d_prev_ytd = d_prev[d_prev['月単体'].isin(active_months)]
        unit = "円" if is_yen else "点"
        unit_label = "金額" if is_yen else "点数"

        st.subheader(f"📊 {selected_year} 【{title_label}】 実績詳細")
        if len(active_months) < 12 and len(active_months) > 0:
            st.caption(f"※前年比は、データが存在する期間（{active_months[0]}〜{active_months[-1]}）の同期間で比較しています。")

        st.write(f"#### 🏆 年間 総{unit_label} TOPランキング")
        if not d_curr.empty:
            curr_sum = d_curr.groupby('診療行為').agg({'回数': 'sum', '総値': 'sum'}).reset_index()
            curr_sum['単価'] = (curr_sum['総値'] / curr_sum['回数']).fillna(0)
            
            prev_sum = d_prev_ytd.groupby('診療行為')['総値'].sum().reset_index().rename(columns={'総値': '前年総値'})
            rank_df = pd.merge(curr_sum, prev_sum, on='診療行為', how='left').fillna(0)
            rank_df['前年比'] = rank_df.apply(lambda x: (x['総値'] / x['前年総値']) * 100 if x['前年総値'] > 0 else 0, axis=1)
            
            top_df = rank_df.sort_values('総値', ascending=False).head(20).reset_index(drop=True)
            top_df.index += 1
            
            disp_df = top_df[['診療行為', '単価', '回数', '総値', '前年比']].rename(columns={'総値': f'総{unit_label}'})
            
            def style_top_table(df):
                styler = df.style.format({
                    '単価': "{:,.1f} " + unit, '回数': "{:,.0f} 回", 
                    f'総{unit_label}': "{:,.0f} " + unit, '前年比': "{:.1f}%"
                })
                def color_yoy(val):
                    try:
                        v = float(val.replace('%',''))
                        if v >= 100: return 'color: #2E86C1; font-weight: bold'
                        elif 0 < v < 100: return 'color: #E74C3C; font-weight: bold'
                        return ''
                    except: return ''
                return styler.map(color_yoy, subset=['前年比'])
                
            st.dataframe(style_top_table(disp_df), use_container_width=True)

        st.write("---")
        st.write("#### 📋 月別詳細テーブル（すべての診療行為・総点数）")
        if not d_curr.empty:
            matrix_df = d_curr.pivot_table(index='診療行為', columns='月単体', values='総値', aggfunc='sum').reindex(columns=month_order).fillna(0)
            matrix_df['年間合計'] = matrix_df.sum(axis=1)
            matrix_df = matrix_df.sort_values('年間合計', ascending=False)
            
            sum_row = matrix_df.sum(numeric_only=True)
            sum_row.name = f'★{unit_label}合計'
            matrix_df = pd.concat([matrix_df, pd.DataFrame([sum_row])])
            
            def style_matrix(df):
                styler = df.style.format("{:,.0f}")
                def apply_bold_total(row):
                    if '★' in str(row.name): return ['font-weight: bold; background-color: #f0f2f6'] * len(row)
                    return [''] * len(row)
                return styler.apply(apply_bold_total, axis=1)
                
            st.dataframe(style_matrix(matrix_df), use_container_width=True)

    if target_cat == "薬剤（院内）・院外処方":
        df_c_1 = df_curr_year[df_curr_year['品目分類'] == 1]
        df_p_1 = df_prev_year[df_prev_year['品目分類'] == 1]
        display_category_section("薬剤（院内）", df_c_1, df_p_1)
        
        st.write("<br><hr style='border: 3px solid #E74C3C;'><br>", unsafe_allow_html=True)
        
        df_c_14 = df_curr_year[df_curr_year['品目分類'] == 14]
        df_p_14 = df_prev_year[df_prev_year['品目分類'] == 14]
        display_category_section("院外処方", df_c_14, df_p_14)
        
    else:
        cat_df_curr = df_curr_year[df_curr_year['カテゴリ名'] == target_cat]
        cat_df_prev = df_prev_year[df_prev_year['カテゴリ名'] == target_cat]
        display_category_section(target_cat, cat_df_curr, cat_df_prev)
