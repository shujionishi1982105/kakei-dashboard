import streamlit as st
import pandas as pd
import plotly.graph_objects as go
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
        if user_input == USER_ID and pw_input == USER_PASS:
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
    /* ========== 共通設定（PC用ベース） ========== */
    [data-testid="stTable"] th:first-child { min-width: 100px !important; }
    div[data-baseweb="select"], div[data-baseweb="select"] * { cursor: pointer !important; }

    div.stButton > button {
        height: 65px !important;
        border-radius: 10px !important;
        font-weight: bold !important;
        font-size: 15px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.2s ease-in-out !important;
        white-space: normal !important;
    }
    div.stButton > button[kind="secondary"] { background-color: #EBF5FB !important; border: 1px solid #AED6F1 !important; color: #154360 !important; }
    div.stButton > button[kind="primary"] { background-color: #E74C3C !important; border: 1px solid #E74C3C !important; color: #FFFFFF !important; }
    div.stButton > button:hover { transform: translateY(-2px) !important; filter: brightness(0.95) !important; }

    .header-container {
        display: flex;
        align-items: center; 
        gap: 15px; 
        margin-bottom: 25px;
        margin-top: 10px;
    }
    .header-logo { height: 55px; width: auto; object-fit: contain; }
    .header-title { margin: 0 !important; color: #2C3E50; font-size: 32px; font-weight: bold; }

    div[data-testid="stMetricValue"] {
        font-size: 2rem !important;
        white-space: nowrap !important;
        overflow: visible !important;
    }
    div[data-testid="stMetricValue"] > div {
        overflow: visible !important;
        text-overflow: clip !important;
    }

    /* ========== 📱スマホ用自動切り替え設定 ========== */
    @media screen and (max-width: 768px) {
        .header-container {
            flex-direction: column !important;
            gap: 5px !important;
            margin-bottom: 10px !important;
        }
        .header-logo { height: 40px !important; }
        .header-title { 
            font-size: 20px !important; 
            text-align: center !important; 
            border-bottom: 2px solid #E74C3C !important;
            padding-bottom: 5px !important;
        }
        div.stButton > button {
            height: 45px !important; 
            font-size: 13px !important; 
            padding: 2px !important;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.5rem !important;
        }
        [data-testid="stTable"] th:first-child { 
            min-width: 70px !important; 
        }
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 3. データ読み込み（デスクトップ/GitHub共通）
# ==========================================
@st.cache_data
def load_kakei_data():
    # デプロイ（GitHub）環境とデスクトップ環境の両方に対応
    search_paths = [
        ".",  # app.py と同じ場所（GitHub上やデスクトップで動かすため一番重要）
        r"G:\マイドライブ\かけい整形外科ダッシュボード",
        r"G:\My Drive\かけい整形外科ダッシュボード"
    ]
    
    all_files = []
    for d_path in search_paths:
        if os.path.exists(d_path):
            all_files.extend(glob.glob(os.path.join(d_path, "*.csv")))
    
    unique_files = {os.path.basename(f): f for f in all_files}
    target_files = [f for f in unique_files.values() if "品目" in f]

    if not target_files: return pd.DataFrame()

    CATEGORY_MAP = {
        1: '薬剤（院内）・院外処方', 
        3: '基本診療料・医学管理料等', 
        4: '調剤・処方',
        5: '注射', 
        6: '処置', 
        7: '手術', 
        8: '検査',
        9: '画像診断', 
        10: 'その他', 
        11: '自費', 
        14: '薬剤（院内）・院外処方'
    }

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
        rename_map = {'品名': '診療行為', '価格１': '単価', '合計　使用量': '回数', '合計 使用量': '回数'}
        df = df.rename(columns=rename_map)
        
        if '診療行為' not in df.columns or '品目分類' not in df.columns: 
            continue

        df = df[~df['品目分類'].isin([2, 12, 13])]
        df['カテゴリ名'] = df['品目分類'].map(CATEGORY_MAP)

        df['単価'] = pd.to_numeric(df['単価'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        df['回数'] = pd.to_numeric(df['回数'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        # ★【修正箇所】ベースアップ評価料など、単価が違うのに名前が同じ項目を分離する
        mask = df['診療行為'].str.contains('外来・在宅ベースアップ評価料', na=False)
        df.loc[mask, '診療行為'] = df.loc[mask, '診療行為'].astype(str) + '（' + df.loc[mask, '単価'].astype(int).astype(str) + '点）'

        df['総値'] = df['単価'] * df['回数']
        
        df['金額_円'] = df['総値'] * 10
        df.loc[df['品目分類'] == 11, '金額_円'] = df.loc[df['品目分類'] == 11, '総値']
        
        match = re.search(r'\d{6}', file)
        month_raw = match.group() if match else "000000"
        year, month = int(month_raw[:4]), int(month_raw[4:])
        df['月単体'] = f"{month}月"
        df['年'] = f"R{year-2018}年"
        combined_list.append(df)
            
    return pd.concat(combined_list, ignore_index=True) if combined_list else pd.DataFrame()

def calc_diff_ratio(curr_val, prev_val):
    if prev_val == 0: return "前年データなし"
    return f"前年比 {(curr_val / prev_val) * 100:.1f}%"

month_order = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
# ==========================================
# 4. メイン画面とナビゲーション
# ==========================================
if logo_base64:
    header_html = f"""
    <div class="header-container">
        <img src="data:image/png;base64,{logo_base64}" class="header-logo">
        <h1 class="header-title">かけい整形外科 経営分析ダッシュボード</h1>
    </div>
    """
    st.markdown(header_html, unsafe_allow_html=True)
else:
    st.markdown('<h1 class="header-title">かけい整形外科 経営分析ダッシュボード</h1>', unsafe_allow_html=True)

df_all = load_kakei_data()
if df_all.empty:
    st.error("データが見つかりません。")
    st.stop()

pages = [
    "経営全体・主要KPI", 
    "薬剤（院内）・院外処方", 
    "基本診療料・医学管理料等", 
    "調剤・処方",
    "注射", 
    "処置", 
    "手術", 
    "検査", 
    "画像診断", 
    "その他", 
    "自費", 
    "AI総合アドバイス"
]

if 'current_page' not in st.session_state: st.session_state.current_page = pages[0]

st.write("### 🔍 分析メニュー")
for i in range(0, len(pages), 4):
    cols = st.columns(4)
    for j in range(4):
        if i + j < len(pages):
            p_name = pages[i + j]
            with cols[j]:
                btn_type = "primary" if st.session_state.current_page == p_name else "secondary"
                if st.button(p_name, use_container_width=True, key=f"nav_btn_{i+j}", type=btn_type):
                    st.session_state.current_page = p_name
                    st.rerun()
st.write("---")

current_page = st.session_state.current_page
available_years = sorted(df_all['年'].dropna().unique())

# ==========================================
# 5. 個別ページの表示内容
# ==========================================
if current_page == "経営全体・主要KPI" or current_page == "AI総合アドバイス":
    
    col_year, col_dummy = st.columns(2)
    with col_year:
        selected_year = st.selectbox("📅 表示する年を選択してください", available_years, index=len(available_years)-1, key="main_year")
    
    df_curr_year = df_all[df_all['年'] == selected_year]
    prev_year_int = int(re.search(r'\d+', selected_year).group()) - 1 if re.search(r'\d+', selected_year) else 0
    df_prev_year = df_all[df_all['年'] == f"R{prev_year_int}年"]

    active_months = df_curr_year['月単体'].unique().tolist()
    # 前年比計算用の「前年同期間」のデータ
    df_prev_year_ytd = df_prev_year[df_prev_year['月単体'].isin(active_months)]

    st.subheader(f"📊 {selected_year} 整形外科 全体業績サマリー")
    if len(active_months) < 12 and len(active_months) > 0:
        st.caption(f"※表の「年間合計」の前年比は、データが存在する期間（{active_months[0]}〜{active_months[-1]}）の累計で比較しています。")

    c1, c2, c3 = st.columns(3)
    
    c1.metric("当年度 総収益", f"¥{int(df_curr_year['金額_円'].sum()):,}", calc_diff_ratio(df_curr_year['金額_円'].sum(), df_prev_year_ytd['金額_円'].sum()))
    
    rehab_curr = df_curr_year[df_curr_year['診療行為'].str.contains('リハビリ|運動器|物理療法', na=False)]['金額_円'].sum()
    rehab_prev = df_prev_year_ytd[df_prev_year_ytd['診療行為'].str.contains('リハビリ|運動器|物理療法', na=False)]['金額_円'].sum()
    dexa_curr = df_curr_year[df_curr_year['診療行為'].str.contains('骨密度|DEXA', na=False)]['金額_円'].sum()
    dexa_prev = df_prev_year_ytd[df_prev_year_ytd['診療行為'].str.contains('骨密度|DEXA', na=False)]['金額_円'].sum()
    
    c2.metric("リハビリ料合計", f"¥{int(rehab_curr):,}", calc_diff_ratio(rehab_curr, rehab_prev))
    c3.metric("骨密度(DEXA)", f"¥{int(dexa_curr):,}", calc_diff_ratio(dexa_curr, dexa_prev))
    
    st.write("#### 📈 月別 全体収益トレンド")
    trend_list = []
    for mon in month_order:
        m_val = df_curr_year[df_curr_year['月単体'] == mon]['金額_円'].sum()
        p_val = df_prev_year[df_prev_year['月単体'] == mon]['金額_円'].sum()
        trend_list.append({'月': mon, '当年': m_val, '前年': p_val})
    df_t = pd.DataFrame(trend_list)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_t['月'], y=df_t['当年'], name=f"総収益({selected_year})", line=dict(color='#2E86C1', width=5)))
    if df_t['前年'].sum() > 0:
        fig.add_trace(go.Scatter(x=df_t['月'], y=df_t['前年'], name=f"前年実績", line=dict(color='#ABB2B9', width=2, dash='dot')))
    
    fig.update_layout(
        hovermode="x unified", xaxis={'categoryorder':'array', 'categoryarray':month_order}, yaxis_title="総収益 (円)",
        margin=dict(l=10, r=10, t=30, b=10)
    )
    st.plotly_chart(fig, use_container_width=True)

    # ★★★ 【今回修正：合計行とYTD前年比の追加】 ★★★
    st.write("---")
    st.write("#### 📋 詳細数値データ（年間一覧）")

    if not df_curr_year.empty:
        # 1. 当年の月別・カテゴリ別集計
        pivot_df = df_curr_year.pivot_table(
            index='月単体', columns='カテゴリ名', values='金額_円', aggfunc='sum'
        ).reindex(month_order).fillna(0)

        # 2. 前年の月別・カテゴリ別集計
        pivot_prev = df_prev_year.pivot_table(
            index='月単体', columns='カテゴリ名', values='金額_円', aggfunc='sum'
        ).reindex(month_order).fillna(0) if not df_prev_year.empty else pd.DataFrame(0, index=month_order, columns=pivot_df.columns)

        # 3. 各月の総収益と前年比の計算
        total_curr_m = pivot_df.sum(axis=1)
        total_prev_m = pivot_prev.sum(axis=1)
        
        yoy_col = []
        for m in month_order:
            c, p = total_curr_m.get(m, 0), total_prev_m.get(m, 0)
            yoy_col.append(f"{(c/p*100):.1f}%" if p > 0 else ("100.0%" if c > 0 else "-"))

        # 4. 「年間合計」行の作成
        # 数値列の単純合計
        sum_row_values = pivot_df.sum()
        total_revenue_sum = total_curr_m.sum()
        
        # 合計行の前年比（YTD: 年初から現在までの累計比較）
        # 前年の同期間のみを合計して比較する
        ytd_prev_total = total_prev_m[total_prev_m.index.isin(active_months)].sum()
        if ytd_prev_total > 0:
            yoy_total = f"{(total_revenue_sum / ytd_prev_total * 100):.1f}%"
        else:
            yoy_total = "-"

        # 5. 表の組み立て
        pivot_df.insert(0, '総診療報酬', total_curr_m)
        pivot_df['前年比'] = yoy_col

        # 合計行のデータフレーム作成
        sum_df = pd.DataFrame([sum_row_values], columns=sum_row_values.index)
        sum_df.insert(0, '総診療報酬', total_revenue_sum)
        sum_df['前年比'] = yoy_total
        sum_df.index = ['年間合計']

        # 結合
        final_df = pd.concat([pivot_df, sum_df])

        # 6. 表示用のリネーム（短縮名）
        rename_cols = {
            '基本診療料・医学管理料等': '基本・管理',
            '薬剤（院内）・院外処方': '薬剤',
            '調剤・処方': '調剤',
            '画像診断': '画像'
        }
        final_df = final_df.rename(columns=rename_cols)

        # 7. スタイル適用（カンマ区切り、合計行の太字化）
        format_dict = {col: "{:,.0f}" for col in final_df.columns if col != '前年比'}
        
        def style_total_row(styler):
            # 最後の行（年間合計）を太字、背景を少しグレーに
            return styler.apply(lambda x: ['font-weight: bold; background-color: #f8f9fa' if x.name == '年間合計' else '' for _ in x], axis=1)

        st.dataframe(
            style_total_row(final_df.style.format(format_dict)),
            use_container_width=True
        )
    else:
        st.info("データがありません。")

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

        # ランキング表
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

        # グラフ
        st.write("---")
        st.write("### 📈 診療行為別 月別推移")
        
        if not d_curr.empty:
            item_order = d_curr.groupby('診療行為')['総値'].sum().sort_values(ascending=False).index.tolist()
            missing_prev = [x for x in d_prev['診療行為'].unique() if x not in item_order]
            all_items = ["🌟 総額"] + item_order + missing_prev
        else:
            all_items = ["🌟 総額"] + sorted(list(d_prev['診療行為'].unique()))
        
        if all_items:
            selected_item = st.selectbox(f"🔍 グラフ表示する項目を選択してください ({title_label})", all_items, key=f"sel_{title_label}")
            
            if selected_item == "🌟 総額":
                curr_item_data = d_curr.groupby('月単体')['総値'].sum().reindex(month_order).fillna(0)
                prev_item_data = d_prev.groupby('月単体')['総値'].sum().reindex(month_order).fillna(0)
            else:
                curr_item_data = d_curr[d_curr['診療行為'] == selected_item].groupby('月単体')['総値'].sum().reindex(month_order).fillna(0)
                prev_item_data = d_prev[d_prev['診療行為'] == selected_item].groupby('月単体')['総値'].sum().reindex(month_order).fillna(0)
            
            plot_df = pd.DataFrame({'月': month_order, '当年': curr_item_data.values, '前年': prev_item_data.values})
            plot_df['前年比'] = plot_df.apply(lambda x: (x['当年'] / x['前年'] * 100) if x['前年'] > 0 else 0, axis=1)
            colors_act = ['#E74C3C' if 0 < val < 100 else ('#2E86C1' if val >= 100 else '#000000') for val in plot_df['前年比']]

            fig_act = go.Figure()
            fig_act.add_trace(go.Scatter(
                x=plot_df['月'], y=plot_df['当年'], mode='lines+markers+text', name=f'当年 ({selected_year})',
                line=dict(color='#2E86C1', width=4),
                text=plot_df['前年比'].apply(lambda x: f"{x:.1f}%" if x > 0 else ""),
                textposition="top center", textfont=dict(color=colors_act, size=13, family="Arial Black"),
                hovertemplate=f"<b>%{{x}}</b><br>当年: %{{y:,.0f}} {unit}<extra></extra>"
            ))
            if plot_df['前年'].sum() > 0:
                fig_act.add_trace(go.Scatter(
                    x=plot_df['月'], y=plot_df['前年'], mode='lines+markers', name=f'前年実績',
                    line=dict(color='#ABB2B9', width=2, dash='dot'),
                    hovertemplate=f"前年: %{{y:,.0f}} {unit}<extra></extra>"
                ))

            fig_act.update_layout(
                hovermode="x unified", xaxis_title="診療月", yaxis_title=f"総{unit_label} ({unit})",
                margin=dict(l=10, r=10, t=30, b=10)
            )
            st.plotly_chart(fig_act, use_container_width=True)

        # マトリクス表
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
