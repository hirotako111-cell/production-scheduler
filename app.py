import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime

# --- ページ設定 ---
st.set_page_config(page_title="生産計画カンバン・ダッシュボード", layout="wide")
st.title("🏭 生産計画カンバン・ダッシュボード")
st.markdown("基幹システムから出力した本日の各種CSVデータをサイドバーからアップロードしてください。")

# --- 初期設定 ---
# 本日の日付（※シミュレーション基準日）
CURRENT_SIM_DATE = datetime(2026, 4, 2)
color_rank = { 'WE': 1, 'YL': 2, 'GD': 3, 'OE': 4, 'PINK': 5, 'RD': 6, 'GN': 7, 'BE': 8, 'PE': 9, 'VIOL': 9, 'MA': 9, 'SV': 10, 'GY': 11, 'BR': 12, 'BK': 13 }

# --- データ処理エンジン ---
@st.cache_data
def process_data(master_file, delivery_file, receiving_file, setup_file):
    try:
        # アップロードされたファイル群を直接読み込む
        master_df = pd.read_csv(master_file, skiprows=3, low_memory=False)
        delivery_df = pd.read_csv(delivery_file, skiprows=3, low_memory=False)
        receiving_df = pd.read_csv(receiving_file, skiprows=4, low_memory=False)
        setup_speed_df = pd.read_csv(setup_file, low_memory=False)

        # カラム名の空白除去
        for df in [master_df, delivery_df, receiving_df, setup_speed_df]:
            df.columns = df.columns.str.strip()

        # 段取り・速度の辞書化
        machine_params = {}
        for _, row in setup_speed_df.iterrows():
            mach = str(row['工程・機械名']).strip()
            cond = str(row['生産条件（色数、木型等）']).strip()
            setup = float(row['段取り時間（分）']) if pd.notna(row['段取り時間（分）']) else 10.0
            speed = float(row['生産速度（枚/時）']) if pd.notna(row['生産速度（枚/時）']) else 100.0
            if mach not in machine_params: machine_params[mach] = {}
            machine_params[mach][cond] = {'setup': setup, 'speed': speed}

        # インク色・木型判定
        def get_color_score(row):
            score = 99
            for i in range(1, 8):
                col = f'COLOR {i}'
                if col in row and pd.notna(row[col]):
                    match = re.search(r'([A-Za-z]+)$', str(row[col]).strip())
                    if match and match.group(1).upper() in color_rank:
                        score = min(score, color_rank[match.group(1).upper()])
            return score

        master_df_unique = master_df.drop_duplicates(subset=['MCS#']).copy()
        master_df_unique['Color_Score'] = master_df_unique.apply(get_color_score, axis=1)
        master_df_unique['Is_DieCut'] = master_df_unique['PD'].apply(lambda x: 1 if isinstance(x, str) and ('DC' in x or 'D/C' in x) else 0)
        mc_dict = master_df_unique.set_index('MCS#').to_dict('index')

        def parse_date(d_str, year="2026"):
            if pd.isna(d_str): return None
            try:
                if " " in d_str: d_str = d_str.split(' ')[0]
                return datetime.strptime(f"{d_str}/{year}", "%d/%m/%Y")
            except: return None

        delivery_df['Delivery_Date'] = delivery_df['DUE DATE'].apply(parse_date) if 'DUE DATE' in delivery_df.columns else None

        # ジョブ抽出と優先度計算
        jobs = []
        for _, d_row in delivery_df.dropna(subset=['MCS#']).iterrows():
            mcs = d_row['MCS#']
            if mcs not in mc_dict: continue
            order_qty = float(d_row['ORDER']) if 'ORDER' in d_row and pd.notna(d_row['ORDER']) else 0
            if order_qty <= 0: continue
            
            mc_info = mc_dict[mcs]
            routing = [str(mc_info[f'MSP{i}']).strip() for i in range(1, 13) if f'MSP{i}' in mc_info and pd.notna(mc_info[f'MSP{i}'])]
            routing = [m for m in routing if m != 'CORR' and m != '']
            if not routing: continue

            slack = (d_row['Delivery_Date'] - CURRENT_SIM_DATE).days if d_row['Delivery_Date'] else 99
            if slack <= 1: rank = 'A (必達)'
            elif slack <= 3: rank = 'B (推奨)'
            else: rank = 'C (調整)'

            jobs.append({
                '優先度': rank,
                'MCS#': mcs,
                '機械': routing[0] if routing else '不明', 
                '出荷日': d_row['Delivery_Date'].strftime('%m/%d') if d_row['Delivery_Date'] else '-',
                '数量': int(order_qty),
                '木型': '有' if mc_info['Is_DieCut'] == 1 else '無',
                'Flute': str(mc_info.get('FLUTE', '')),
                'ColorScore': mc_info['Color_Score']
            })

        df_jobs = pd.DataFrame(jobs)
        df_jobs.sort_values(by=['優先度', '木型', 'Flute', 'ColorScore'], ascending=[True, False, True, True], inplace=True)
        df_jobs.drop(columns=['ColorScore'], inplace=True)
        df_jobs.insert(0, '実行順', range(1, len(df_jobs) + 1))
        
        return df_jobs, machine_params, None
    except Exception as e:
        return None, None, str(e)

# --- フロントエンドUI ---
with st.sidebar:
    st.header("📂 データアップロード")
    st.caption("※アップロードされたデータは一時的なものであり、システムに保存されません。")
    f_master = st.file_uploader("1. MasterCard (CSV)", type=['csv'])
    f_delivery = st.file_uploader("2. Delivery (CSV)", type=['csv'])
    f_recv = st.file_uploader("3. Receiving (CSV)", type=['csv'])
    f_setup = st.file_uploader("4. Setup Speed (CSV)", type=['csv'])

if f_master and f_delivery and f_recv and f_setup:
    if "schedule_data" not in st.session_state:
        with st.spinner('スケジュールを自動生成しています...'):
            initial_df, m_params, error_msg = process_data(f_master, f_delivery, f_recv, f_setup)
            if error_msg:
                st.error(f"ファイルの読み込みエラー: {error_msg}")
            else:
                st.session_state.schedule_data = initial_df
                st.session_state.machine_params = m_params
                st.success("データの読み込みとスケジュールの自動生成が完了しました！")

    if "schedule_data" in st.session_state:
        st.markdown("### 📊 本日の稼働状況サマリー (P2機械)")
        col1, col2, col3 = st.columns(3)
        p2_jobs = st.session_state.schedule_data[st.session_state.schedule_data['機械'] == 'P2']
        total_qty = p2_jobs['数量'].sum()
        must_do_count = len(p2_jobs[p2_jobs['優先度'] == 'A (必達)'])
        
        col1.metric("P2 総生産予定数", f"{total_qty:,} 枚")
        col2.metric("【必達】タスク数", f"{must_do_count} 件")
        col3.metric("本日の稼働設定", "通常 + 残業 (21:00迄)")

        st.markdown("### 📋 計画の調整 (ドラッグ＆ドロップの代わりに数値を直接編集できます)")
        st.info("💡 **実行順** の数字を書き換えて列名をクリックすると並び替わります。**数量**を減らすと分納扱いになります。")
        
        edited_df = st.data_editor(
            st.session_state.schedule_data,
            column_config={
                "実行順": st.column_config.NumberColumn("実行順", help="数字を変更して並び替えます", step=1),
                "数量": st.column_config.NumberColumn("数量", min_value=0, step=100),
                "優先度": st.column_config.TextColumn("優先度", disabled=True),
                "MCS#": st.column_config.TextColumn("MCS#", disabled=True),
                "出荷日": st.column_config.TextColumn("出荷日", disabled=True),
            },
            hide_index=True,
            use_container_width=True,
            height=500
        )

        st.session_state.schedule_data = edited_df

        st.markdown("---")
        csv = edited_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 確定したスケジュールをCSVでダウンロード",
            data=csv,
            file_name=f'production_schedule_{datetime.now().strftime("%Y%m%d")}.csv',
            mime='text/csv',
            type="primary"
        )
else:
    st.info("👈 左側のサイドバーから、本日の各種CSVファイル（4つ）をすべてアップロードしてください。")
