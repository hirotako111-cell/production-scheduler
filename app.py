import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta

# --- ページ設定 ---
st.set_page_config(page_title="生産計画カンバン・ダッシュボード", layout="wide")
st.title("🏭 生産計画カンバン・ダッシュボード")

# --- 初期設定 ---
CURRENT_SIM_DATE = datetime(2026, 4, 2)
color_rank = { 'WE': 1, 'YL': 2, 'GD': 3, 'OE': 4, 'PINK': 5, 'RD': 6, 'GN': 7, 'BE': 8, 'PE': 9, 'VIOL': 9, 'MA': 9, 'SV': 10, 'GY': 11, 'BR': 12, 'BK': 13 }

# --- 時間計算ロジック ---
def add_working_time(start_dt, duration_mins):
    # 簡易稼働時間カレンダー（08:00-12:00, 13:00-17:00, 17:15-21:00）
    current_time = start_dt
    remaining = duration_mins
    
    while remaining > 0:
        hour = current_time.hour
        minute = current_time.minute
        
        # 休憩時間のスキップ
        if hour == 12:
            current_time = current_time.replace(hour=13, minute=0)
        elif hour == 17 and minute < 15:
            current_time = current_time.replace(minute=15)
        elif hour >= 21:
            current_time = (current_time + timedelta(days=1)).replace(hour=8, minute=0)
            continue
            
        # 次の休憩までの時間を計算
        if hour < 12: next_break = current_time.replace(hour=12, minute=0)
        elif hour < 17: next_break = current_time.replace(hour=17, minute=0)
        else: next_break = current_time.replace(hour=21, minute=0)
        
        available_mins = (next_break - current_time).total_seconds() / 60
        
        if remaining <= available_mins:
            current_time += timedelta(minutes=remaining)
            remaining = 0
        else:
            current_time = next_break
            remaining -= available_mins
            
    return current_time

# --- データ処理エンジン ---
@st.cache_data
def process_data(master_file, delivery_file, receiving_file, setup_file):
    def load_file(file_obj, skip_rows=0):
        if file_obj.name.lower().endswith('.csv'): return pd.read_csv(file_obj, skiprows=skip_rows, low_memory=False)
        else: return pd.read_excel(file_obj, skiprows=skip_rows)

    master_df = load_file(master_file, skip_rows=3)
    delivery_df = load_file(delivery_file, skip_rows=3)
    setup_speed_df = load_file(setup_file, skip_rows=0)

    for df in [master_df, delivery_df, setup_speed_df]: df.columns = df.columns.str.strip()

    machine_params = {}
    for _, row in setup_speed_df.iterrows():
        mach = str(row['工程・機械名']).strip()
        cond = str(row['生産条件（色数、木型等）']).strip()
        machine_params.setdefault(mach, {})[cond] = {
            'setup': float(row['段取り時間（分）']) if pd.notna(row['段取り時間（分）']) else 10.0,
            'speed': float(row['生産速度（枚/時）']) if pd.notna(row['生産速度（枚/時）']) else 100.0
        }

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

    def parse_date(d_str):
        if pd.isna(d_str): return None
        try: return datetime.strptime(f"{d_str.split(' ')[0]}/2026", "%d/%m/%Y")
        except: return None
    delivery_df['Delivery_Date'] = delivery_df['DUE DATE'].apply(parse_date) if 'DUE DATE' in delivery_df.columns else None

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
        if slack <= 1: rank = 'A(必達)'
        elif slack <= 3: rank = 'B(推奨)'
        else: rank = 'C(調整)'

        # 初期パラメータの取得
        mach = routing[0]
        params = machine_params.get(mach, {})
        setup_time, speed = 10.0, 100.0
        if mach.startswith('P') and mc_info['Is_DieCut'] == 1 and '木型あり' in params:
            setup_time, speed = params['木型あり']['setup'], params['木型あり']['speed']

        jobs.append({
            '優先度': rank,
            '機械': mach,
            'MCS#': mcs,
            '出荷日': d_row['Delivery_Date'].strftime('%Y-%m-%d') if d_row['Delivery_Date'] else '2099-12-31',
            '数量': int(order_qty),
            '段取り(分)': setup_time,
            '基準速度(枚/時)': speed,
            'エラー状態': '正常'
        })

    df_jobs = pd.DataFrame(jobs)
    df_jobs.sort_values(by=['優先度'], inplace=True)
    df_jobs.insert(0, '実行順', range(1, len(df_jobs) + 1))
    return df_jobs, machine_params

# --- フロントエンドUI ---
with st.sidebar:
    st.header("📂 データアップロード")
    f_master = st.file_uploader("1. MasterCard", type=['csv', 'xlsx'])
    f_delivery = st.file_uploader("2. Delivery", type=['csv', 'xlsx'])
    f_recv = st.file_uploader("3. Receiving", type=['csv', 'xlsx'])
    f_setup = st.file_uploader("4. Setup Speed", type=['csv', 'xlsx'])

    st.markdown("---")
    st.header("🛠 シミュレーション設定")
    sim_speed_up = st.checkbox("⚡ 応援投入（全機械 生産速度20%UP）", value=False)
    sim_no_overtime = st.checkbox("🚫 残業なし（17:00終了）", value=False)

if f_master and f_delivery and f_recv and f_setup:
    if "raw_data" not in st.session_state:
        df, m_params = process_data(f_master, f_delivery, f_recv, f_setup)
        st.session_state.raw_data = df
        st.session_state.machine_params = m_params

    # 1. 機械フィルター
    machine_list = ["すべて"] + sorted(st.session_state.raw_data['機械'].unique().tolist())
    selected_machine = st.selectbox("🎯 対象の機械を選択して予定を表示", machine_list)

    # フィルタリング
    if selected_machine == "すべて":
        display_df = st.session_state.raw_data.copy()
    else:
        display_df = st.session_state.raw_data[st.session_state.raw_data['機械'] == selected_machine].copy()

    # 2. 編集機能（実行順、数量の変更）
    st.markdown(f"### 📋 【{selected_machine}】の計画調整 (実行順・数量を編集可能)")
    edited_df = st.data_editor(
        display_df,
        column_config={
            "実行順": st.column_config.NumberColumn("実行順 (並べ替え)", step=1),
            "数量": st.column_config.NumberColumn("数量 (分納調整)", min_value=0, step=100),
            "機械": st.column_config.TextColumn(disabled=True),
            "優先度": st.column_config.TextColumn(disabled=True),
            "MCS#": st.column_config.TextColumn(disabled=True),
            "出荷日": st.column_config.TextColumn(disabled=True),
            "段取り(分)": None, # 非表示
            "基準速度(枚/時)": None, # 非表示
            "エラー状態": None # 非表示
        },
        hide_index=True,
        use_container_width=True
    )

    # 3. リアルタイム再計算ロジック
    edited_df.sort_values(by='実行順', inplace=True)
    
    start_dt = CURRENT_SIM_DATE.replace(hour=8, minute=0)
    current_times = {} # 機械ごとの現在時刻管理
    
    calculated_records = []
    delay_count = 0

    for _, row in edited_df.iterrows():
        mach = row['機械']
        if mach not in current_times: current_times[mach] = start_dt
        
        # シミュレーション適用
        speed = row['基準速度(枚/時)'] * (1.2 if sim_speed_up else 1.0)
        duration = row['段取り(分)'] + (row['数量'] / speed) * 60
        
        c_start = current_times[mach]
        c_end = add_working_time(c_start, duration)
        
        # エラー検知（納期遅延）
        due_date_str = row['出荷日']
        if due_date_str != '2099-12-31':
            due_dt = datetime.strptime(due_date_str, "%Y-%m-%d").replace(hour=23, minute=59)
            if c_end > due_dt:
                status = "🚨 納期遅延"
                delay_count += 1
            else:
                status = "✅ 正常"
        else: status = "✅ 正常"

        row['開始予定'] = c_start.strftime("%m/%d %H:%M")
        row['終了予定'] = c_end.strftime("%m/%d %H:%M")
        row['ステータス'] = status
        
        calculated_records.append(row)
        
        # 次のジョブは30分後
        current_times[mach] = add_working_time(c_end, 30)

    final_df = pd.DataFrame(calculated_records)

    # 4. サマリー表示 (画面上部へ配置)
    st.markdown("---")
    st.markdown("### 📊 稼働状況サマリー")
    col1, col2, col3 = st.columns(3)
    col1.metric("総予定数量", f"{final_df['数量'].sum():,} 枚")
    
    if delay_count > 0:
        col2.error(f"⚠️ 納期遅延リスク: {delay_count} 件")
    else:
        col2.success("✨ 納期遅延リスク: 0 件")
        
    col3.metric("シミュレーション状態", "応援+20%" if sim_speed_up else "標準速度")

    # 5. 計算結果の最終確認表 (ステータス・時刻入り)
    st.markdown("#### 🕒 時系列スケジュール結果")
    
    # 納期遅延行をハイライトする処理
    def highlight_errors(s):
        return ['background-color: #ffcccc' if s['ステータス'] == '🚨 納期遅延' else '' for _ in s]
    
    st.dataframe(final_df[['実行順', '機械', 'MCS#', '数量', '出荷日', '開始予定', '終了予定', 'ステータス']].style.apply(highlight_errors, axis=1), use_container_width=True)

else:
    st.info("👈 左側のサイドバーからデータファイルをアップロードしてください。")
