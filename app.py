import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import altair as alt

# --- ページ設定 ---
st.set_page_config(page_title="生産計画調整ダッシュボード", layout="wide")
st.title("🏭 生産計画調整ダッシュボード")

# --- 初期設定 ---
color_rank = { 'WE': 1, 'YL': 2, 'GD': 3, 'OE': 4, 'PINK': 5, 'RD': 6, 'GN': 7, 'BE': 8, 'PE': 9, 'VIOL': 9, 'MA': 9, 'SV': 10, 'GY': 11, 'BR': 12, 'BK': 13 }

# --- 時間計算ロジック（休憩除外） ---
def add_working_time(start_dt, duration_mins):
    current_time = start_dt
    remaining = duration_mins
    while remaining > 0:
        hour, minute = current_time.hour, current_time.minute
        if hour == 12: current_time = current_time.replace(hour=13, minute=0)
        elif hour == 17 and minute < 15: current_time = current_time.replace(minute=15)
        elif hour >= 21:
            current_time = (current_time + timedelta(days=1)).replace(hour=8, minute=0)
            continue
            
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
def process_data(master_file, delivery_file, setup_file, track_file, recv_file, target_date):
    def load_file(file_obj, skip_rows=0):
        if file_obj.name.lower().endswith('.csv'): return pd.read_csv(file_obj, skiprows=skip_rows, low_memory=False)
        else: return pd.read_excel(file_obj, skiprows=skip_rows)

    def find_col(df, keywords):
        for col in df.columns:
            if any(k in col.upper() for k in keywords): return col
        return None

    # 各種ファイルの読み込み
    master_df = load_file(master_file, skip_rows=3)
    delivery_df = load_file(delivery_file, skip_rows=3)
    setup_speed_df = load_file(setup_file, skip_rows=0)
    
    for df in [master_df, delivery_df, setup_speed_df]: df.columns = df.columns.str.strip()

    # MSC# への修正対応を含めた列名特定
    mcs_col_delivery = find_col(delivery_df, ['MCS#', 'MSC#', 'M/CARD SEQ#'])
    if not mcs_col_delivery:
        st.error(f"Deliveryファイル内に 'MCS#' または 'MSC#' 列が見つかりません。列名を確認してください。")
        st.stop()

    # 材料受入データの読み込み
    recv_dict = {}
    if recv_file:
        recv_df = load_file(recv_file, skip_rows=4)
        recv_df.columns = recv_df.columns.str.strip()
        mcs_col_recv = find_col(recv_df, ['MCS#', 'MSC#', 'PART'])
        status_col_recv = find_col(recv_df, ['STATUS', 'RECV'])
        if mcs_col_recv and status_col_recv:
            recv_dict = recv_df.set_index(mcs_col_recv)[status_col_recv].to_dict()

    machine_params = {}
    for _, row in setup_speed_df.iterrows():
        mach = str(row['工程・機械名']).strip()
        cond = str(row['生産条件（色数、木型等）']).strip()
        machine_params.setdefault(mach, {})[cond] = {
            'setup': float(row['段取り時間（分）']) if pd.notna(row['段取り時間（分）']) else 10.0,
            'speed': float(row['生産速度（枚/時）']) if pd.notna(row['生産速度（枚/時）']) else 100.0
        }

    mc_dict = master_df.drop_duplicates(subset=['MCS#']).set_index('MCS#').to_dict('index')

    def get_routing(mcs):
        if mcs not in mc_dict: return None
        mc_info = mc_dict[mcs]
        routing = [str(mc_info[f'MSP{i}']).strip() for i in range(1, 13) if f'MSP{i}' in mc_info and pd.notna(mc_info[f'MSP{i}'])]
        return routing[0] if routing else None

    jobs = []
    
    # 1. 実績ファイル(Floor Track)から「未完了分」を抽出
    if track_file:
        track_df = load_file(track_file, skip_rows=4)
        track_df.columns = track_df.columns.str.strip()
        mcs_col_track = find_col(track_df, ['MCS#', 'MSC#'])
        for _, t_row in track_df.iterrows():
            mcs = str(t_row.get(mcs_col_track, '')).strip()
            if not mcs or mcs == 'nan': continue
            remaining = float(t_row.get('PLAN OUT', 0)) - float(t_row.get('GOOD', 0))
            if remaining > 0:
                mach = get_routing(mcs)
                if not mach: continue
                jobs.append({
                    '優先度': 'A (前日繰越)', '機械': mach, 'MCS#': mcs, 
                    '出荷日': '繰越分', '数量': int(remaining), '材料': '入荷済'
                })

    # 2. 新規Deliveryデータの抽出
    def parse_date(d_str):
        if pd.isna(d_str): return None
        try: return datetime.strptime(f"{d_str.split(' ')[0]}/{target_date.year}", "%d/%m/%Y")
        except: return None
    delivery_df['Delivery_Date'] = delivery_df['DUE DATE'].apply(parse_date) if 'DUE DATE' in delivery_df.columns else None

    for _, d_row in delivery_df.dropna(subset=[mcs_col_delivery]).iterrows():
        mcs = str(d_row[mcs_col_delivery]).strip()
        order_qty = float(d_row['ORDER']) if 'ORDER' in d_row and pd.notna(d_row['ORDER']) else 0
        if order_qty <= 0: continue
        mach = get_routing(mcs)
        if not mach: continue
        
        # 材料受入ステータスの確認
        mat_status = "入荷済" if recv_dict.get(mcs) == "RECEIVED" else "未入荷" if mcs in recv_dict else "要確認"

        ddate = d_row['Delivery_Date']
        slack = (ddate - target_date).days if ddate else 99
        if slack <= 1: rank = 'B (本日/翌日必達)'
        elif slack <= 3: rank = 'C (推奨)'
        else: rank = 'D (調整)'

        jobs.append({
            '優先度': rank, '機械': mach, 'MCS#': mcs, 
            '出荷日': ddate.strftime('%Y-%m-%d') if ddate else '2099-12-31',
            '数量': int(order_qty), '材料': mat_status
        })

    df_jobs = pd.DataFrame(jobs)
    if not df_jobs.empty:
        df_jobs.sort_values(by=['優先度'], inplace=True)
        df_jobs.insert(0, '実行順', range(1, len(df_jobs) + 1))
    return df_jobs, machine_params

# --- フロントエンドUI ---
with st.sidebar:
    st.header("⚙️ 計画設定")
    target_date = st.date_input("📅 計画日", datetime(2026, 4, 2))
    st.markdown("---")
    st.header("📂 データアップロード")
    f_master = st.file_uploader("1. MasterCard", type=['csv', 'xlsx'])
    f_delivery = st.file_uploader("2. Delivery", type=['csv', 'xlsx'])
    f_setup = st.file_uploader("3. Setup Speed", type=['csv', 'xlsx'])
    f_track = st.file_uploader("4. Floor Track (実績)", type=['csv', 'xlsx'])
    f_recv = st.file_uploader("5. Receiving Schedule (受入)", type=['csv', 'xlsx'])

if f_master and f_delivery and f_setup:
    raw_df, m_params = process_data(f_master, f_delivery, f_setup, f_track, f_recv, datetime.combine(target_date, datetime.min.time()))
    
    if raw_df.empty:
        st.warning("計画対象のジョブが見つかりませんでした。")
        st.stop()

    # 全機械の負荷グラフ
    st.markdown("### 📊 全機械 負荷状況")
    machine_list = sorted(raw_df['機械'].unique().tolist())
    workload_data = []
    for m in machine_list:
        m_df = raw_df[raw_df['機械'] == m]
        # 簡易計算（段取り25分+速度100枚/hで暫定計算）
        total_mins = len(m_df) * 25 + (m_df['数量'].sum() / 100 * 60)
        workload_data.append({'機械': m, '稼働予定(分)': round(total_mins, 1)})
    
    chart = alt.Chart(pd.DataFrame(workload_data)).mark_bar().encode(
        x='稼働予定(分):Q', y=alt.Y('機械:N', sort='-x'),
        color=alt.condition(alt.datum['稼働予定(分)'] > 480, alt.value('#e74c3c'), alt.value('#3498db'))
    ).properties(height=200)
    st.altair_chart(chart, use_container_width=True)

    # 機械ごとの開始時刻設定
    cols = st.columns(min(len(machine_list), 6))
    start_times = {}
    for i, m in enumerate(machine_list):
        with cols[i % len(cols)]:
            start_times[m] = st.time_input(f"{m} 開始", value=datetime.strptime("08:00", "%H:%M").time(), key=f"t_{m}")

    # 調整ボード
    selected_machine = st.selectbox("🎯 調整する機械", ["すべて"] + machine_list)
    display_df = raw_df if selected_machine == "すべて" else raw_df[raw_df['機械'] == selected_machine]

    edited_df = st.data_editor(
        display_df,
        column_config={
            "実行順": st.column_config.NumberColumn("順序", step=1),
            "数量": st.column_config.NumberColumn("数量"),
            "材料": st.column_config.TextColumn("材料状況")
        },
        hide_index=True, use_container_width=True
    )

    # 再計算
    edited_df.sort_values(by='実行順', inplace=True)
    current_times = {m: datetime.combine(target_date, start_times[m]) for m in machine_list}
    prev_mcs = {m: None for m in machine_list}
    final_recs = []

    for _, row in edited_df.iterrows():
        m = row['機械']
        # 段取り時間の決定（同一アイテムなら0分、それ以外は25分想定）
        setup = 0 if prev_mcs[m] == row['MCS#'] else 25
        duration = setup + (row['数量'] / 100 * 60)
        c_start = current_times[m]
        c_end = add_working_time(c_start, duration)
        
        row['開始'], row['終了'] = c_start.strftime("%H:%M"), c_end.strftime("%H:%M")
        
        # === 修正箇所：エラーが起きていた状況判定ロジックをきれいに整理 ===
        status = "✅ OK"
        
        # 1. 納期遅延チェック（繰越分や日付未定のものは除外）
        if row['出荷日'] not in ['繰越分', '2099-12-31'] and '-' in row['出荷日']:
            due_dt = datetime.strptime(row['出荷日'], "%Y-%m-%d").replace(hour=23, minute=59)
            if c_end > due_dt:
                status = "🚨 遅延"
                
        # 2. 材料ステータスの上書き（材料がなければそもそも生産できないため）
        if row['材料'] == "未入荷": 
            status = "❌ 材料待ち"
            
        row['状況'] = status
        # =========================================================
        
        final_recs.append(row)
        current_times[m] = add_working_time(c_end, 30)
        prev_mcs[m] = row['MCS#']

    st.dataframe(pd.DataFrame(final_recs)[['実行順', '優先度', '機械', 'MCS#', '数量', '材料', '開始', '終了', '状況']], use_container_width=True)

else:
    st.info("👈 左側のサイドバーから必要なファイルをアップロードしてください。")
