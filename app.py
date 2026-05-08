import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import altair as alt

# --- ページ設定 ---
st.set_page_config(page_title="生産計画調整ダッシュボード", layout="wide")
st.title("🏭 生産計画調整ダッシュボード (社内工程専用)")

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
            normalized_col = str(col).upper().replace('＃', '#').replace(' ', '')
            for k in keywords:
                if k.replace(' ', '') in normalized_col: return str(col)
        return None

    # 各種ファイルの読み込み
    master_df = load_file(master_file, skip_rows=3)
    delivery_df = load_file(delivery_file, skip_rows=3)
    setup_speed_df = load_file(setup_file, skip_rows=0)
    for df in [master_df, delivery_df, setup_speed_df]: df.columns = df.columns.astype(str).str.strip()

    # 材料受入データの読み込みと正規化
    recv_schedule = {}
    if recv_file:
        recv_df = load_file(recv_file, skip_rows=4)
        recv_df.columns = recv_df.columns.astype(str).str.strip()
        m_col = find_col(recv_df, ['MCS#'])
        d_col = find_col(recv_df, ['RECV', 'DATE', 'STATUS'])
        if m_col and d_col:
            for _, r in recv_df.iterrows():
                mcs_val = str(r[m_col]).strip()
                # 入荷日の解析
                try: 
                    dt_val = pd.to_datetime(r[d_col])
                    recv_schedule[mcs_val] = dt_val
                except:
                    recv_schedule[mcs_val] = datetime(2099, 12, 31) # 不明な場合は遠い未来

    mc_dict = master_df.drop_duplicates(subset=[find_col(master_df, ['MCS#']) or 'MCS#']).set_index(find_col(master_df, ['MCS#']) or 'MCS#').to_dict('index')

    def get_internal_routing(mcs):
        if mcs not in mc_dict: return None
        mc_info = mc_dict[mcs]
        # CORR以外の自社工程を抽出
        routing = [str(mc_info[f'MSP{i}']).strip() for i in range(1, 13) if f'MSP{i}' in mc_info and pd.notna(mc_info[f'MSP{i}'])]
        internal = [m for m in routing if m not in ['CORR', 'nan', '', 'None']]
        return internal[0] if internal else None

    jobs = []
    
    # 1. 実績・繰越
    if track_file:
        track_df = load_file(track_file, skip_rows=4)
        track_df.columns = track_df.columns.astype(str).str.strip()
        m_col_t = find_col(track_df, ['MCS#'])
        for _, t_row in track_df.iterrows():
            mcs = str(t_row.get(m_col_t, '')).strip()
            rem = float(t_row.get('PLAN OUT', 0)) - float(t_row.get('GOOD', 0))
            if rem > 0:
                mach = get_internal_routing(mcs)
                if mach:
                    jobs.append({'優先度': 'A (繰越)', '機械': mach, 'MCS#': mcs, '出荷日': '繰越分', '数量': int(rem), '入荷予定': '入荷済'})

    # 2. 新規Delivery
    mcs_col_d = find_col(delivery_df, ['MCS#']) or delivery_df.columns[3]
    for _, d_row in delivery_df.dropna(subset=[mcs_col_d]).iterrows():
        mcs = str(d_row[mcs_col_d]).strip()
        qty = float(d_row.get(find_col(delivery_df, ['ORDER', 'QTY']) or 'ORDER', 0))
        if qty <= 0: continue
        mach = get_internal_routing(mcs)
        if not mach: continue
        
        recv_dt = recv_schedule.get(mcs, None)
        recv_str = recv_dt.strftime('%m/%d') if recv_dt and recv_dt < datetime(2099,1,1) else "確認中"
        
        jobs.append({
            '優先度': 'B (通常)', '機械': mach, 'MCS#': mcs, 
            '出荷日': str(d_row.get('DUE DATE', '-')), '数量': int(qty), '入荷予定': recv_str, 'recv_dt': recv_dt
        })

    df_jobs = pd.DataFrame(jobs)
    if not df_jobs.empty:
        df_jobs.sort_values(by=['優先度'], inplace=True)
        df_jobs.insert(0, '実行順', range(1, len(df_jobs) + 1))
    return df_jobs

# --- UI ---
with st.sidebar:
    st.header("📂 データ読み込み")
    f_master = st.file_uploader("1. MasterCard", type=['csv', 'xlsx'])
    f_delivery = st.file_uploader("2. Delivery", type=['csv', 'xlsx'])
    f_setup = st.file_uploader("3. Setup Speed", type=['csv', 'xlsx'])
    f_track = st.file_uploader("4. Floor Track (進捗実績)", type=['csv', 'xlsx'])
    f_recv = st.file_uploader("5. Receiving Schedule (CORR入荷)", type=['csv', 'xlsx'])

if f_master and f_delivery and f_setup:
    raw_df = process_data(f_master, f_delivery, f_setup, f_track, f_recv, datetime.now())
    
    st.markdown("### 📊 自社機械 負荷状況")
    machine_list = sorted(raw_df['機械'].unique().tolist())
    # グラフ表示... (中略: 前回と同様のAltairグラフ)
    
    selected_machine = st.selectbox("🎯 機械フィルタ", ["すべて"] + machine_list)
    display_df = raw_df if selected_machine == "すべて" else raw_df[raw_df['機械'] == selected_machine]

    edited_df = st.data_editor(display_df, hide_index=True, use_container_width=True)

    # 再計算
    edited_df.sort_values(by='実行順', inplace=True)
    final_recs = []
    current_times = {m: datetime.now().replace(hour=8, minute=0) for m in machine_list}
    
    for _, row in edited_df.iterrows():
        m = row['機械']
        # 材料入荷日による開始制約
        earliest_start = current_times[m]
        if 'recv_dt' in row and row['recv_dt'] and row['recv_dt'] > earliest_start:
            earliest_start = row['recv_dt'].replace(hour=8, minute=0)
            
        duration = 25 + (row['数量'] / 100 * 60)
        c_end = add_working_time(earliest_start, duration)
        
        row['開始'], row['終了'] = earliest_start.strftime("%H:%M"), c_end.strftime("%H:%M")
        row['状況'] = "✅ OK"
        if 'recv_dt' in row and row['recv_dt'] and row['recv_dt'] > datetime.now():
            row['状況'] = f"⏳ {row['入荷予定']} 入荷待ち"
        
        final_recs.append(row)
        current_times[m] = add_working_time(c_end, 30)

    st.dataframe(pd.DataFrame(final_recs)[['実行順', '優先度', '機械', 'MCS#', '数量', '入荷予定', '開始', '終了', '状況']], use_container_width=True)
