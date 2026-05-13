import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import altair as alt

# --- ページ設定 ---
st.set_page_config(page_title="生産計画調整ダッシュボード", layout="wide")
st.title("🏭 生産計画調整ダッシュボード (社内工程専用)")

# --- 時間計算ロジック（休憩除外） ---
def add_working_time(start_dt, duration_mins):
    if pd.isna(duration_mins) or duration_mins == float('inf'):
        duration_mins = 0
        
    if start_dt.year > 2100:
        start_dt = start_dt.replace(year=2100)
        
    current_time = start_dt
    remaining = float(duration_mins)
    
    if remaining > 500000: remaining = 500000 
    
    while remaining > 0:
        if current_time.year > 2100:
            break
            
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
        if file_obj.name.lower().endswith('.csv'): 
            return pd.read_csv(file_obj, skiprows=skip_rows, low_memory=False)
        else: 
            return pd.read_excel(file_obj, skiprows=skip_rows)

    def find_col(df, keywords, default_idx=None):
        for i, col in enumerate(df.columns):
            norm = str(col).upper().replace(' ', '').replace('＃', '#')
            for k in keywords:
                if k in norm: return col
        if default_idx is not None and len(df.columns) > default_idx:
            return df.columns[default_idx]
        return None

    # 各種ファイルの読み込み
    master_df = load_file(master_file, skip_rows=3)
    delivery_df = load_file(delivery_file, skip_rows=3)
    setup_speed_df = load_file(setup_file, skip_rows=0)
    
    for df in [master_df, delivery_df, setup_speed_df]:
        df.columns = df.columns.astype(str).str.strip()

    mcs_col_m = find_col(master_df, ['MCS#', 'MSC#'], default_idx=5)
    mcs_col_d = find_col(delivery_df, ['MCS#', 'MSC#'], default_idx=3)

    if not mcs_col_m or not mcs_col_d:
        st.error(f"MCS#列が特定できませんでした。")
        st.stop()

    # 受入(Receiving)データの構築
    recv_schedule = {}
    if recv_file:
        recv_df = load_file(recv_file, skip_rows=4)
        recv_df.columns = recv_df.columns.astype(str).str.strip()
        r_mcs = find_col(recv_df, ['MCS#', 'MSC#', 'PART'], default_idx=7)
        r_date = find_col(recv_df, ['RECV', 'DATE', 'STATUS'], default_idx=1)
        if r_mcs and r_date:
            for _, r in recv_df.iterrows():
                m_val = str(r[r_mcs]).strip()
                try:
                    dt = pd.to_datetime(r[r_date], errors='coerce')
                    if pd.notna(dt) and dt.year < 2100:
                        recv_schedule[m_val] = dt
                    else:
                        recv_schedule[m_val] = None
                except:
                    recv_schedule[m_val] = None

    mc_dict = master_df.drop_duplicates(subset=[mcs_col_m]).set_index(mcs_col_m).to_dict('index')

    def get_internal_routing(mcs):
        mcs_s = str(mcs).strip()
        if mcs_s not in mc_dict: return None
        mc_info = mc_dict[mcs_s]
        routing = [str(mc_info.get(f'MSP{i}', '')).strip() for i in range(1, 13)]
        internal = [m for m in routing if m not in ['CORR', 'nan', '', 'None', 'FALSE']]
        return internal[0] if internal else None

    jobs = []
    
    # 1. 実績・繰越
    if track_file:
        track_df = load_file(track_file, skip_rows=4)
        track_df.columns = track_df.columns.astype(str).str.strip()
        t_mcs = find_col(track_df, ['MCS#', 'MSC#'], default_idx=7)
        if t_mcs:
            for _, t_row in track_df.iterrows():
                mcs = str(t_row.get(t_mcs, '')).strip()
                plan = pd.to_numeric(t_row.get('PLAN OUT', 0), errors='coerce')
                good = pd.to_numeric(t_row.get('GOOD', 0), errors='coerce')
                plan = plan if pd.notna(plan) else 0
                good = good if pd.notna(good) else 0
                rem = plan - good
                
                if rem > 0:
                    mach = get_internal_routing(mcs)
                    if mach:
                        jobs.append({'優先度': 'A (繰越)', '機械': mach, 'MCS#': mcs, '出荷日': '繰越分', '数量': int(rem), '入荷予定': '入荷済', 'recv_dt': None})

    # 2. 新規Delivery (P列 = Index 15)
    qty_col = find_col(delivery_df, ['ORDER'], default_idx=15)
    due_col = find_col(delivery_df, ['DUE DATE', 'DELIVERY'])
    
    for _, d_row in delivery_df.dropna(subset=[mcs_col_d]).iterrows():
        mcs = str(d_row[mcs_col_d]).strip()
        qty_raw = d_row.get(qty_col, 0)
        qty = pd.to_numeric(qty_raw, errors='coerce')
        if pd.isna(qty) or qty <= 0:
            continue
        
        mach = get_internal_routing(mcs)
        if not mach: continue
        
        recv_dt = recv_schedule.get(mcs, None)
        recv_str = recv_dt.strftime('%m/%d') if pd.notna(recv_dt) and recv_dt.year < 2100 else "確認中"
        
        due_val = str(d_row.get(due_col, '-')) if due_col else '-'
        if due_val == 'nan' or pd.isna(due_val): due_val = '-'
        
        # 納期に基づく優先度の計算（当日・翌日必達をBとする）
        rank = 'D (調整)'
        try:
            if due_col and pd.notna(d_row[due_col]):
                # 日付をパースして計画日との差分（slack）を計算
                ddate = pd.to_datetime(d_row[due_col], dayfirst=True, errors='coerce')
                if pd.notna(ddate):
                    slack = (ddate.date() - target_date).days
                    if slack <= 1: rank = 'B (本日/翌日必達)'
                    elif slack <= 3: rank = 'C (推奨)'
                    due_val = ddate.strftime('%Y-%m-%d')
        except:
            pass
        
        jobs.append({
            '優先度': rank, '機械': mach, 'MCS#': mcs, 
            '出荷日': due_val, '数量': int(qty), '入荷予定': recv_str, 'recv_dt': recv_dt
        })

    if not jobs: return pd.DataFrame()

    df_jobs = pd.DataFrame(jobs)
    df_jobs.sort_values(by=['優先度'], inplace=True)
    df_jobs.insert(0, '実行順', range(1, len(df_jobs) + 1))
    return df_jobs

# --- UI部 ---
with st.sidebar:
    st.header("📂 データ読み込み")
    target_date = st.date_input("📅 計画日", datetime(2026, 4, 2))
    st.markdown("---")
    f_master = st.file_uploader("1. MasterCard", type=['csv', 'xlsx'])
    f_delivery = st.file_uploader("2. Delivery", type=['csv', 'xlsx'])
    f_setup = st.file_uploader("3. Setup Speed", type=['csv', 'xlsx'])
    f_track = st.file_uploader("4. Floor Track (実績)", type=['csv', 'xlsx'])
    f_recv = st.file_uploader("5. Receiving Schedule", type=['csv', 'xlsx'])

if f_master and f_delivery and f_setup:
    raw_df = process_data(f_master, f_delivery, f_setup, f_track, f_recv, target_date)

    if raw_df.empty:
        st.warning("⚠️ ジョブが見つかりませんでした。")
        st.stop()

    machine_list = sorted(raw_df['機械'].dropna().unique().tolist())
    
    # グラフの計算対象を「ランクAとB」に絞る
    st.markdown("### 📊 本日・翌日必達分の負荷状況 (単位: 時間)")
    st.caption("※前日からの繰越分（ランクA）と、本日・翌日納期分（ランクB）のみを合算してグラフ化しています。遠い将来の受注分は含まれません。")
    
    workload_data = []
    # AとBの優先度を持つジョブだけを抽出
    target_df = raw_df[raw_df['優先度'].str.contains('A|B', na=False)]
    
    for m in machine_list:
        m_df = target_df[target_df['機械'] == m]
        total_mins = len(m_df) * 25 + (m_df['数量'].sum() / 100 * 60)
        total_hours = total_mins / 60.0  # 分を時間に変換
        workload_data.append({'機械': m, '稼働予定(時間)': round(total_hours, 1)})
        
    chart = alt.Chart(pd.DataFrame(workload_data)).mark_bar().encode(
        x='稼働予定(時間):Q', 
        y=alt.Y('機械:N', sort='-x'), 
        color=alt.condition(alt.datum['稼働予定(時間)'] > 8, alt.value('#e74c3c'), alt.value('#3498db')) # 8時間を超えると赤色
    ).properties(height=200)
    st.altair_chart(chart, use_container_width=True)

    st.markdown("---")
    st.markdown("### ⏱ 機械ごとの開始時刻設定")
    cols = st.columns(min(len(machine_list), 6))
    start_times = {m: cols[i % len(cols)].time_input(f"{m} 開始", value=datetime.strptime("08:00", "%H:%M").time(), key=f"t_{m}") for i, m in enumerate(machine_list)}
    
    selected_machine = st.selectbox("🎯 機械フィルタ", ["すべて"] + machine_list)
    display_df = raw_df if selected_machine == "すべて" else raw_df[raw_df['機械'] == selected_machine]

    edited_df = st.data_editor(display_df.drop(columns=['recv_dt'], errors='ignore'), hide_index=True, use_container_width=True)
    edited_df.sort_values(by='実行順', inplace=True)

    # スケジュール再計算
    current_times = {m: datetime.combine(target_date, start_times[m]) for m in machine_list}
    final_recs = []
    for _, row in edited_df.iterrows():
        m = row['機械']
        orig = raw_df[raw_df['MCS#'] == row['MCS#']].iloc[0]
        rdt = orig['recv_dt']
        
        safe_rdt = rdt if pd.notna(rdt) and rdt.year < 2100 else None
        start = max(current_times[m], safe_rdt.replace(hour=8, minute=0) if safe_rdt else current_times[m])
        
        duration = 25 + (row['数量'] / 100 * 60)
        end = add_working_time(start, duration)
        
        row['開始'], row['終了'] = start.strftime("%H:%M"), end.strftime("%H:%M")
        row['状況'] = "✅ OK"
        if safe_rdt and safe_rdt > datetime.combine(target_date, datetime.min.time()): 
            row['状況'] = f"⏳ {row['入荷予定']} 入荷待ち"
            
        final_recs.append(row)
        current_times[m] = add_working_time(end, 30)

    st.dataframe(pd.DataFrame(final_recs)[['実行順', '優先度', '機械', 'MCS#', '数量', '入荷予定', '開始', '終了', '状況']], use_container_width=True)
    st.download_button("📥 確定したスケジュールをダウンロード", data=pd.DataFrame(final_recs).to_csv(index=False, encoding='utf-8-sig'), file_name=f'plan_{target_date.strftime("%Y%m%d")}.csv', mime='text/csv')
else:
    st.info("👈 左側のサイドバーから必要なファイルをアップロードしてください。")
