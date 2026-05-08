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
def process_data(master_file, delivery_file, setup_file, track_file, target_date):
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

    mc_dict = master_df.drop_duplicates(subset=['MCS#']).set_index('MCS#').to_dict('index')

    def get_routing(mcs):
        if mcs not in mc_dict: return None
        mc_info = mc_dict[mcs]
        routing = [str(mc_info[f'MSP{i}']).strip() for i in range(1, 13) if f'MSP{i}' in mc_info and pd.notna(mc_info[f'MSP{i}'])]
        routing = [m for m in routing if m != 'CORR' and m != '']
        return routing[0] if routing else None

    def get_setup_speed(mach, mcs):
        if mcs not in mc_dict: return 10.0, 100.0
        is_diecut = 1 if 'DC' in str(mc_dict[mcs].get('PD', '')) else 0
        params = machine_params.get(mach, {})
        setup_time, speed = 10.0, 100.0
        if mach.startswith('P') and is_diecut == 1 and '木型あり' in params:
            setup_time, speed = params['木型あり']['setup'], params['木型あり']['speed']
        return setup_time, speed

    jobs = []
    
    # 1. 実績ファイル(Floor Track)から「未完了分（キャリーオーバー）」を抽出
    if track_file is not None:
        track_df = load_file(track_file, skip_rows=4)
        track_df.columns = track_df.columns.str.strip()
        for _, t_row in track_df.iterrows():
            mcs = str(t_row.get('MCS#', '')).strip()
            if not mcs or mcs == 'nan': continue
            
            plan_out = float(t_row.get('PLAN OUT', 0)) if pd.notna(t_row.get('PLAN OUT')) else 0
            good = float(t_row.get('GOOD', 0)) if pd.notna(t_row.get('GOOD')) else 0
            remaining = int(plan_out - good)
            
            if remaining > 0:
                mach = get_routing(mcs)
                if not mach: continue
                setup_time, speed = get_setup_speed(mach, mcs)
                jobs.append({
                    '優先度': 'A (前日繰越)', '機械': mach, 'MCS#': mcs, 
                    '出荷日': '繰越分', '数量': remaining, '段取り(分)': setup_time, '基準速度(枚/時)': speed
                })

    # 2. 新規Deliveryデータの抽出
    def parse_date(d_str):
        if pd.isna(d_str): return None
        try: return datetime.strptime(f"{d_str.split(' ')[0]}/{target_date.year}", "%d/%m/%Y")
        except: return None
    delivery_df['Delivery_Date'] = delivery_df['DUE DATE'].apply(parse_date) if 'DUE DATE' in delivery_df.columns else None

    for _, d_row in delivery_df.dropna(subset=['MCS#']).iterrows():
        mcs = str(d_row['MCS#']).strip()
        order_qty = float(d_row['ORDER']) if 'ORDER' in d_row and pd.notna(d_row['ORDER']) else 0
        if order_qty <= 0: continue
        
        mach = get_routing(mcs)
        if not mach: continue
        
        ddate = d_row['Delivery_Date']
        slack = (ddate - target_date).days if ddate else 99
        if slack <= 1: rank = 'B (本日/翌日必達)'
        elif slack <= 3: rank = 'C (推奨)'
        else: rank = 'D (調整)'

        setup_time, speed = get_setup_speed(mach, mcs)
        jobs.append({
            '優先度': rank, '機械': mach, 'MCS#': mcs, 
            '出荷日': ddate.strftime('%Y-%m-%d') if ddate else '2099-12-31',
            '数量': int(order_qty), '段取り(分)': setup_time, '基準速度(枚/時)': speed
        })

    df_jobs = pd.DataFrame(jobs)
    if not df_jobs.empty:
        df_jobs.sort_values(by=['優先度'], inplace=True)
        df_jobs.insert(0, '実行順', range(1, len(df_jobs) + 1))
    return df_jobs

# --- フロントエンドUI ---
with st.sidebar:
    st.header("⚙️ 計画シミュレーション設定")
    target_date = st.date_input("📅 計画を作成する日付", datetime(2026, 4, 2))
    
    st.markdown("---")
    st.header("📂 データアップロード")
    f_master = st.file_uploader("1. MasterCard", type=['csv', 'xlsx'])
    f_delivery = st.file_uploader("2. Delivery", type=['csv', 'xlsx'])
    f_setup = st.file_uploader("3. Setup Speed", type=['csv', 'xlsx'])
    st.caption("👇 以下のファイルをアップロードすると、未完了分が自動で最優先として組み込まれます。")
    f_track = st.file_uploader("4. Floor Track Status (実績)", type=['csv', 'xlsx'])

if f_master and f_delivery and f_setup:
    # データ処理
    raw_df = process_data(f_master, f_delivery, f_setup, f_track, datetime.combine(target_date, datetime.min.time()))
    if raw_df.empty:
        st.warning("計画対象のジョブが見つかりませんでした。")
        st.stop()

    machine_list = sorted(raw_df['機械'].unique().tolist())

    # --- 1. 全機械の作業量比較グラフ ---
    st.markdown("### 📊 本日の全機械 負荷状況 (予定総稼働時間)")
    workload_data = []
    for m in machine_list:
        m_df = raw_df[raw_df['機械'] == m]
        total_mins = m_df['段取り(分)'].sum() + (m_df['数量'] / m_df['基準速度(枚/時)'] * 60).sum()
        workload_data.append({'機械': m, '総稼働予定(分)': round(total_mins, 1)})
    
    wl_df = pd.DataFrame(workload_data)
    chart = alt.Chart(wl_df).mark_bar().encode(
        x='総稼働予定(分):Q',
        y=alt.Y('機械:N', sort='-x'),
        color=alt.condition(alt.datum['総稼働予定(分)'] > 480, alt.value('#e74c3c'), alt.value('#3498db')),
        tooltip=['機械', '総稼働予定(分)']
    ).properties(height=200)
    st.altair_chart(chart, use_container_width=True)

    # --- 2. 稼働スタート時間の手修正設定 ---
    st.markdown("### ⏱ 機械ごとの稼働開始時間設定")
    st.caption("担当者の掛け持ち状況等に合わせて、各機械の最初のスタート時間を修正してください。")
    cols = st.columns(min(len(machine_list), 6))
    start_times = {}
    for i, m in enumerate(machine_list):
        with cols[i % len(cols)]:
            start_times[m] = st.time_input(f"{m} 開始時刻", value=datetime.strptime("08:00", "%H:%M").time(), key=f"time_{m}")

    st.markdown("---")
    
    # --- 3. 対象機械の選択と編集 ---
    selected_machine = st.selectbox("🎯 調整する機械を選択してください", ["すべて"] + machine_list)
    display_df = raw_df.copy() if selected_machine == "すべて" else raw_df[raw_df['機械'] == selected_machine].copy()

    st.markdown(f"#### 📋 計画調整ボード (順序・数量を変更して再計算)")
    edited_df = st.data_editor(
        display_df,
        column_config={
            "実行順": st.column_config.NumberColumn("順序", step=1, width="small"),
            "数量": st.column_config.NumberColumn("数量", min_value=0, step=100),
            "段取り(分)": None, "基準速度(枚/時)": None
        },
        hide_index=True, use_container_width=True
    )

    # --- 4. リアルタイム再計算（同アイテム段取りカット含む） ---
    edited_df.sort_values(by='実行順', inplace=True)
    calculated_records = []
    
    current_times = {m: datetime.combine(target_date, start_times[m]) for m in machine_list}
    prev_mcs_map = {m: None for m in machine_list}

    for _, row in edited_df.iterrows():
        mach = row['機械']
        c_start = current_times[mach]
        
        # 同一アイテム連続時の段取りカット
        is_continuous = (prev_mcs_map[mach] == row['MCS#'])
        actual_setup = 0 if is_continuous else row['段取り(分)']
        duration = actual_setup + (row['数量'] / row['基準速度(枚/時)']) * 60
        
        c_end = add_working_time(c_start, duration)
        
        status = "🚨 納期遅延" if row['出荷日'] not in ['繰越分', '2099-12-31'] and c_end > datetime.strptime(row['出荷日'], "%Y-%m-%d").replace(hour=23, minute=59) else "✅ 正常"

        row['開始'] = c_start.strftime("%m/%d %H:%M")
        row['終了'] = c_end.strftime("%m/%d %H:%M")
        row['連続生産'] = "🔄 結合" if is_continuous else "-"
        row['ステータス'] = status
        
        calculated_records.append(row)
        
        current_times[mach] = add_working_time(c_end, 30) # 30分インターバル
        prev_mcs_map[mach] = row['MCS#']

    final_df = pd.DataFrame(calculated_records)

    st.markdown("#### 🕒 シミュレーション結果 (時系列スケジュール)")
    def highlight_errors(s): return ['background-color: #ffcccc' if s['ステータス'] == '🚨 納期遅延' else 'background-color: #e8f8f5' if s['連続生産'] == '🔄 結合' else '' for _ in s]
    st.dataframe(final_df[['実行順', '優先度', '機械', 'MCS#', '数量', '連続生産', '開始', '終了', 'ステータス']].style.apply(highlight_errors, axis=1), use_container_width=True)

    csv = final_df.to_csv(index=False, encoding='utf-8-sig')
    st.download_button("📥 確定したスケジュールをダウンロード", data=csv, file_name=f'plan_{target_date.strftime("%Y%m%d")}.csv', mime='text/csv', type="primary")

else:
    st.info("👈 サイドバーから必要なファイルをアップロードしてください。")
