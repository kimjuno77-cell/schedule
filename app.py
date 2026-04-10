import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import re
from datetime import date, timedelta

# 페이지 설정
st.set_page_config(page_title="월간 진도 보고서 (Monthly Progress Report)", layout="wide")

st.title("📊 월간 진도 보고서 생성기")
st.markdown("---")

# 사이드바: 프로젝트 설정
with st.sidebar:
    st.header("📂 프로젝트 설정")
    project_name = st.text_input("프로젝트 이름 (Project Name)", value="New Project", key="project_name")
    
    # 프로젝트 시작일 및 납품일 설정
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        project_start_date = st.date_input("프로젝트 시작일", value=date.today(), key="project_start_date")
    with col_d2:
        contract_delivery_date = st.date_input("계약 납품일", value=date.today() + timedelta(weeks=40), key="contract_delivery_date")
        
    st.divider()
    uploaded_file = st.file_uploader("기존 엑셀 파일 불러오기", type=["xlsx", "csv"])
    st.info("💡 팀원 배포용: 이 프로그램을 폴더째로 공유하면 됩니다.")

# 기본 항목 리스트 및 제작 기간 정의
default_items_map = {
    "Ammonia Injection Grid": 10,
    "Catalyst": 54,
    "Dilution Tank": 12,
    "Ammonia Tank": 20,
    "Vaporizer": 12,
    "Ammonia_RA Supply Pump": 20,
    "Fan(Hot Gas_Cooling Air)": 24,
    "Ammonia Flow Control Unit Skid": 16,
    "Manual Valve": 16,
    "PRV & PSV": 16,
    "Ystrainer": 16,
    "Instrument": 24,
    "Control Valve & Damper": 24,
    "Analyzer": 36,
    "Catalyst Structure": 12,
    "Header & Manifold": 12,
    "Piping Spool": 20,
    "Pipe Support": 12,
    "Raw Material": 12,
    "Panel(MCC, LCP, E.H.T, Chemical)": 24,
    "Hook Up Material": 20,
    "Cable": 20,
    "Cable Tray": 16,
    "etc.": 16,
    "Lighting System": 16,
    "Atomizing Nozzle": 12,
    "Breather Valve": 24,
    "Manual Damper": 24,
    "Expansion Joint": 12,
    "Eye Shower": 8,
    "Insulation": 8
}
default_items = list(default_items_map.keys())

# 데이터 초기화
# 데이터 초기화 및 파일 로드 로직 개선
if 'data' not in st.session_state:
    st.session_state.data = None

# 파일값 변경 감지 (새 파일 업로드 시 데이터 갱신)
if uploaded_file is not None:
    # 기존에 로드한 파일과 다른지 확인 (또는 최초 로드)
    curr_file_id = uploaded_file.file_id if hasattr(uploaded_file, 'file_id') else uploaded_file.name
    
    if 'loaded_file_id' not in st.session_state or st.session_state.loaded_file_id != curr_file_id:
        try:
            if uploaded_file.name.endswith('.csv'):
                st.session_state.data = pd.read_csv(uploaded_file)
            else:
                # Load Data
                xl = pd.ExcelFile(uploaded_file)
                st.session_state.data = xl.parse(0) # Assume data is first sheet or 'Schedule'
                
                # Load Metadata if exists
                if 'ProjectInfo' in xl.sheet_names:
                    meta_df = xl.parse('ProjectInfo')
                    if not meta_df.empty:
                        # Expecting columns: Key, Value or single row with col headers
                        # Let's assume structure: Columns [ProjectName, StartDate, DeliveryDate]
                        try:
                            if 'ProjectName' in meta_df.columns:
                                st.session_state['project_name'] = str(meta_df.iloc[0]['ProjectName'])
                            if 'StartDate' in meta_df.columns:
                                st.session_state['project_start_date'] = pd.to_datetime(meta_df.iloc[0]['StartDate']).date()
                            if 'DeliveryDate' in meta_df.columns:
                                st.session_state['contract_delivery_date'] = pd.to_datetime(meta_df.iloc[0]['DeliveryDate']).date()
                            st.success(f"프로젝트 정보 복구 완료: {st.session_state.get('project_name')}")
                        except Exception as meta_ex:
                            st.warning(f"메타데이터 로드 중 일부 오류: {meta_ex}")

            st.session_state.loaded_file_id = curr_file_id
            st.success(f"파일이 성공적으로 로드되었습니다: {uploaded_file.name}")
            st.rerun() # Rerun to apply loaded session state to widgets
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

# 업로드된 파일이 없고 데이터도 없으면 기본 데이터 생성
if st.session_state.data is None:
        # 빈 데이터 프레임 생성
        data = {
            '항목 (Item)': default_items,
            '금액 (Amount)': [0] * len(default_items), # New Amount Column
            # 매핑된 제작 기간 적용, 없으면 0
            '제작 기간 (Weeks)': [default_items_map.get(item, 0) for item in default_items], 
            '가중치 (Weight)': [0] * len(default_items),
            '전월 계획 (Plan Prev)': [0] * len(default_items),
            '전월 실적 (Actual Prev)': [0] * len(default_items),
            '금월 계획 (Plan Curr)': [0] * len(default_items),
            '금월 실적 (Actual Curr)': [0] * len(default_items),
        }
        # 날짜 컬럼 추가
        cols = [
            '설계 계획 시작', '설계 계획 종료', '설계 실적 시작', '설계 진행률 (%)', '설계 실적 종료',
            '구매 계획 시작', '구매 계획 종료', '구매 실적 시작', '구매 진행률 (%)', '구매 실적 종료',
            '제작 계획 시작', '제작 계획 종료', '제작 실적 시작', '제작 진행률 (%)', '제작 실적 종료',
            '검사 계획 시작', '검사 계획 종료', '검사 실적 시작', '검사 진행률 (%)', '검사 실적 종료',
            '납품 계획 시작', '납품 계획 종료', '납품 실적 시작', '납품 진행률 (%)', '납품 실적 종료'
        ]
        for c in cols:
            data[c] = [None] * len(default_items)
            
        st.session_state.data = pd.DataFrame(data)

df = st.session_state.data

# 자동 스케줄링 로직
# 자동 스케줄링 로직
def auto_schedule(df, start_date):
    # Ensure start_date is a proper python date object
    if isinstance(start_date, pd.Timestamp):
        base_start = start_date.date()
    elif isinstance(start_date, date):
        base_start = start_date
    else:
        # Fallback for strings or other types
        base_start = pd.to_datetime(start_date).date()

    for i, row in df.iterrows():
        try:
            # 제작 기간 처리 (float conversion)
            raw_weeks = row.get('제작 기간 (Weeks)', 0)
            manuf_weeks = float(raw_weeks) if pd.notnull(raw_weeks) and raw_weeks != '' else 0
            
            # 1. 구매 (15일)
            p_start = base_start
            p_end = p_start + timedelta(days=15)
            df.at[i, '구매 계획 시작'] = p_start
            df.at[i, '구매 계획 종료'] = p_end

            # 2. 설계 (120일) - 구매 종료 + 3일 후 시작
            d_start = p_end + timedelta(days=3)
            d_end = d_start + timedelta(days=120)
            df.at[i, '설계 계획 시작'] = d_start
            df.at[i, '설계 계획 종료'] = d_end
            
            # 3. 제작 (기간 = 주 * 7일) - 설계 종료 + 1일 후 시작
            if manuf_weeks > 0:
                m_start = d_end + timedelta(days=1)
                m_end = m_start + timedelta(days=int(manuf_weeks * 7))
            else:
                m_start = None
                m_end = None
                
            df.at[i, '제작 계획 시작'] = m_start
            df.at[i, '제작 계획 종료'] = m_end
            
            # 4. 검사 (14일) - 제작 종료(없으면 설계 종료) + 1일 후 시작
            base_for_insp = m_end if (m_end is not None) else d_end
            
            i_start = base_for_insp + timedelta(days=1)
            i_end = i_start + timedelta(days=14)
            df.at[i, '검사 계획 시작'] = i_start
            df.at[i, '검사 계획 종료'] = i_end

            # 5. 납품 (7일) - 검사 종료 + 7일 후 시작
            del_start = i_end + timedelta(days=7)
            del_end = del_start + timedelta(days=7)
            df.at[i, '납품 계획 시작'] = del_start
            df.at[i, '납품 계획 종료'] = del_end
            
        except Exception as ex:
            st.error(f"Row {i} ('{row.get('항목 (Item)', 'Unknown')}') 처리 중 오류: {ex}")
            continue
            
    return df

# 상단 툴바
col_tool1, col_tool2 = st.columns([1, 4])
with col_tool1:
    if st.button("📅 일정 자동 계산 (Auto Plan)"):
        st.session_state.data = auto_schedule(st.session_state.data, project_start_date)
        st.success("일정이 자동 계산되었습니다! (구매 15일, 설계 120일 등 설정된 규칙 적용)")
        st.rerun()
with col_tool2:
    st.info("ℹ️ 항목별 '제작 기간 (Weeks)'이 기본값으로 설정되어 있습니다. 필요시 수정한 후 자동 계산 버튼을 누르세요.")

df = st.session_state.data

# 단계 정의 (순서 변경: 구매 -> 설계)
phases_info = [
    ('구매 (Procurement)', '구매 계획 시작', '구매 계획 종료', '구매 실적 시작', '구매 진행률 (%)', '구매 실적 종료'),
    ('설계 (Design)', '설계 계획 시작', '설계 계획 종료', '설계 실적 시작', '설계 진행률 (%)', '설계 실적 종료'),
    ('제작 (Manufacturing)', '제작 계획 시작', '제작 계획 종료', '제작 실적 시작', '제작 진행률 (%)', '제작 실적 종료'),
    ('검사 (Inspection)', '검사 계획 시작', '검사 계획 종료', '검사 실적 시작', '검사 진행률 (%)', '검사 실적 종료'),
    ('납품 (Delivery)', '납품 계획 시작', '납품 계획 종료', '납품 실적 시작', '납품 진행률 (%)', '납품 실적 종료'),
]
all_date_cols = []
all_prog_cols = []
for p in phases_info:
    all_date_cols.extend([p[1], p[2], p[3], p[5]])
    all_prog_cols.append(p[4])

# 날짜 형변환 및 컬럼 순서 재정렬 (구매 -> 설계 -> 제작...)
# phases_info의 순서대로 날짜 컬럼을 정렬한다.
ordered_columns = ['항목 (Item)', '금액 (Amount)', '제작 기간 (Weeks)', '가중치 (Weight)', '전월 계획 (Plan Prev)', '전월 실적 (Actual Prev)', '금월 계획 (Plan Curr)', '금월 실적 (Actual Curr)']
for p in phases_info:
    ordered_columns.extend([p[1], p[2], p[3], p[4], p[5]])

# 데이터프레임에 없는 컬럼이 있을 수 있으므로 교집합만 사용하거나 새로 생성
for col in ordered_columns:
    if col not in df.columns:
        df[col] = None 
        # 금액 컬럼 기본값 0 처리
        if col == '금액 (Amount)':
             df[col] = 0
        elif '진행률 (%)' in col:
             df[col] = 0

# 날짜 형변환
for col in all_date_cols:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

# 컬럼 순서 강제 적용
df = df[ordered_columns]

# 메인 입력 화면
st.subheader(f"📝 {project_name} - 상세 진도 및 일정 입력")

column_config = {
    "항목 (Item)": st.column_config.TextColumn(width="medium", disabled=False),
    "금액 (Amount)": st.column_config.NumberColumn(format="%d"),
    "제작 기간 (Weeks)": st.column_config.NumberColumn(format="%d주"),
    "가중치 (Weight)": st.column_config.NumberColumn(format="%.2f%%"), # 가중치는 자동 계산되지만 필요 시 수정 가능
    "전월 계획 (Plan Prev)": st.column_config.NumberColumn(format="%d%%"),
    "전월 실적 (Actual Prev)": st.column_config.NumberColumn(format="%d%%"),
    "금월 계획 (Plan Curr)": st.column_config.NumberColumn(format="%d%%"),
    "금월 실적 (Actual Curr)": st.column_config.NumberColumn(format="%d%%"),
}
for col in all_date_cols:
    column_config[col] = st.column_config.DateColumn(format="YYYY-MM-DD")
for col in all_prog_cols:
    column_config[col] = st.column_config.NumberColumn(format="%d%%", min_value=0, max_value=100)

with st.form("entry_form"):
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config=column_config,
        key="data_editor_v7" # Key updated to force refresh
    )
    
    submitted = st.form_submit_button("💾 입력 데이터 적용 (Apply Changes)")
    
    if submitted:
        st.session_state.data = edited_df
        st.success("데이터가 적용되었습니다. (Data Updated)")


# --- 계산 및 시각화 ---
import traceback

try:
    # Ensure columns are numeric (Handle string inputs like '50%' or '50')
    # Ensure columns are numeric (Handle string inputs like '50%' or '50')
    num_cols = ['가중치 (Weight)', '전월 계획 (Plan Prev)', '전월 실적 (Actual Prev)', '금월 계획 (Plan Curr)', '금월 실적 (Actual Curr)']
    
    # 0. Amount Cleaning
    if '금액 (Amount)' in edited_df.columns:
        edited_df['금액 (Amount)'] = edited_df['금액 (Amount)'].astype(str).str.replace(',', '').str.strip()
        edited_df['금액 (Amount)'] = pd.to_numeric(edited_df['금액 (Amount)'], errors='coerce').fillna(0)
        
        # 1. Clean Weight Column First (Always ensure it's numeric)
        w_col = '가중치 (Weight)'
        if w_col in edited_df.columns:
             edited_df[w_col] = edited_df[w_col].astype(str).str.replace('%', '').str.replace(',', '').str.strip()
             edited_df[w_col] = pd.to_numeric(edited_df[w_col], errors='coerce').fillna(0)

        # 2. Logic: Amount vs Weight
        total_amount = edited_df['금액 (Amount)'].sum()
        if total_amount > 0:
             # Case A: Amount exists -> Calculate Weight % based on Amount
             edited_df['가중치 (Weight)'] = (edited_df['금액 (Amount)'] / total_amount) * 100
        else:
             # Case B: Amount is 0 -> Use Manual Weight or Fallback
             if edited_df[w_col].sum() == 0 and len(edited_df) > 0:
                 # Fallback: All weights are 0 -> Apply Equal Weights
                 edited_df[w_col] = 100.0 / len(edited_df)
    
    # 3. Clean other numeric columns
    for c in num_cols:
        if c == '가중치 (Weight)': continue # Already cleaned/calculated
        
        # Convert column to string, strip '%', then to numeric
        if c in edited_df.columns:
            edited_df[c] = edited_df[c].astype(str).str.replace('%', '').str.replace(',', '').str.strip()
            edited_df[c] = pd.to_numeric(edited_df[c], errors='coerce').fillna(0)
            
    # --- Automatic Progress Calculation (New Request) ---
    # Global Phase Weights: Procurement 10, Design 20, Mfg 40, Insp 25, Delivery 5
    # Logic: Start=50%, End=100% of Phase Weight
    phase_ratios = {
        '구매 (Procurement)': 10.0,
        '설계 (Design)': 20.0,
        '제작 (Manufacturing)': 40.0,
        '검사 (Inspection)': 25.0,
        '납품 (Delivery)': 5.0
    }
    
    # Reference Dates
    now_ts = pd.Timestamp.now()
    first_day_of_month = now_ts.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    # Last day of current month (First day of next month - 1 day)
    next_month = (first_day_of_month + pd.DateOffset(months=1))
    last_day_of_month = next_month - pd.Timedelta(days=1)
    
    curr_actual_list = []
    prev_actual_list = []
    curr_plan_list = []
    prev_plan_list = []
    
    # Iterate over rows to calculate progress per Item
    for index, row in edited_df.iterrows():
        progress_accum_curr_act = 0.0
        progress_accum_prev_act = 0.0
        progress_accum_curr_plan = 0.0
        progress_accum_prev_plan = 0.0
        
        # Calculate progress based on phase status
        for phase_name, p_s, p_e, a_s, a_prog, a_e in phases_info:
            weight = phase_ratios.get(phase_name, 0)
            
            # --- Actual Calculation ---
            start_val = row[a_s]
            end_val = row[a_e]
            
            prog = 0
            try:
                prog_val = row.get(a_prog, 0)
                if pd.notnull(prog_val):
                    if isinstance(prog_val, str):
                        prog_val = float(prog_val.replace('%', '').strip())
                    prog = float(prog_val) / 100.0
            except:
                prog = 0
                
            if pd.notnull(end_val):
                prog = 1.0
            
            # Current Actual
            if prog > 0:
                progress_accum_curr_act += weight * prog
            elif pd.notnull(start_val): # In Progress (Fallback if prog is 0)
                progress_accum_curr_act += weight * 0.5
                
            # Previous Actual
            try:
                s_date = pd.to_datetime(start_val) if pd.notnull(start_val) else None
                e_date = pd.to_datetime(end_val) if pd.notnull(end_val) else None
                
                if e_date and e_date < first_day_of_month:
                    progress_accum_prev_act += weight * 1.0
                elif s_date and s_date < first_day_of_month:
                    progress_accum_prev_act += weight * 0.5
            except: pass

            # --- Plan Calculation ---
            p_start_val = row[p_s]
            p_end_val = row[p_e]
            
            try:
                ps_date = pd.to_datetime(p_start_val) if pd.notnull(p_start_val) else None
                pe_date = pd.to_datetime(p_end_val) if pd.notnull(p_end_val) else None
                
                # Previous Plan (Scheduled before this month)
                if pe_date and pe_date < first_day_of_month:
                    progress_accum_prev_plan += weight * 1.0
                elif ps_date and ps_date < first_day_of_month:
                    progress_accum_prev_plan += weight * 0.5
                    
                # Current Plan (Scheduled up to end of this month)
                # Note: This is cumulative target
                if pe_date and pe_date <= last_day_of_month:
                    progress_accum_curr_plan += weight * 1.0
                elif ps_date and ps_date <= last_day_of_month:
                    progress_accum_curr_plan += weight * 0.5
            except: pass
        
        curr_actual_list.append(progress_accum_curr_act)
        prev_actual_list.append(progress_accum_prev_act)
        curr_plan_list.append(progress_accum_curr_plan)
        prev_plan_list.append(progress_accum_prev_plan)
        
    # Apply calculated progress
    edited_df['금월 실적 (Actual Curr)'] = curr_actual_list
    edited_df['전월 실적 (Actual Prev)'] = prev_actual_list
    edited_df['금월 계획 (Plan Curr)'] = curr_plan_list
    edited_df['전월 계획 (Plan Prev)'] = prev_plan_list
    # ----------------------------------------------------

    edited_df['월간 진도 (Monthly Progress)'] = edited_df['금월 실적 (Actual Curr)'] - edited_df['전월 실적 (Actual Prev)']
    
    total_weight = edited_df['가중치 (Weight)'].sum()
    if total_weight > 0:
        overall_plan = (edited_df['금월 계획 (Plan Curr)'] * edited_df['가중치 (Weight)']).sum() / total_weight
        overall_actual = (edited_df['금월 실적 (Actual Curr)'] * edited_df['가중치 (Weight)']).sum() / total_weight
    else:
        overall_plan = 0; overall_actual = 0
        
    status_msg = "정상 (On Track)"
    if overall_actual < overall_plan: status_msg = "지연 (Delayed)"
    elif overall_actual > overall_plan: status_msg = "초과 달성 (Ahead)"

    # 2. 지연 분석
    delay_alerts = []
    
    # Contract delivery date (ensure date object)
    if isinstance(contract_delivery_date, pd.Timestamp):
        contract_delivery_date_obj = contract_delivery_date.date()
    elif isinstance(contract_delivery_date, date):
        contract_delivery_date_obj = contract_delivery_date
    else:
        contract_delivery_date_obj = pd.to_datetime(contract_delivery_date).date()
    
    for index, row in edited_df.iterrows():
        item_name = row['항목 (Item)']
        
        # 단계별 지연
        for phase_name, p_start, p_end, a_start, a_prog, a_end in phases_info:
            plan_end = row[p_end]
            actual_end = row[a_end]
            
            # Check if values are valid dates (not NaT or None)
            if pd.notnull(plan_end) and pd.notnull(actual_end):
                try:
                    p_e = pd.to_datetime(plan_end).date()
                    a_e = pd.to_datetime(actual_end).date()
                    
                    if a_e > p_e:
                        days_diff = (a_e - p_e).days
                        delay_alerts.append(f"⚠️ **{item_name}** - {phase_name}: {days_diff}일 지연 (계획: {p_e}, 실적: {a_e})")
                except: continue
        
        # 납품일 체크 (Smart Alert Logic)
        last_plan_end = row['납품 계획 종료']
        if pd.notnull(last_plan_end):
             try:
                 l_p_e = pd.to_datetime(last_plan_end).date()
                 if l_p_e > contract_delivery_date_obj:
                     delay_days = (l_p_e - contract_delivery_date_obj).days
                     alert_msg = f"🚨 **{item_name}** - 계약 납품일({contract_delivery_date}) {delay_days}일 초과! (계획: {l_p_e})"
                     
                     # Smart Compression Logic
                     solutions = []
                     remaining_delay = delay_days
                     
                     # 1. Compress Design Phase
                     d_start_val = row.get('설계 계획 시작')
                     d_end_val = row.get('설계 계획 종료')
                     
                     if pd.notnull(d_start_val) and pd.notnull(d_end_val):
                         ds = pd.to_datetime(d_start_val).date()
                         de = pd.to_datetime(d_end_val).date()
                         if de > ds:
                             curr_d_days = (de - ds).days
                             # Minimum 30 days (approx 1 month)
                             reduceable_d = max(0, curr_d_days - 30)
                             
                             if reduceable_d > 0:
                                 reduce_amount = min(remaining_delay, reduceable_d)
                                 solutions.append(f"설계 기간 {reduce_amount}일 단축 (현재 {curr_d_days}일 -> 권장 {curr_d_days - reduce_amount}일)")
                                 remaining_delay -= reduce_amount
                     
                     # 2. Compress Manufacturing Phase (If delay remains)
                     if remaining_delay > 0:
                         solutions.append(f"제작 기간 {remaining_delay}일 단축 필요")
                         
                     if solutions:
                         alert_msg += " 👉 [제안] " + ", ".join(solutions)
                         
                     delay_alerts.append(alert_msg)
             except: continue

    st.markdown("---")
    st.subheader(f"📊 {project_name} 종합 리포트")
    c1, c2, c3 = st.columns(3)
    c1.metric("전체 계획 공정률", f"{overall_plan:.2f}%")
    c2.metric("전체 실적 공정률", f"{overall_actual:.2f}%", delta=f"{overall_actual - overall_plan:.2f}%")
    c3.metric("종합 상태", status_msg)
    
    if delay_alerts:
        st.error("🚨 **주요 이슈 및 지연 알림**")
        for alert in delay_alerts:
            st.write(alert)

    # 3. 상세 진도율 테이블 표시 (UI에 표시)
    st.markdown("---")
    st.subheader("📋 상세 진도율 검토 (Detailed Progress Review)")
    
    # 표시할 컬럼 정의
    review_cols = ['항목 (Item)', '가중치 (Weight)', '전월 실적 (Actual Prev)', '금월 실적 (Actual Curr)', '월간 진도 (Monthly Progress)']
    # 존재하는 컬럼만 필터링
    final_review_cols = [c for c in review_cols if c in edited_df.columns]
    
    # 데이터프레임 표시 (포맷팅 적용)
    st.dataframe(
        edited_df[final_review_cols].style.format({
            '가중치 (Weight)': '{:.2f}%',
            '전월 실적 (Actual Prev)': '{:.2f}%',
            '금월 실적 (Actual Curr)': '{:.2f}%',
            '월간 진도 (Monthly Progress)': '{:.2f}%'
        }),
        use_container_width=True,
        hide_index=True
    )

    # --- Chart Generation Functions ---
    def create_gantt_chart(df, phases, title):
        # Prepare data for Plotly Gantt
        plan_data = [] # For bar chart (px.timeline)
        
        # Color mapping (Pastel Tones for Plan)
        phase_colors = {
            '구매 (Procurement)': '#A0C4FF',       # Pastel Blue
            '설계 (Design)': '#9BF6FF',            # Pastel Cyan
            '제작 (Manufacturing)': '#FFADAD',     # Pastel Red
            '검사 (Inspection)': '#FFD6A5',        # Pastel Orange
            '납품 (Delivery)': '#CAFFBF'           # Pastel Green
        }
        
        # Lists for the Vertical Progress Line (Actual)
        line_dates = []
        line_items = []
        
        # We need to iterate in the order they appear in the DataFrame to maintain vertical connection
        # Ensure df is sorted if needed, but usually it's in logic order. 
        # If we sort for Plotly Y-axis (which is reversed usually), we should match that order.
        # Plotly draws Y axis from bottom up by default, but we use 'reversed' in update_yaxes.
        # So top row in DF = Top row in Chart.
        
        for index, row in df.iterrows():
            item_name = row['항목 (Item)']
            if pd.isna(item_name) or str(item_name).strip() == "": continue
            
            # 1. Collect Plan Data
            item_has_plan = False
            for phase_name, p_start, p_end, a_start, a_prog, a_end in phases:
                if pd.notnull(row[p_start]) and pd.notnull(row[p_end]):
                    plan_data.append(dict(
                        Item=item_name, 
                        Y_Label=item_name,  
                        Phase=phase_name, 
                        Start=row[p_start], 
                        Finish=row[p_end],
                        Type="Plan"
                    ))
                    item_has_plan = True
            
            # 2. Find Latest Status Date (Earned Schedule) for this Item
            # Strategy: If Actual End exists -> Plot at Plan End.
            #           If Actual Start exists -> Plot at Plan Start.
            #           This visualizes "How much planned work has been achieved".
            #           Right of Today = Ahead (Completed future work).
            #           Left of Today  = Delay (Only completed past work).
            
            valid_plan_dates = []
            
            for phase_name, p_start, p_end, a_start, a_prog, a_end in phases:
                if pd.notnull(row[p_start]) and pd.notnull(row[p_end]):
                    ps_dt = pd.to_datetime(row[p_start])
                    pe_dt = pd.to_datetime(row[p_end])
                    total_duration = (pe_dt - ps_dt).days
                    
                    prog = 0
                    try:
                        prog_val = row.get(a_prog, 0)
                        if pd.notnull(prog_val):
                            if isinstance(prog_val, str):
                                prog_val = float(prog_val.replace('%', '').strip())
                            prog = float(prog_val) / 100.0
                    except:
                        prog = 0
                        
                    if pd.notnull(row[a_end]):
                        prog = 1.0
                        
                    if prog > 0:
                        earned_days = total_duration * prog
                        earned_date = ps_dt + pd.Timedelta(days=earned_days)
                        valid_plan_dates.append(earned_date.date())
                    elif pd.notnull(row[a_start]):
                        valid_plan_dates.append(ps_dt.date())
            
            if valid_plan_dates:
                # Take the latest Plan Date achieved
                line_dates.append(max(valid_plan_dates))
                line_items.append(item_name)
            else:
                pass

        if not plan_data and not line_dates:
            return None
        
        # --- Create Figure ---
        fig = go.Figure()
        
        # 1. Add Plan Bars
        if plan_data:
            g_df = pd.DataFrame(plan_data)
            g_df['Start'] = pd.to_datetime(g_df['Start'])
            g_df['Finish'] = pd.to_datetime(g_df['Finish'])
            
            # Important: We want the Y-axis order to follow df order.
            # Plotly maps categorical Y based on appearance or sort.
            # We can force the category order.
            
            # Instead of separate px.timeline, let's add traces to go.Figure manually or use px and add scatter.
            # Using px.timeline is easier for the bars.
            fig = px.timeline(
                g_df, x_start="Start", x_end="Finish", y="Y_Label", color="Phase", 
                color_discrete_map=phase_colors,
                opacity=0.5, # Background opacity
                hover_data=["Item", "Phase", "Start", "Finish"], 
                title=title
            )
        else:
            fig = go.Figure()
            fig.update_layout(title=title)

        # 2. Add Vertical Progress Line
        if line_dates:
            fig.add_trace(go.Scatter(
                x=line_dates,
                y=line_items,
                mode='lines+markers',
                name='Actual Progress (Original)',
                marker=dict(symbol='circle', size=10, color='#FF5733'), # Red-Orange dot
                line=dict(color='#FF5733', width=3), # Connection line
                hoverinfo='x+y+text',
                hovertext=[f"Latest: {d}" for d in line_dates]
            ))

        # --- Layout Adjustments ---
        fig.update_yaxes(
            autorange="reversed", # Start from top
            title_text="항목 (Item)",
            type='category', # Ensure categorical
            categoryorder='array', # Force order
            categoryarray=df['항목 (Item)'].tolist(), # Use exact DF order
            showgrid=True,
            gridwidth=1,
            gridcolor='#888888',
        )
        
        # Add 'Today' Line
        today_ts = pd.Timestamp.now().timestamp() * 1000
        fig.add_vline(x=today_ts, line_width=2, line_dash="solid", line_color="red", annotation_text="Today")
        
        fig.update_xaxes(
            type='date', 
            showgrid=True, 
            gridwidth=0.5, 
            gridcolor='#E0E0E0',
            dtick=864000000.0, # 10 Days
            tickformat="%m-%d"
        )
        
        fig.update_layout(
            height=max(600, len(df) * 40), 
            template='plotly_white',
            barmode='overlay',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        
        return fig

    def create_plan_vs_actual_gantt(df, phases):
        # Prepare data for Plan vs Actual Gantt
        gantt_data = []
        
        # Color mapping for phases (Synced with Main Gantt)
        phase_colors = {
            '구매 (Procurement)': '#A0C4FF',       # Pastel Blue
            '설계 (Design)': '#9BF6FF',            # Pastel Cyan
            '제작 (Manufacturing)': '#FFADAD',     # Pastel Red
            '검사 (Inspection)': '#FFD6A5',        # Pastel Orange
            '납품 (Delivery)': '#CAFFBF'           # Pastel Green
        }
        
        for index, row in df.iterrows():
            item_name = row['항목 (Item)']
            if pd.isna(item_name) or str(item_name).strip() == "": continue
            
            for phase_name, p_start, p_end, a_start, a_prog, a_end in phases:
                # 1. Plan Bar (Grey)
                if pd.notnull(row[p_start]) and pd.notnull(row[p_end]):
                    gantt_data.append(dict(
                        Item=item_name, 
                        Y_Label=f"{item_name}", 
                        Phase="Plan",
                        Start=row[p_start], 
                        Finish=row[p_end],
                        ColorKey="Plan" 
                    ))
                
                # 2. Actual Bar (Colored by Phase)
                if pd.notnull(row[a_start]):
                     start_date = row[a_start]
                     # If end date is missing, assume it's ongoing (ends today)
                     finish_date = row[a_end] if pd.notnull(row[a_end]) else pd.Timestamp.now().date()

                     gantt_data.append(dict(
                        Item=item_name, 
                        Y_Label=f"{item_name}", 
                        Phase=phase_name, 
                        Start=start_date, 
                        Finish=finish_date,
                        ColorKey=phase_name 
                    ))

        if not gantt_data: return None
        
        g_df = pd.DataFrame(gantt_data)
        g_df['Start'] = pd.to_datetime(g_df['Start'])
        g_df['Finish'] = pd.to_datetime(g_df['Finish'])
        
        # Define color map including 'Plan'
        color_map = {'Plan': '#d3d3d3'} # Light Grey
        color_map.update(phase_colors)
        
        # To make Plan appear "behind" or clearly distinguishable, we might want to separate rows or overlap.
        # User said "Plan Grey, Actual Color Comparison".
        # If we share Y_Label, they overlap in 'stack' mode (not ideal) or 'group' mode.
        # 'overlay' mode isn't standard in timeline. 
        # Best approach for comparison: Two rows per Item? 
        # "Item 1 (Plan)" and "Item 1 (Actual)"?
        # User request: "Gantt차트를 이용하여 plan 차트 회색으로 실행차트는 비교하는 것으로 항목별... 날짜로 표현해줘."
        # Let's align them on the SAME row if possible, but distinct visuals, OR separate rows.
        # Separate rows (Plan row, Actual row) is clearest for Gantt.
        # Let's adjust Y_Label to separate Plan/Actual.
        
        g_df['Y_Cat'] = g_df.apply(lambda x: x['Item'] if x['ColorKey'] != 'Plan' else x['Item'] + " (Plan)", axis=1) # Naive approach
        # Better: Group by Item, but differentiate Plan/Actual bars.
        # Let's try: Item Name as Y Axis.
        # But distinguish bars by opacity or width? Plotly Express timeline is limited.
        
        # Let's go with: 
        # Row 1: Item A (Plan) -> Grey Bars
        # Row 2: Item A (Actual) -> Colored Bars
        # This is clear.
        
        
        g_df['Y_Label_Final'] = g_df.apply(lambda x: f"{x['Item']} [Plan]" if x['ColorKey'] == 'Plan' else f"<span style='color: #0000FF; font-weight: bold; font-size: 14px;'>{x['Item']} [Actual]</span>", axis=1)
        
        # Sorting to keep Plan/Actual together
        g_df.sort_values(by=['Item', 'ColorKey'], ascending=[True, False], inplace=True) 
        # Plan (P) vs Phase Name... P comes after most? 
        # Let's force verify order.
        
        fig = px.timeline(
            g_df, x_start="Start", x_end="Finish", y="Y_Label_Final", color="ColorKey",
            color_discrete_map=color_map,
            opacity=0.9,
            hover_data=["Item", "Phase", "Start", "Finish"],
            title="상세 공정 비교 (Plan vs Actual)"
        )
        
        fig.update_yaxes(
            autorange="reversed", 
            title_text="항목 (Item)",
            showgrid=True,
            gridwidth=1,          # Thicker line
            gridcolor='#888888',  # Darker grey for clear separation
            zeroline=True,
            zerolinewidth=2,
            zerolinecolor='#888888'
        )
        fig.update_xaxes(
            type='date',
            showgrid=True,
            gridwidth=0.5,
            gridcolor='#E0E0E0',
            dtick=864000000.0, # 10 Days
            tickformat="%m-%d"
        )
        fig.update_layout(height=max(600, len(df)*50), showlegend=True, template='plotly_white') # Force white template
        return fig

    def create_data_table_html(df, phases):
        # Select columns: Item, Weight, Prev Actual, Curr Actual, Monthly Progress, Duration
        # Ensure these columns exist
        base_cols = ['항목 (Item)', '가중치 (Weight)', '전월 실적 (Actual Prev)', '금월 실적 (Actual Curr)', '월간 진도 (Monthly Progress)', '제작 기간 (Weeks)']
        cols_to_show = [c for c in base_cols if c in df.columns]
        
        # Add date columns from phases
        date_cols = []
        for p in phases:
             # Add Plan Start/End, Actual Start/Progress/End
             date_cols.extend([p[1], p[2], p[3], p[4], p[5]])
        
        # Filter only existing columns
        existing_date_cols = [c for c in date_cols if c in df.columns]
        final_cols = cols_to_show + existing_date_cols
        
        table_df = df[final_cols].copy()
        
        # Format Numeric Columns to 2 decimal places
        numeric_format_cols = ['가중치 (Weight)', '전월 실적 (Actual Prev)', '금월 실적 (Actual Curr)', '월간 진도 (Monthly Progress)']
        for col in numeric_format_cols:
            if col in table_df.columns:
                 # Check if numeric
                 table_df[col] = pd.to_numeric(table_df[col], errors='coerce').fillna(0)
                 table_df[col] = table_df[col].apply(lambda x: f"{x:.2f}")

        # Format Dates
        for col in existing_date_cols:
            table_df[col] = pd.to_datetime(table_df[col]).dt.strftime('%Y-%m-%d').fillna('-')
            
        # Rename columns for better readability (Optional)
        # e.g. remove ' (Item)' etc.
        
        # Convert to HTML
        html = table_df.to_html(index=False, classes='data-table', border=0)
        return html

    # --- 4. Main Chart View (Tabs Removed) ---
    st.subheader("📅 통합 공정 스케줄 (Project Schedule Gantt)")
    
    fig_gantt = create_gantt_chart(edited_df, phases_info, f"통합 공정 스케줄 ({project_name})")
    if fig_gantt:
        # Add Delivery Line
        delivery_ts = pd.to_datetime(contract_delivery_date).timestamp() * 1000
        fig_gantt.add_vline(x=delivery_ts, line_width=2, line_dash="dash", line_color="red", annotation_text="계약 납품일")
        fig_gantt.update_layout(template='plotly_white') # Ensure white background
        st.plotly_chart(fig_gantt, use_container_width=True)
    else:
        st.info("차트를 표시할 날짜 데이터가 부족합니다.")
        
    # 엑셀 다운로드 (자동 계산된 데이터 포함)
    st.markdown("---")
    
    # 엑셀 다운로드 (In-Memory)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False, sheet_name="Schedule")
        
        # Metadata
        meta_data = {
            'ProjectName': [project_name],
            'StartDate': [project_start_date],
            'DeliveryDate': [contract_delivery_date]
        }
        pd.DataFrame(meta_data).to_excel(writer, index=False, sheet_name="ProjectInfo")
    
    output.seek(0)
    
    server_filename = f"{project_name}_Schedule_Calculated.xlsx"
    import re
    safe_server_name = re.sub(r'[\\/*?:"<>|]', "", server_filename).strip()

    st.download_button(
        label="💾 엑셀 스케줄 다운로드 (Download Excel)",
        data=output,
        file_name=safe_server_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # temp file for stability (optional, can keep or remove, keeping for now)
    temp_filename = "temp_export.xlsx"
    with pd.ExcelWriter(temp_filename, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False, sheet_name="Schedule")

    # ... [Existing Chart Code] ...
    
    # --- 5. Report Generation ---
    st.markdown("---")
    st.subheader("📑 보고서 생성 (Report Generation)")
    
    import base64
    
    def get_image_base64(path):
        try:
            with open(path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode()
        except Exception:
            return ""

    # Session State for Report
    if 'report_html' not in st.session_state:
        st.session_state.report_html = None
    if 'report_name' not in st.session_state:
        st.session_state.report_name = None

    if st.button("🔄 종합 보고서 생성 (Generate Report)"):
        with st.spinner("보고서를 생성 중입니다... (Generating Report...)"):
            # 1. Prepare Assets
            
            # 2. Capture Charts (Plotly to HTML div)
            
            # Chart 1: Gantt Chart
            fig_gantt = create_gantt_chart(edited_df, phases_info, "") # Clean title for report
            if fig_gantt:
                # Add Delivery Line
                delivery_ts = pd.to_datetime(contract_delivery_date).timestamp() * 1000
                fig_gantt.add_vline(x=delivery_ts, line_width=2, line_dash="dash", line_color="red", annotation_text="계약 납품일")
                gantt_html = fig_gantt.to_html(full_html=False, include_plotlyjs='cdn')
            else:
                gantt_html = "<p>일정 데이터 부족</p>"

            # 3. Data Table HTML (Existing Function)
            data_table_html = create_data_table_html(edited_df, phases_info)
            
            # 4. Prepare New Sections
            
            # A. Overall Metrics HTML
            diff_val = overall_actual - overall_plan
            diff_color = "red" if diff_val < 0 else "green"
            diff_sign = "" if diff_val < 0 else "+"
            
            metrics_html = f"""
            <div class="metrics-container" style="display: flex; gap: 20px; justify-content: space-between; background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                <div class="metric-card" style="flex: 1; text-align: center; background: white; padding: 15px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <h3 style="margin-top: 0; color: #555;">전체 계획 공정률</h3>
                    <p style="font-size: 24px; font-weight: bold; margin: 0; color: #0056b3;">{overall_plan:.2f}%</p>
                </div>
                <div class="metric-card" style="flex: 1; text-align: center; background: white; padding: 15px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <h3 style="margin-top: 0; color: #555;">전체 실적 공정률</h3>
                    <p style="font-size: 24px; font-weight: bold; margin: 0; color: #0056b3;">{overall_actual:.2f}% <span style="font-size: 16px; color: {diff_color};">({diff_sign}{diff_val:.2f}%)</span></p>
                </div>
                <div class="metric-card" style="flex: 1; text-align: center; background: white; padding: 15px; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <h3 style="margin-top: 0; color: #555;">종합 상태</h3>
                    <p style="font-size: 24px; font-weight: bold; margin: 0; color: #333;">{status_msg}</p>
                </div>
            </div>
            """
            
            # B. Detailed Progress Review Table HTML
            review_cols = ['항목 (Item)', '가중치 (Weight)', '전월 실적 (Actual Prev)', '금월 실적 (Actual Curr)', '월간 진도 (Monthly Progress)']
            final_review_cols = [c for c in review_cols if c in edited_df.columns]
            
            # Create a copy for formatting
            review_df = edited_df[final_review_cols].copy()
            for col in final_review_cols:
                if col in ['가중치 (Weight)', '전월 실적 (Actual Prev)', '금월 실적 (Actual Curr)', '월간 진도 (Monthly Progress)']:
                     review_df[col] = pd.to_numeric(review_df[col], errors='coerce').fillna(0).apply(lambda x: f"{x:.2f}%")
            
            review_table_html = review_df.to_html(index=False, classes='data-table', border=0)

            # 5. HTML Template
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>{project_name} - Monthly Progress Report</title>
                <style>
                    body {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: #333; line-height: 1.6; max-width: 1200px; margin: 0 auto; padding: 40px; }}
                    .page-break {{ page-break-before: always; }}
                    .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0056b3; padding-bottom: 20px; margin-bottom: 30px; }}
                    .logo {{ font-size: 24px; font-weight: bold; color: #0056b3; }}
                    .title-box {{ text-align: right; }}
                    .title {{ font-size: 28px; font-weight: bold; margin: 0; color: #2c3e50; }}
                    .subtitle {{ font-size: 14px; color: #7f8c8d; margin-top: 5px; }}
                    
                    .section {{ margin-bottom: 50px; }}
                    .section-title {{ font-size: 20px; font-weight: bold; color: #0056b3; border-bottom: 1px solid #eee; padding-bottom: 10px; margin-bottom: 20px; }}
                    
                    table.data-table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 11px; }}
                    table.data-table th {{ background-color: #0056b3; color: white; padding: 8px; text-align: center; border: 1px solid #ddd; }}
                    table.data-table td {{ padding: 6px; border: 1px solid #ddd; text-align: center; }}
                    table.data-table tr:nth-child(even) {{ background-color: #f9f9f9; }}
                    table.data-table tr:hover {{ background-color: #f1f1f1; }}
                    td {{ padding: 10px; border-bottom: 1px solid #ddd; }}
                    tr:nth-child(even) {{ background-color: #f2f2f2; }}
                    
                    @media print {{
                        .page-break {{ break-before: page; }}
                        body {{ padding: 0; }}
                    }}
                </style>
            </head>
            <body>
                <!-- Header -->
                <div class="header">
                    <div class="logo">EMKO</div>
                    <div class="title-box">
                        <div class="title">Monthly Progress Report</div>
                        <div class="subtitle">Project: {project_name}</div>
                        <div class="subtitle">Date: {pd.Timestamp.now().strftime('%Y-%m-%d')}</div>
                    </div>
                </div>

                <!-- 1. Overall Metrics (New) -->
                <div class="section">
                    <div class="section-title">📊 종합 공정 현황 (Overall Status)</div>
                    {metrics_html}
                </div>

                <!-- Issues & Delays -->
                <div class="section">
                     <div class="section-title">🚨 주요 이슈 및 지연 알림 (Major Issues)</div>
                     <ul>
                     {''.join([f'<li style="color:red; font-weight:bold;">{alert}</li>' for alert in delay_alerts]) if delay_alerts else '<li>No major issues found. (정상)</li>'}
                     </ul>
                </div>

                <div class="page-break"></div>

                <!-- Gantt Chart -->
                <div class="section">
                    <div class="section-title">📅 통합 공정 스케줄 (Project Schedule)</div>
                    <div style="width:100%; overflow-x: auto;">
                        {gantt_html}
                    </div>
                </div>

                <div class="page-break"></div>
                
                <!-- 2. Detailed Progress Review (New) -->
                <div class="section">
                    <div class="section-title">📋 상세 진도율 검토 (Detailed Progress Review)</div>
                    {review_table_html}
                </div>
                
                <div class="page-break"></div>

                <!-- Detailed Data (Full Table) -->
                <div class="section">
                    <div class="section-title">📑 전체 데이터 (Full Data)</div>
                    {data_table_html}
                </div>
                
                <div class="footer">
                    &copy; {date.today().year} EMKO. All rights reserved. Generated by Gantt Chat Project.
                </div>
            </body>
            </html>
            """
            
            # Save to Session State
            st.session_state.report_html = html_content
            
            report_filename = f"{project_name}_Progress_Report.html"
            safe_report_name = re.sub(r'[\\/*?:"<>|]', "", report_filename).strip()
            st.session_state.report_name = safe_report_name
            
            st.success("보고서가 생성되었습니다! 아래 다운로드 버튼을 눌러주세요.")

    # Show Download Button if Report is Ready
    if st.session_state.report_html:
        st.download_button(
            label="💾 종합 보고서 다운로드 (Download Report)",
            data=st.session_state.report_html,
            file_name=st.session_state.report_name,
            mime="text/html"
        )

except Exception as e:
    st.error(f"오류 발생: {e}")
    st.text(traceback.format_exc())
