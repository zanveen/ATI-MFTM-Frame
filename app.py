import streamlit as st
import json
import copy
import os
import base64
import pandas as pd
import gspread
import zipfile
import calendar
from google.oauth2.service_account import Credentials
from datetime import datetime, date, timedelta
from pathlib import Path
from openpyxl import load_workbook
import io

# ─── 설정 ───
st.set_page_config(
    page_title="Frame 제작 진행현황 관리",
    layout="wide",
    initial_sidebar_state="expanded"
)

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
PROJECTS_FILE = DATA_DIR / "projects.json"
TEMPLATE_FILE = Path("template.xlsx")
LOGO_FILE = "ati_logo.png"

# ─── 사용자 권한 정보 ───
USER_CREDENTIALS = {
    "rladbstn5344": {"role": "admin", "name": "관리자", "company": "ALL"},
    "11053": {"role": "vendor", "name": "한울산업", "company": "한울산업"},
    "10738": {"role": "vendor", "name": "정한테크", "company": "정한테크"}
}

# ─── 체크리스트 항목 정의 ───
CHECKLIST_ITEMS = [
    {"no": 1, "category": "판금설계", "item": "수정 도면 반영 확인"},
    {"no": 2, "category": "제작", "item": "용접 방향 확인"},
    {"no": 3, "category": "제작", "item": "비드 상태 확인"},
    {"no": 4, "category": "제작", "item": "주요 치수 정밀 검사 (대각 기장, 폭, 높이 등)"},
    {"no": 5, "category": "제작", "item": "절곡부 마감 상태"},
    {"no": 6, "category": "제작", "item": "SUS FRAME 얼룩 확인"},
    {"no": 7, "category": "제작", "item": "용접 후 뒤틀림 및 열변형 교정 상태 확인"},
    {"no": 8, "category": "도장", "item": "도장일 확인 및 건조 상태"},
    {"no": 9, "category": "도장", "item": "도장면 외관 검사 (오염, 흐름, 뭉침, 찍힘 등)"},
    {"no": 10, "category": "도장", "item": "마스킹 부위 및 탭(Tap) 구멍 도료 유입 여부"},
    {"no": 11, "category": "조립", "item": "도어 미부착 개소 있는지 확인"},
    {"no": 12, "category": "조립", "item": "도어 개폐시 간섭 혹은 소음 확인"},
    {"no": 13, "category": "조립", "item": "최종 조립 수평 상태 및 프레임 단차 확인"},
    {"no": 14, "category": "조립", "item": "명판, 경고 라벨 등 부착물 상태 확인"},
    {"no": 15, "category": "납품", "item": "납품일까지 미입고품 확인 및 전달"},
    {"no": 16, "category": "납품", "item": "레벨풋 장착 확인"},
    {"no": 17, "category": "납품", "item": "프레임 리스트 동봉 확인"},
    {"no": 18, "category": "절곡 단품", "item": "납품일까지 미입고품 확인 및 전달"},
    {"no": 19, "category": "절곡 단품", "item": "표면 스크래치, 찍힘 및 모서리 버(Burr) 제거 상태"},
    {"no": 20, "category": "절곡 단품", "item": "파트별 수량 확인 및 라벨링(식별표) 부착 상태"},
]

# ─── 로고 이미지 함수 ───
def get_logo_html(height="34px"):
    if os.path.exists(LOGO_FILE):
        try:
            with open(LOGO_FILE, "rb") as f:
                encoded = base64.b64encode(f.read()).decode()
            return f'<img src="data:image/png;base64,{encoded}" style="height:{height}; vertical-align:middle; margin-right:8px; margin-bottom:4px;">'
        except: pass
    return "🏭 "

# ─── 🛡️ 구글 시트 연동 및 복구 로직 (에러 완벽 차단) ───
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(credentials)

def load_from_sheets():
    try:
        client = get_gspread_client()
        sh = client.open("Frame_Data")
        worksheet = sh.sheet1 # 에러의 원인이었던 부분 완벽 수정!!
        all_records = worksheet.get_all_records()
        
        projects = {}
        for row in all_records:
            pid = str(row['pid'])
            projects[pid] = {
                "info": {
                    "company": str(row.get('company', '미상')),
                    "equipment": str(row.get('equipment', '미상')),
                    "order_date": str(row.get('order_date', '')),
                    "delivery_date": str(row.get('delivery_date', '')),
                    "frame_parts": int(row.get('frame_parts', 1)) if row.get('frame_parts') else 1,
                    "frame_options": json.loads(row.get('frame_options', '[]')) if row.get('frame_options') else [],
                    "exterior_spec": str(row.get('exterior_spec', '')),
                    "interior_spec": str(row.get('interior_spec', '')),
                    "notes_top": str(row.get('notes_top', '')),
                    "delivery_delay_count": int(row.get('delivery_delay_count', 0)) if row.get('delivery_delay_count') else 0,
                    "delay_request": json.loads(row.get('delay_request', '{}')) if row.get('delay_request') else {}
                },
                "checks": json.loads(row.get('checks', '{}')) if row.get('checks') else {},
                "special_notes": str(row.get('special_notes', '')),
                "history": json.loads(row.get('history', '[]')) if row.get('history') else []
            }
        return projects
    except Exception as e:
        # 시트가 비어있거나 처음 실행될 때 무시하고 빈 딕셔너리 반환
        return {}

def save_to_sheets(projects):
    try:
        client = get_gspread_client()
        sh = client.open("Frame_Data")
        worksheet = sh.sheet1
        
        header = ["pid", "company", "equipment", "order_date", "delivery_date", "frame_parts", 
                  "frame_options", "exterior_spec", "interior_spec", "notes_top", 
                  "delivery_delay_count", "delay_request", "checks", "special_notes", "history"]
        
        data_to_save = [header]
        for pid, p in projects.items():
            info = p['info']
            data_to_save.append([
                pid, info.get('company',''), info.get('equipment',''), info.get('order_date',''), info.get('delivery_date',''),
                info.get('frame_parts', 1), json.dumps(info.get('frame_options', []), ensure_ascii=False),
                info.get('exterior_spec',''), info.get('interior_spec',''), info.get('notes_top',''),
                info.get('delivery_delay_count', 0), json.dumps(info.get('delay_request', {}), ensure_ascii=False),
                json.dumps(p.get('checks', {}), ensure_ascii=False), p.get('special_notes',''), 
                json.dumps(p.get('history', []), ensure_ascii=False)
            ])
        
        worksheet.clear()
        try:
            worksheet.update(values=data_to_save, range_name='A1')
        except TypeError: # 호환성 처리
            worksheet.update('A1', data_to_save)
            
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

# ─── 계산 및 유틸 함수 ───
def calc_progress(checks):
    return sum(1 for c in checks.values() if c.get("status") == "확인") / 20

def calc_score(checks):
    return sum(5 if c.get("status") == "확인" else 2 if c.get("status") == "미비" else 0 for c in checks.values())

def get_project_status(proj):
    progress = calc_progress(proj.get("checks", {}))
    dd = proj["info"].get("delivery_date", "")
    if progress >= 1.0: return "완료"
    if dd:
        try:
            d = datetime.strptime(dd, "%Y-%m-%d").date()
            if 0 <= (d - date.today()).days <= 7: return "납기임박"
        except: pass
    return "진행중"

def filter_projects_by_role(all_projects, user_info):
    if user_info["role"] == "admin": return all_projects
    return {pid: p for pid, p in all_projects.items() if p["info"]["company"] == user_info["company"]}

def generate_checklist_excel(project):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb["점검표"]
    info = project["info"]
    checks = project.get("checks", {})

    ws["C2"], ws["C3"], ws["C4"], ws["C5"] = info.get("company", ""), info.get("equipment", ""), info.get("order_date", ""), info.get("delivery_date", "")
    ws["C6"], ws["C7"], ws["C8"] = info.get("exterior_spec", ""), info.get("interior_spec", ""), info.get("frame_parts", "")
    
    frame_options = info.get("frame_options", [])
    notes = info.get("notes_top", "")
    combined_d3 = ""
    if frame_options: combined_d3 += f"[옵션: {', '.join(frame_options)}]\n"
    if notes: combined_d3 += notes
    if combined_d3: ws["D3"] = combined_d3.strip()

    for i, item in enumerate(CHECKLIST_ITEMS):
        row = 13 + i
        key = str(item["no"])
        status = checks.get(key, {}).get("status", "")
        if status == "확인": ws[f"D{row}"] = "○"
        elif status == "미비": ws[f"E{row}"] = "○"
        c_date = checks.get(key, {}).get("date", "")
        if c_date: ws[f"F{row}"] = c_date
        c_memo = checks.get(key, {}).get("memo", "")
        if c_memo: ws[f"H{row}"] = c_memo

    ws["F34"] = f"{calc_score(checks)} / 100" 
    ws["A37"] = project.get("special_notes", "")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─── CSS 스타일 (사이드바 20:80 및 버튼 강제 고정) ───
st.markdown("""
<style>
    .login-container { max-width: 420px; margin: 80px auto; padding: 40px; background: white; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); text-align: center; border-top: 6px solid #4A90D9; }
    
    [data-testid="stSidebar"] { min-width: 20vw !important; max-width: 20vw !important; background-color: #f0f4f8; }
    [data-testid="stSidebar"] [data-testid="stRadio"] label { font-size: 20px !important; line-height: 2.2 !important; padding: 8px 0 !important; }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] span { font-size: 18px !important; line-height: 1.8 !important; }
    
    .main .block-container p, .main .block-container span, .main .block-container label, .main .block-container li, .main .block-container td, .main .block-container th { font-size: 18px !important; line-height: 1.9 !important; }
    .main .block-container h1 { font-size: 34px !important; font-weight: 800; color: #1e293b; }
    
    .project-card { background: white; border-radius: 12px; padding: 22px; margin-bottom: 14px; border-left: 5px solid #4A90D9; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .progress-bar-bg { background: #e8ecf1; border-radius: 10px; height: 26px; overflow: hidden; margin: 12px 0 8px 0; }
    .progress-bar-fill { height: 100%; border-radius: 10px; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 13px; transition: width 0.4s ease; }
    .status-badge { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 14px; font-weight: 700; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .badge-green { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
    .badge-yellow { background: #fff3cd; color: #856404; border: 1px solid #ffeeba; }
    .badge-red { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
    .badge-blue { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
    .category-header { background: #4A90D9; color: white; padding: 10px 18px; border-radius: 8px; margin: 22px 0 10px 0; font-weight: bold; font-size: 19px; }

    /* 대시보드 4개 버튼 */
    div[data-testid="stHorizontalBlock"] button[kind="primary"] {
        height: 115px !important; border-radius: 16px !important; background-color: white !important;
        border: 2px solid #cbd5e1 !important; transition: transform 0.2s cubic-bezier(0.175, 0.885, 0.32, 1.275), box-shadow 0.2s, border-color 0.2s !important; padding: 0 !important;
    }
    div[data-testid="stHorizontalBlock"] button[kind="primary"]:hover {
        transform: scale(1.05) !important; box-shadow: 0 8px 16px rgba(0,0,0,0.1) !important; border-color: #4A90D9 !important; z-index: 10 !important;
    }
    div[data-testid="stHorizontalBlock"] button[kind="primary"] p {
        display: block !important; white-space: pre-wrap !important; font-size: 17px !important; font-weight: 600 !important; color: #475569 !important; text-align: center; margin: 0; line-height: 1.4;
    }
    div[data-testid="stHorizontalBlock"] button[kind="primary"] p::first-line {
        font-size: 42px !important; font-weight: 900 !important; color: #4A90D9 !important; line-height: 1.2;
    }

    /* 🔥 4번째 컬럼(납기임박) 완벽한 빨간색 고정 🔥 */
    div[data-testid="column"]:nth-of-type(4) button[kind="primary"] p { color: #e74c3c !important; }
    div[data-testid="column"]:nth-of-type(4) button[kind="primary"] p::first-line { color: #e74c3c !important; }
    div[data-testid="column"]:nth-of-type(4) button[kind="primary"] { border-color: #fca5a5 !important; }
    div[data-testid="column"]:nth-of-type(4) button[kind="primary"]:hover { border-color: #e74c3c !important; background-color: #fef2f2 !important; }
</style>
""", unsafe_allow_html=True)

# ─── 세션 초기화 및 시트 로드 ───
if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "user_info" not in st.session_state: st.session_state.user_info = None
if "dashboard_filter" not in st.session_state: st.session_state.dashboard_filter = "전체"
if "inspection_project" not in st.session_state: st.session_state.inspection_project = None

if "projects" not in st.session_state or not st.session_state.projects:
    st.session_state.projects = load_from_sheets()

# ═══════════════════════════════════════════════════
# 🔐 1. 로그인
# ═══════════════════════════════════════════════════
if not st.session_state.logged_in:
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown(f"""
        <div class="login-container">
            <div style="font-size: 28px; font-weight: 800; color: #1a1a2e; margin-bottom: 20px;">
                {get_logo_html('34px')} FRAME 제작 통합 관리
            </div>
            <p style="color:#666; margin-bottom:30px;">부여받은 비밀번호를 입력해주세요.</p>
        """, unsafe_allow_html=True)
        pwd = st.text_input("🔑 비밀번호", type="password", label_visibility="collapsed")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("로그인", use_container_width=True, type="primary"):
            if pwd in USER_CREDENTIALS:
                st.session_state.logged_in = True
                st.session_state.user_info = USER_CREDENTIALS[pwd]
                st.session_state.projects = load_from_sheets() # 시트 최신화
                st.rerun()
            else:
                st.error("비밀번호가 일치하지 않습니다.")
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════
# 🖥️ 2. 메인 애플리케이션
# ═══════════════════════════════════════════════════
else:
    user = st.session_state.user_info

    # ─── 사이드바 ───
    st.sidebar.markdown(f"""
        <div style="font-size:30px; font-weight:bold; margin-bottom:20px; line-height:1.2;">
            {get_logo_html('32px')} FRAME 관리
        </div>
    """, unsafe_allow_html=True)
        
    st.sidebar.markdown(f"**접속자: {user['name']}님**")
    st.sidebar.markdown("---")

    menu_options = ["📊 대시보드", "📅 납기 캘린더", "📋 점검 기록", "📥 체크리스트 추출"]
    if user["role"] == "admin":
        menu_options.insert(1, "➕ 프로젝트 등록")

    menu = st.sidebar.radio("메뉴", menu_options, label_visibility="collapsed")
    if menu != "📋 점검 기록": st.session_state.inspection_project = None

    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 로그아웃", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.user_info = None
        st.rerun()
        
    projects = filter_projects_by_role(st.session_state.projects, user)

    # ═══════════════════════════════════════════════════
    # 📊 대시보드
    # ═══════════════════════════════════════════════════
    if menu == "📊 대시보드":
        col_title, col_stats = st.columns([5, 5])
        with col_title:
            st.markdown(f"""
                <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px;">
                    {get_logo_html('34px')} {'전체 진행 현황 대시보드' if user['role']=='admin' else '우리 회사 배정 프로젝트'}
                </div>
            """, unsafe_allow_html=True)
            
        with col_stats:
            if user["role"] == "admin":
                pending_reqs = [p for p in projects.values() if p["info"].get("delay_request", {}).get("status") == "pending"]
                if pending_reqs:
                    st.markdown(f"""
                    <div style="background-color:#fff3cd; border-left:5px solid #ffc107; padding:12px; border-radius:6px; margin-bottom:15px;">
                        <h4 style="margin:0; color:#856404; font-size:18px;">🚨 납기 변경 요청 접수 ({len(pending_reqs)}건)</h4>
                        <p style="margin:4px 0 0 0; font-size:14px; color:#856404;">업체에서 납기일 조율을 요청했습니다. '점검 기록' 메뉴에서 승인해주세요.</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                months = set(p["info"].get("delivery_date", "")[:7] for p in st.session_state.projects.values() if p["info"].get("delivery_date"))
                month_list = sorted(list(months), reverse=True)
                selected_period = st.selectbox("📅 통계 조회 기간", ["전체 누적"] + month_list)
                
                h_delays, h_projs, j_delays, j_projs = 0, 0, 0, 0
                for p in st.session_state.projects.values():
                    dd = p["info"].get("delivery_date", "")
                    if selected_period != "전체 누적" and not dd.startswith(selected_period): continue
                    comp, delay = p["info"].get("company", ""), p["info"].get("delivery_delay_count", 0)
                    if comp == "한울산업": h_delays += delay; h_projs += 1
                    elif comp == "정한테크": j_delays += delay; j_projs += 1

                st.markdown(f"""
                <div style="background-color:#f8f9fa; border:1px solid #dee2e6; padding:15px; border-radius:8px;">
                    <b style="color:#2c3e50; font-size:16px;">📈 {selected_period} 업체별 현황</b>
                    <div style="display:flex; gap:10px; margin-top:10px;">
                        <div style="flex:1; background:white; padding:12px; border-radius:6px; border:1px solid #e0e0e0; text-align:center;">
                            <b style="color:#4A90D9; font-size:16px;">한울산업</b><br>
                            <span style="font-size:14px;">진행 프로젝트: <b>{h_projs}</b>건</span><br>
                            <span style="font-size:14px; color:#e74c3c;">납기 지연: <b>{h_delays}</b>회</span>
                        </div>
                        <div style="flex:1; background:white; padding:12px; border-radius:6px; border:1px solid #e0e0e0; text-align:center;">
                            <b style="color:#4A90D9; font-size:16px;">정한테크</b><br>
                            <span style="font-size:14px;">진행 프로젝트: <b>{j_projs}</b>건</span><br>
                            <span style="font-size:14px; color:#e74c3c;">납기 지연: <b>{j_delays}</b>회</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
        st.markdown("---")

        if not projects:
            st.info("표시할 프로젝트가 없습니다.")
        else:
            completed_pids = [pid for pid, p in projects.items() if get_project_status(p) == "완료"]
            urgent_pids = [pid for pid, p in projects.items() if get_project_status(p) == "납기임박"]
            in_progress_pids = [pid for pid, p in projects.items() if get_project_status(p) == "진행중"]

            filters = [("전체", len(projects), "전체 프로젝트"), ("진행중", len(in_progress_pids), "진행중"), ("완료", len(completed_pids), "완료"), ("납기임박", len(urgent_pids), "납기임박")]
            cols = st.columns(4)
            for i, (key, num, label) in enumerate(filters):
                with cols[i]:
                    if st.session_state.dashboard_filter == key:
                        active_border, active_bg = ("#e74c3c", "#fef2f2") if key == "납기임박" else ("#4A90D9", "#eff6ff")
                        st.markdown(f"""
                        <style>
                        div[data-testid="column"]:nth-of-type({i+1}) button[kind="primary"] {{
                            transform: scale(1.06) !important; border: 3px solid {active_border} !important; background-color: {active_bg} !important; box-shadow: 0 8px 20px rgba(0,0,0,0.1) !important;
                        }}
                        </style>
                        """, unsafe_allow_html=True)
                    if st.button(f"{num}\n{label}", key=f"filter_{key}", use_container_width=True, type="primary"):
                        st.session_state.dashboard_filter = key
                        st.rerun()

            show_pids = list(projects.keys()) if st.session_state.dashboard_filter == "전체" else in_progress_pids if st.session_state.dashboard_filter == "진행중" else completed_pids if st.session_state.dashboard_filter == "완료" else urgent_pids
            st.markdown("<br>", unsafe_allow_html=True)

            if not show_pids:
                st.info("해당하는 프로젝트가 없습니다.")
            else:
                sort_option = st.selectbox("정렬", ["납품 예정일순", "진척률 낮은순", "진척률 높은순"], label_visibility="collapsed")
                sorted_pids = show_pids.copy()
                if sort_option == "납품 예정일순": sorted_pids.sort(key=lambda x: projects[x]["info"].get("delivery_date", "9999"))
                elif sort_option == "진척률 낮은순": sorted_pids.sort(key=lambda x: calc_progress(projects[x].get("checks", {})))
                else: sorted_pids.sort(key=lambda x: calc_progress(projects[x].get("checks", {})), reverse=True)

                for pid in sorted_pids:
                    proj = projects.get(pid)
                    if not proj: continue
                    
                    info = proj["info"]
                    checks = proj.get("checks", {})
                    pct = int(calc_progress(checks) * 100)
                    score = calc_score(checks)
                    delay_cnt = info.get("delivery_delay_count", 0)

                    if pct >= 100: badge, bar_color = '<span class="status-badge badge-green">✅ 완료</span>', "#27ae60"
                    elif pct >= 50: badge, bar_color = '<span class="status-badge badge-blue">🔧 진행중</span>', "#4A90D9"
                    elif pct > 0: badge, bar_color = '<span class="status-badge badge-yellow">⚠️ 초기</span>', "#f39c12"
                    else: badge, bar_color = '<span class="status-badge badge-red">⏳ 미착수</span>', "#e74c3c"

                    dd = info.get("delivery_date", "")
                    days_text = ""
                    if dd:
                        try:
                            diff = (datetime.strptime(dd, "%Y-%m-%d").date() - date.today()).days
                            if diff < 0: days_text = f'<span style="color:#e74c3c;font-weight:bold;">⚠️ 납기 {abs(diff)}일 초과</span>'
                            elif diff == 0: days_text = '<span style="color:#e74c3c;font-weight:bold;">📌 오늘 납품!</span>'
                            elif diff <= 7: days_text = f'<span style="color:#e74c3c;font-weight:bold;">⏰ D-{diff}</span>'
                            else: days_text = f'<span style="color:#888;">D-{diff}</span>'
                        except: pass

                    opts_html = f'<span style="background:#eff6ff; color:#3b82f6; padding:4px 12px; border-radius:20px; font-size:14px; font-weight:700; border:1px solid #bfdbfe; margin-left:15px;">옵션: {", ".join(info.get("frame_options", []))}</span>' if info.get("frame_options") else ''
                    delay_text = f" &nbsp;│ <span style='color:#e74c3c;font-weight:bold;'>납기변경 {delay_cnt}회</span>" if delay_cnt > 0 else ""

                    st.markdown(f"""
                    <div class="project-card">
                        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap; margin-bottom:10px;">
                            <div style="display:flex; align-items:center;">
                                <h3 style="margin:0;">🔩 {info.get('equipment', pid)}</h3>{opts_html}
                            </div>
                            <div style="display:flex; align-items:center; gap:10px;">
                                {badge}<span>{days_text}</span>
                            </div>
                        </div>
                        <p style="margin-top:5px; color:#475569;">업체: <b style="color:#1e293b;">{info.get('company','-')}</b> │ 납품예정: {dd} │ 검사점수: <b>{score}/100</b> {delay_text}</p>
                        <div class="progress-bar-bg">
                            <div class="progress-bar-fill" style="width:{max(pct,3)}%;background:{bar_color};">{pct}%</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════
    # 📅 납기 캘린더 (HTML 기반 실제 달력)
    # ═══════════════════════════════════════════════════
    elif menu == "📅 납기 캘린더":
        st.markdown(f"""
            <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                {get_logo_html('34px')} 프로젝트 납기 캘린더
            </div>
        """, unsafe_allow_html=True)
        st.markdown("달력에서 프로젝트별 납품 예정일을 한눈에 확인하세요.")
        st.markdown("---")

        # 년도 및 월 선택 UI
        today = date.today()
        col_y, col_m, _ = st.columns([2, 2, 6])
        year = col_y.selectbox("년도", range(today.year - 1, today.year + 3), index=1)
        month = col_m.selectbox("월", range(1, 13), index=today.month - 1)

        # 달력 생성을 위한 HTML/CSS
        cal = calendar.monthcalendar(year, month)
        html_cal = '<table style="width:100%; border-collapse: collapse; background:white; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-radius:10px; overflow:hidden;">'
        
        # 요일 헤더
        html_cal += '<tr>'
        for day_name in ["월", "화", "수", "목", "금", "토", "일"]:
            color = "#e74c3c" if day_name == "일" else "#3b82f6" if day_name == "토" else "#333"
            html_cal += f'<th style="border: 1px solid #e2e8f0; padding: 12px; text-align: center; background-color: #f8fafc; font-size:18px; color:{color};">{day_name}</th>'
        html_cal += '</tr>'
        
        # 달력 날짜 채우기
        for week in cal:
            html_cal += '<tr>'
            for idx, day in enumerate(week):
                if day == 0:
                    html_cal += '<td style="border: 1px solid #e2e8f0; background-color: #f1f5f9; height: 120px;"></td>'
                else:
                    date_str = f"{year}-{month:02d}-{day:02d}"
                    is_today = (date_str == today.strftime("%Y-%m-%d"))
                    bg_color = "#eff6ff" if is_today else "white"
                    day_color = "#e74c3c" if idx == 6 else "#3b82f6" if idx == 5 else "#333"
                    weight = "900" if is_today else "bold"
                    
                    # 해당 날짜에 납품인 프로젝트 찾기
                    day_projs = []
                    for pid, p in projects.items():
                        if p['info']['delivery_date'] == date_str:
                            day_projs.append(p)
                    
                    cell_content = f'<div style="font-weight: {weight}; font-size: 18px; margin-bottom: 8px; color: {day_color};">{day}{" (오늘)" if is_today else ""}</div>'
                    
                    for p in day_projs:
                        status = get_project_status(p)
                        badge_color = "#e74c3c" if status == "납기임박" else "#27ae60" if status == "완료" else "#4A90D9"
                        cell_content += f'<div style="background-color: {badge_color}; color: white; padding: 4px 6px; border-radius: 4px; font-size: 13px; font-weight:bold; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; box-shadow: 0 1px 2px rgba(0,0,0,0.1);">{p["info"]["equipment"]} ({p["info"]["company"]})</div>'
                    
                    html_cal += f'<td style="border: 1px solid #e2e8f0; padding: 10px; height: 120px; vertical-align: top; background-color: {bg_color}; width: 14.28%;">{cell_content}</td>'
            html_cal += '</tr>'
        html_cal += '</table>'
        
        st.markdown(html_cal, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════
    # ➕ 프로젝트 등록 (관리자 전용 / 배치 수정)
    # ═══════════════════════════════════════════════════
    elif menu == "➕ 프로젝트 등록" and user["role"] == "admin":
        st.markdown(f"""
            <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                {get_logo_html('34px')} 신규 프로젝트 등록
            </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        with st.form("new_project_form"):
            company = st.radio("업체명 *", ["한울산업", "정한테크"], horizontal=True)
            equipment = st.text_input("장비명 (설비명) *", placeholder="예: PALM2")
            
            # 🔥 레이아웃 수정: 발주일과 덩어리 수 나란히 배치 🔥
            col_d1, col_d2, _ = st.columns([2, 2, 4])
            with col_d1: order_date = st.date_input("발주일")
            with col_d2: frame_parts = st.number_input("Frame Part 수 (덩어리) *", min_value=1, value=1, step=1)
            
            delivery_date = st.date_input("납품 예정일자")

            st.markdown("**프레임 옵션**")
            opt_cols = st.columns(3)
            with opt_cols[0]: opt_clean = st.checkbox("클린부스")
            with opt_cols[1]: opt_table = st.checkbox("테이블")
            with opt_cols[2]: opt_jig = st.checkbox("전도방지지그")

            frame_options = [opt for opt, checked in zip(["클린부스", "테이블", "전도방지지그"], [opt_clean, opt_table, opt_jig]) if checked]

            st.markdown("---")
            spec_col1, spec_col2 = st.columns(2)
            with spec_col1: exterior_spec = st.radio("외관 사양", ["SUS", "도장"], horizontal=True)
            with spec_col2: interior_spec = st.radio("내부 사양", ["SUS", "도장"], horizontal=True)

            notes_top = st.text_area("특이사항", placeholder="메모 사항", height=80)
            submitted = st.form_submit_button("✅ 구글 시트에 프로젝트 등록", use_container_width=True)

            if submitted:
                if not equipment: st.error("장비명은 필수 입력입니다.")
                else:
                    pid = f"{company}_{equipment}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                    st.session_state.projects[pid] = repair_project({
                        "info": {
                            "company": company, "equipment": equipment, "order_date": str(order_date),
                            "delivery_date": str(delivery_date), "frame_parts": frame_parts,
                            "frame_options": frame_options, "exterior_spec": exterior_spec,
                            "interior_spec": interior_spec, "notes_top": notes_top,
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "delivery_delay_count": 0, "delay_request": {}
                        },
                        "checks": {}, "special_notes": "", "history": []
                    })
                    save_to_sheets(st.session_state.projects)
                    st.success("✅ 구글 시트에 안전하게 등록되었습니다!")
                    st.rerun()

    # ═══════════════════════════════════════════════════
    # 📋 점검 기록
    # ═══════════════════════════════════════════════════
    elif menu == "📋 점검 기록":
        if not projects: st.info("점검할 프로젝트가 없습니다.")
        else:
            active_projects = {pid: p for pid, p in projects.items() if calc_progress(p.get("checks", {})) < 1.0}

            if st.session_state.inspection_project is None:
                st.markdown(f"""
                    <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                        {get_logo_html('34px')} 진행중인 프로젝트 점검
                    </div>
                """, unsafe_allow_html=True)
                
                for comp_name in ["한울산업", "정한테크"]:
                    comp_projs = [(k, v) for k, v in active_projects.items() if v["info"]["company"] == comp_name]
                    if comp_projs:
                        st.markdown(f'<div style="background:#f8f9fa; border-radius:12px; padding:16px 20px 6px 20px; margin-bottom:8px;"><div style="font-size:22px; font-weight:bold; color:#2c3e50; padding-bottom:8px; border-bottom:3px solid #4A90D9;">🏢 {comp_name}</div></div>', unsafe_allow_html=True)
                        for pid, proj in comp_projs:
                            pct = int(calc_progress(proj.get("checks", {})) * 100)
                            dd = proj["info"].get("delivery_date", "")
                            bar_c = ('#27ae60' if pct>=80 else '#4A90D9' if pct>=50 else '#f39c12' if pct>0 else '#e74c3c')
                            
                            col1, col2 = st.columns([8, 2])
                            with col1:
                                st.markdown(f"""
                                <div style="background:white; border-radius:10px; padding:16px; margin-bottom:10px; border-left:5px solid #4A90D9; box-shadow:0 1px 6px rgba(0,0,0,0.07);">
                                    <b style="font-size:18px;">{proj['info']['equipment']}</b> (납기: {dd})
                                    <div style="background:#e8ecf1;border-radius:6px;height:10px;margin-top:8px;">
                                        <div style="background:{bar_c};height:100%;width:{max(pct,2)}%;border-radius:6px;"></div>
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)
                            with col2:
                                st.markdown("<br>", unsafe_allow_html=True)
                                if st.button("점검 →", key=f"ins_{pid}", use_container_width=True):
                                    st.session_state.inspection_project = pid
                                    st.rerun()
            else:
                pid = st.session_state.inspection_project
                proj = projects.get(pid)
                if not proj:
                    st.session_state.inspection_project = None
                    st.rerun()

                info = proj["info"]
                checks = proj.get("checks", {})
                
                if st.button("← 프로젝트 목록으로 돌아가기"):
                    st.session_state.inspection_project = None
                    st.rerun()

                delay_req = info.get("delay_request", {})
                st.markdown("---")
                st.markdown("### 🗓️ 일정 관리 및 조율")
                
                if user["role"] == "vendor":
                    if delay_req and delay_req.get("status") == "pending":
                        st.info(f"⏳ 관리자에게 납기 변경을 요청했습니다. (희망일: {delay_req.get('requested_date')} / 승인 대기중)")
                    else:
                        with st.expander("👉 관리자에게 납기 변경 요청하기"):
                            with st.form(f"delay_form_{pid}"):
                                new_date = st.date_input("변경 희망 납품일")
                                reason = st.text_input("납기 변경(지연) 사유")
                                if st.form_submit_button("요청 전송", use_container_width=True):
                                    proj["info"]["delay_request"] = {"requested_date": str(new_date), "reason": reason, "status": "pending"}
                                    save_to_sheets(st.session_state.projects)
                                    st.success("납기 변경 요청이 전송되었습니다.")
                                    st.rerun()

                if user["role"] == "admin":
                    if delay_req and delay_req.get("status") == "pending":
                        st.warning(f"🚨 **[{info['company']}] 납기 변경 요청** : 기존 {info['delivery_date']} ➡️ **희망 {delay_req.get('requested_date')}** (사유: {delay_req.get('reason')})")
                        col_a, col_b, col_c = st.columns([2, 2, 6])
                        if col_a.button("✅ 요청 승인 (납기 적용)", use_container_width=True):
                            proj["info"]["delivery_date"] = delay_req["requested_date"]
                            proj["info"]["delivery_delay_count"] = info.get("delivery_delay_count", 0) + 1
                            proj["info"]["delay_request"]["status"] = "approved"
                            save_to_sheets(st.session_state.projects)
                            st.rerun()
                        if col_b.button("❌ 요청 반려", use_container_width=True):
                            proj["info"]["delay_request"]["status"] = "rejected"
                            save_to_sheets(st.session_state.projects)
                            st.rerun()
                    
                    with st.expander("🔧 관리자 전용: 납기일 임의 변경 (지연 누적 적용)"):
                        with st.form(f"admin_delay_form_{pid}"):
                            try: default_d = datetime.strptime(info["delivery_date"], "%Y-%m-%d").date()
                            except: default_d = date.today()
                            
                            new_date_admin = st.date_input("새로운 납기일로 변경", value=default_d)
                            st.warning("⚠️ 납기일을 변경하면 해당 업체의 '납기 지연 횟수'가 1회 증가합니다.")
                            confirm_change = st.checkbox("네, 위 내용을 확인하였으며 납기일을 강제로 변경합니다.")
                            
                            if st.form_submit_button("납기일 강제 변경 적용", use_container_width=True):
                                if str(new_date_admin) == info.get("delivery_date", ""):
                                    st.error("기존 납기일과 동일합니다. 다른 날짜를 선택하세요.")
                                elif not confirm_change:
                                    st.error("변경 사항을 적용하시려면 위 확인 체크박스에 체크해주세요.")
                                else:
                                    proj["info"]["delivery_date"] = str(new_date_admin)
                                    proj["info"]["delivery_delay_count"] = info.get("delivery_delay_count", 0) + 1
                                    save_to_sheets(st.session_state.projects)
                                    st.success("✅ 구글 시트에 납기일이 변경/저장되었습니다!")
                                    st.rerun()

                st.markdown("---")
                opts_html = f'<span style="background:#eff6ff; color:#3b82f6; padding:4px 12px; border-radius:20px; font-size:14px; font-weight:700; border:1px solid #bfdbfe; margin-left:15px;">옵션: {", ".join(info.get("frame_options", []))}</span>' if info.get("frame_options") else ''

                st.markdown(f"""
                <div style="background:#eef2f7;padding:16px 22px;border-radius:10px;margin-bottom:24px;">
                    <div style="display:flex; align-items:center; margin-bottom:8px;">
                        <h3 style="margin:0;font-size:23px;">🔩 {info['equipment']}</h3>{opts_html}
                    </div>
                    <p style="margin:4px 0;font-size:17px;">
                        <b>업체:</b> {info['company']} │ <b>납품일:</b> <span style="color:#e74c3c;font-weight:bold;">{info['delivery_date']}</span> (납기변경 누적 {info.get('delivery_delay_count', 0)}회) │
                        <b>진척률:</b> {int(calc_progress(checks)*100)}%
                    </p>
                </div>
                """, unsafe_allow_html=True)

                check_date = st.date_input("📅 당일 점검 일자 지정", value=date.today(), key="cd")
                updated_checks = copy.deepcopy(checks)
                current_cat = ""

                for item in CHECKLIST_ITEMS:
                    key = str(item["no"])
                    if item["category"] != current_cat:
                        current_cat = item["category"]
                        st.markdown(f'<div style="background:#4A90D9; color:white; padding:10px 18px; border-radius:8px; margin:22px 0 10px 0; font-weight:bold; font-size:19px;">📂 {current_cat}</div>', unsafe_allow_html=True)

                    existing = checks.get(key, {})
                    c1, c2, c3 = st.columns([5, 2, 3])
                    c1.markdown(f"**{item['no']}.** {item['item']}")
                    status = c2.radio(f"s{key}", ["미점검", "확인", "미비"], index=["미점검", "확인", "미비"].index(existing.get("status", "미점검")), horizontal=True, label_visibility="collapsed")
                    memo = c3.text_input(f"m{key}", value=existing.get("memo", ""), placeholder="비고", label_visibility="collapsed")
                    updated_checks[key] = {"status": status, "date": str(check_date) if status != "미점검" else existing.get("date", ""), "memo": memo}

                st.markdown("---")
                special_notes = st.text_area("▶ 현장 특이사항 기록", value=proj.get("special_notes", ""), height=120)

                if st.button("💾 구글 시트에 점검 결과 저장", type="primary", use_container_width=True):
                    proj["checks"] = updated_checks
                    proj["special_notes"] = special_notes
                    if "history" not in proj: proj["history"] = []
                    proj["history"].append({"date": str(check_date), "progress": int(calc_progress(updated_checks) * 100), "score": calc_score(updated_checks), "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M")})
                    save_to_sheets(st.session_state.projects)
                    st.success("✅ 구글 시트에 안전하게 저장되었습니다!")

                if proj.get("history"):
                    st.markdown("### 📈 점검 히스토리")
                    st.dataframe(pd.DataFrame(proj["history"])[["date", "progress", "score"]].rename(columns={"date":"점검일", "progress":"진척률(%)", "score":"점수"}), use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════
    # 📥 체크리스트 추출 (다중 선택 및 ZIP 일괄 다운로드 기능)
    # ═══════════════════════════════════════════════════
    elif menu == "📥 체크리스트 추출":
        st.markdown(f"""
            <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                {get_logo_html('34px')} 엑셀 양식 일괄 추출
            </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        
        if not projects: st.info("추출할 프로젝트가 없습니다.")
        else:
            # 1. 월별 조회 필터
            months = sorted(list(set(p['info']['delivery_date'][:7] for p in projects.values() if p['info']['delivery_date'])), reverse=True)
            col_filter, _ = st.columns([3, 7])
            selected_month = col_filter.selectbox("📅 조회 기간 선택 (년/월)", ["전체"] + months)
            
            filtered_projs = {pid: p for pid, p in projects.items() if selected_month == "전체" or p['info']['delivery_date'].startswith(selected_month)}
            
            if not filtered_projs:
                st.warning("해당 기간에 납품 예정인 프로젝트가 없습니다.")
            else:
                st.markdown("### ✅ 다운로드할 프로젝트를 선택하세요")
                
                # 2. 데이터 프레임을 이용한 직관적인 체크박스 선택 UI
                df_data = []
                for pid, p in filtered_projs.items():
                    df_data.append({
                        "선택": False,
                        "pid": pid,
                        "장비명": p['info']['equipment'],
                        "업체": p['info']['company'],
                        "납기일": p['info']['delivery_date'],
                        "진척률": f"{int(calc_progress(p['checks'])*100)}%"
                    })
                df = pd.DataFrame(df_data)
                
                # 사용자가 체크박스로 선택할 수 있도록 데이터 에디터 제공
                edited_df = st.data_editor(
                    df,
                    column_config={
                        "선택": st.column_config.CheckboxColumn("선택", default=False),
                        "pid": None  # pid 컬럼은 화면에서 숨김
                    },
                    disabled=["장비명", "업체", "납기일", "진척률"],
                    hide_index=True,
                    use_container_width=True
                )
                
                # 선택된 프로젝트 PID 추출
                selected_pids = edited_df[edited_df["선택"] == True]["pid"].tolist()
                
                if selected_pids:
                    st.success(f"총 {len(selected_pids)}개의 프로젝트가 선택되었습니다.")
                    
                    # 3. ZIP 파일 생성 로직
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for pid in selected_pids:
                            proj = projects[pid]
                            excel_data = generate_checklist_excel(proj)
                            filename = f"Frame_점검표_{proj['info']['company']}_{proj['info']['equipment']}.xlsx"
                            zip_file.writestr(filename, excel_data.getvalue())
                    
                    st.download_button(
                        label=f"📦 선택한 {len(selected_pids)}개 프로젝트 일괄 다운로드 (ZIP 압축파일)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Frame_점검표_일괄다운로드_{date.today().strftime('%Y%m%d')}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.info("위 표에서 다운로드할 프로젝트의 '선택' 칸을 체크해주세요.")