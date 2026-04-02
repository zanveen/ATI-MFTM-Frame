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
    return "<span style='color:#4A90D9; font-weight:900;'>ATI</span> "

# ─── 구글 시트 연동 및 복구 로직 ───
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(credentials)

def load_from_sheets():
    try:
        client = get_gspread_client()
        sh = client.open("Frame_Data")
        worksheet = sh.sheet1 
        all_records = worksheet.get_all_records()
        
        projects = {}
        for row in all_records:
            pid = str(row.get('pid', '')).strip()
            if not pid: continue 
            
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
                    "delay_request": json.loads(row.get('delay_request', '{}')) if row.get('delay_request') else {},
                    "is_delivered": bool(row.get('is_delivered', False))
                },
                "checks": json.loads(row.get('checks', '{}')) if row.get('checks') else {},
                "special_notes": str(row.get('special_notes', '')),
                "history": json.loads(row.get('history', '[]')) if row.get('history') else []
            }
        return projects
    except Exception as e:
        return {}

def save_to_sheets(projects):
    try:
        client = get_gspread_client()
        sh = client.open("Frame_Data")
        worksheet = sh.sheet1
        
        header = ["pid", "company", "equipment", "order_date", "delivery_date", "frame_parts", 
                  "frame_options", "exterior_spec", "interior_spec", "notes_top", 
                  "delivery_delay_count", "delay_request", "checks", "special_notes", "history", "is_delivered"]
        
        data_to_save = [header]
        for pid, p in projects.items():
            if not p.get('info'): continue 
            info = p['info']
            data_to_save.append([
                pid, info.get('company',''), info.get('equipment',''), info.get('order_date',''), info.get('delivery_date',''),
                info.get('frame_parts', 1), json.dumps(info.get('frame_options', []), ensure_ascii=False),
                info.get('exterior_spec',''), info.get('interior_spec',''), info.get('notes_top',''),
                info.get('delivery_delay_count', 0), json.dumps(info.get('delay_request', {}), ensure_ascii=False),
                json.dumps(p.get('checks', {}), ensure_ascii=False), p.get('special_notes',''), 
                json.dumps(p.get('history', []), ensure_ascii=False), info.get('is_delivered', False)
            ])
        
        worksheet.clear()
        try:
            worksheet.update(values=data_to_save, range_name='A1')
        except TypeError: 
            worksheet.update('A1', data_to_save)
            
    except Exception as e:
        st.error(f"저장 실패: {e}")

# ─── 계산 및 유틸 함수 ───
def repair_project(p):
    if not isinstance(p, dict): p = {}
    if "info" not in p: p["info"] = {}
    info = p["info"]
    info.setdefault("company", "미상")
    info.setdefault("equipment", "알수없음")
    info.setdefault("delivery_date", "")
    info.setdefault("delivery_delay_count", 0)
    info.setdefault("frame_parts", 1)
    info.setdefault("frame_options", [])
    info.setdefault("is_delivered", False)
    if "delay_request" not in info or not isinstance(info["delay_request"], dict):
        info["delay_request"] = {}
    if "checks" not in p: p["checks"] = {}
    if "history" not in p: p["history"] = []
    if "special_notes" not in p: p["special_notes"] = ""
    return p

def calc_progress(checks):
    total_score = sum(5 if c.get("status") == "확인" else 2 if c.get("status") == "미비" else 0 for c in checks.values())
    return total_score / 100

def calc_score(checks):
    return sum(5 if c.get("status") == "확인" else 2 if c.get("status") == "미비" else 0 for c in checks.values())

def get_project_status(proj):
    if proj.get("info", {}).get("is_delivered", False):
        return "완료"
    
    dd = proj.get("info", {}).get("delivery_date", "")
    if dd:
        try:
            d = datetime.strptime(dd, "%Y-%m-%d").date()
            if (d - date.today()).days <= 7: return "납기임박"
        except: pass
    return "진행중"

def filter_projects_by_role(all_projects, user_info):
    if user_info["role"] == "admin": return all_projects
    return {pid: p for pid, p in all_projects.items() if p.get("info", {}).get("company") == user_info["company"]}

def generate_checklist_excel(project):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb["점검표"]
    info = project.get("info", {})
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

# ─── CSS 스타일 ───
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

    /* 대시보드 납기임박 빨간색 고정 CSS */
    div[data-testid="column"]:nth-child(4) button p,
    div[data-testid="column"]:nth-of-type(4) button p,
    div[data-testid="stHorizontalBlock"] > div:nth-child(4) button p,
    div[data-testid="stHorizontalBlock"] > div:last-child button p { color: #e74c3c !important; }
    
    div[data-testid="column"]:nth-child(4) button p::first-line,
    div[data-testid="column"]:nth-of-type(4) button p::first-line,
    div[data-testid="stHorizontalBlock"] > div:nth-child(4) button p::first-line,
    div[data-testid="stHorizontalBlock"] > div:last-child button p::first-line { color: #e74c3c !important; }
    
    div[data-testid="column"]:nth-child(4) button,
    div[data-testid="column"]:nth-of-type(4) button { border-color: #fca5a5 !important; }
    
    div[data-testid="column"]:nth-child(4) button:hover,
    div[data-testid="column"]:nth-of-type(4) button:hover { border-color: #e74c3c !important; background-color: #fef2f2 !important; }
</style>
""", unsafe_allow_html=True)

# ─── 세션 초기화 및 시트 로드 ───
if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "user_info" not in st.session_state: st.session_state.user_info = None
if "dashboard_filter" not in st.session_state: st.session_state.dashboard_filter = "전체"
if "inspection_project" not in st.session_state: st.session_state.inspection_project = None

# 🔥 액션 완료 후 띄워줄 알림창을 저장하는 변수 (화면 새로고침 대응)
if "flash_msg" not in st.session_state: st.session_state.flash_msg = None

if "projects" not in st.session_state:
    st.session_state.projects = load_from_sheets()

# ═══════════════════════════════════════════════════
# 1. 로그인
# ═══════════════════════════════════════════════════
if not st.session_state.logged_in:
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown(f"""
        <div class="login-container">
            <div style="font-size: 28px; font-weight: 800; color: #1a1a2e; margin-bottom: 20px;">
                {get_logo_html('34px')} FRAME 제작 통합 관리
            </div>
            <p style="color:#666; margin-bottom:30px;">비밀번호를 입력해주세요.</p>
        """, unsafe_allow_html=True)
        pwd = st.text_input("비밀번호", type="password", label_visibility="collapsed")
        st.markdown("<br>", unsafe_allow_html=True)
        login_col1, login_col2, login_col3 = st.columns([1, 2, 1])
        with login_col2:
            if st.button("로그인", use_container_width=True, type="primary"):
                if pwd in USER_CREDENTIALS:
                    st.session_state.logged_in = True
                    st.session_state.user_info = USER_CREDENTIALS[pwd]
                    st.session_state.projects = load_from_sheets() 
                    st.rerun()
                else:
                    st.error("비밀번호가 일치하지 않습니다.")
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════
# 2. 메인 애플리케이션
# ═══════════════════════════════════════════════════
else:
    user = st.session_state.user_info

    # ─── 사이드바 ───
    st.sidebar.markdown(f"""
        <div style="font-size:30px; font-weight:bold; margin-bottom:20px; line-height:1.2;">
            {get_logo_html('32px')} FRAME 진척률 관리
        </div>
    """, unsafe_allow_html=True)
        
    st.sidebar.markdown(f"**환영합니다. {user['name']}님**")
    st.sidebar.markdown("---")

    menu_options = ["ALL", "캘린더", "점검", "양식 추출"]
    if user["role"] == "admin":
        menu_options.insert(1, "신규 등록")

    menu = st.sidebar.radio("메뉴", menu_options, label_visibility="collapsed")
    
    # Handle navigation from board/calendar to inspection
    nav = st.query_params.get("nav", "")
    if nav == "inspect" and st.session_state.inspection_project:
        menu = "점검"
        st.query_params.clear()
    
    if menu != "점검": st.session_state.inspection_project = None

    st.sidebar.markdown("---")
    if st.sidebar.button("로그아웃", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.user_info = None
        st.rerun()
        
    projects = filter_projects_by_role(st.session_state.projects, user)

    # 📌 글로벌 플래시 메시지 렌더링 (새로고침되어도 최상단에 알림 팝업 생성!)
    if st.session_state.flash_msg:
        st.toast(st.session_state.flash_msg, icon="✅")
        st.success(st.session_state.flash_msg)
        if "등록" in st.session_state.flash_msg:
            st.balloons()
        st.session_state.flash_msg = None  # 한 번 보여주고 삭제

    # ═══════════════════════════════════════════════════
    # 프로젝트 보드
    # ═══════════════════════════════════════════════════
    if menu == "ALL":
        col_title, col_stats = st.columns([5, 5])
        with col_title:
            st.markdown(f"""
                <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px;">
                    {get_logo_html('34px')} {'전체 진행 현황' if user['role']=='admin' else '우리 회사 배정 프로젝트'}
                </div>
            """, unsafe_allow_html=True)
            
        with col_stats:
            if user["role"] == "admin":
                pending_reqs = [p for p in projects.values() if p.get("info", {}).get("delay_request", {}).get("status") == "pending"]
                if pending_reqs:
                    st.markdown(f"""
                    <div style="background-color:#fff3cd; border-left:5px solid #ffc107; padding:12px; border-radius:6px; margin-bottom:15px;">
                        <h4 style="margin:0; color:#856404; font-size:18px;">[안내] 납기 변경 요청 접수 ({len(pending_reqs)}건)</h4>
                        <p style="margin:4px 0 0 0; font-size:14px; color:#856404;">업체에서 납기일 조율을 요청했습니다. '점검' 메뉴에서 승인해주세요.</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                months = set(p.get("info", {}).get("delivery_date", "")[:7] for p in st.session_state.projects.values() if p.get("info", {}).get("delivery_date"))
                month_list = sorted(list(months), reverse=True)
                selected_period = st.selectbox("통계 조회 기간", ["전체 누적"] + month_list)
                
                h_delays, h_projs, j_delays, j_projs = 0, 0, 0, 0
                for p in st.session_state.projects.values():
                    info = p.get("info", {})
                    dd = info.get("delivery_date", "")
                    if selected_period != "전체 누적" and not dd.startswith(selected_period): continue
                    comp, delay = info.get("company", ""), info.get("delivery_delay_count", 0)
                    if comp == "한울산업": h_delays += delay; h_projs += 1
                    elif comp == "정한테크": j_delays += delay; j_projs += 1

                st.markdown(f"""
                <div style="background-color:#f8f9fa; border:1px solid #dee2e6; padding:15px; border-radius:8px;">
                    <b style="color:#2c3e50; font-size:16px;">{selected_period} 업체별 현황</b>
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
                sort_option = st.selectbox("정렬 기준", ["납품 예정일순", "진척률 낮은순", "진척률 높은순"], label_visibility="collapsed")
                sorted_pids = show_pids.copy()
                if sort_option == "납품 예정일순": sorted_pids.sort(key=lambda x: projects[x].get("info", {}).get("delivery_date", "9999"))
                elif sort_option == "진척률 낮은순": sorted_pids.sort(key=lambda x: calc_progress(projects[x].get("checks", {})))
                else: sorted_pids.sort(key=lambda x: calc_progress(projects[x].get("checks", {})), reverse=True)

                for pid in sorted_pids:
                    proj = projects.get(pid)
                    if not proj: continue
                    
                    info = proj.get("info", {})
                    checks = proj.get("checks", {})
                    pct = int(calc_progress(checks) * 100)
                    score = calc_score(checks)
                    delay_cnt = info.get("delivery_delay_count", 0)
                    is_deliv = info.get("is_delivered", False)

                    if is_deliv: badge, bar_color = '<span class="status-badge badge-green">[완료] 납품완료</span>', "#95a5a6"
                    elif pct >= 50: badge, bar_color = '<span class="status-badge badge-blue">[진행] 조립중</span>', "#4A90D9"
                    elif pct > 0: badge, bar_color = '<span class="status-badge badge-yellow">[초기] 제작중</span>', "#f39c12"
                    else: badge, bar_color = '<span class="status-badge badge-red">[대기] 미착수</span>', "#e74c3c"

                    dd = info.get("delivery_date", "")
                    days_text = ""
                    if dd and not is_deliv:
                        try:
                            diff = (datetime.strptime(dd, "%Y-%m-%d").date() - date.today()).days
                            if diff < 0: days_text = f'<span style="color:#e74c3c;font-weight:bold;">납기 {abs(diff)}일 초과</span>'
                            elif diff == 0: days_text = '<span style="color:#e74c3c;font-weight:bold;">오늘 납품</span>'
                            elif diff <= 7: days_text = f'<span style="color:#e74c3c;font-weight:bold;">D-{diff}</span>'
                            else: days_text = f'<span style="color:#888;">D-{diff}</span>'
                        except: pass

                    opts_html = f'<span style="background:#eff6ff; color:#3b82f6; padding:4px 12px; border-radius:20px; font-size:14px; font-weight:700; border:1px solid #bfdbfe; margin-left:15px;">옵션: {", ".join(info.get("frame_options", []))}</span>' if info.get("frame_options") else ''
                    delay_text = f" &nbsp;│ <span style='color:#e74c3c;font-weight:bold;'>납기변경 {delay_cnt}회</span>" if delay_cnt > 0 else ""

                    card_html = (
                        f'<div class="project-card">'
                        f'<div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; margin-bottom:10px;">'
                        f'<div style="display:flex; align-items:center;"><h3 style="margin:0;">{info.get("equipment", pid)}</h3>{opts_html}</div>'
                        f'<div style="display:flex; align-items:center; gap:10px;">{badge}<span>{days_text}</span></div>'
                        f'</div>'
                        f'<p style="margin-top:5px; color:#475569;">업체: <b style="color:#1e293b;">{info.get("company","-")}</b> │ 납품예정: {dd} │ 검사점수: <b>{score}/100</b> {delay_text}</p>'
                        f'<div class="progress-bar-bg"><div class="progress-bar-fill" style="width:{max(pct,3)}%;background:{bar_color};">{pct}%</div></div>'
                        f'</div>'
                    )
                    st.markdown(card_html, unsafe_allow_html=True)
                    
                    if not is_deliv:
                        col1, col2, col3 = st.columns([6, 2, 2])
                        with col2:
                            if st.button("📋 점검 보기", key=f"btn_inspect_{pid}", use_container_width=True):
                                st.session_state.inspection_project = pid
                                # Switch to inspection menu
                                st.query_params["nav"] = "inspect"
                                st.rerun()
                        with col3:
                            if user["role"] == "admin":
                                if st.button("납품 완료 처리", key=f"btn_done_{pid}", use_container_width=True):
                                    st.session_state.projects[pid]["info"]["is_delivered"] = True
                                    save_to_sheets(st.session_state.projects)
                                    st.session_state.flash_msg = f"✅ [{info.get('equipment')}] 납품 처리가 완료되었습니다!"
                                    st.rerun()
                    else:
                        if user["role"] == "admin":
                            pass  # Already delivered, no actions

    # ═══════════════════════════════════════════════════
    # 납기 캘린더
    # ═══════════════════════════════════════════════════
    elif menu == "캘린더":
        st.markdown(f"""
            <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                {get_logo_html('34px')} 프로젝트 납기 캘린더
            </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        today = date.today()
        col_y, col_m, _ = st.columns([2, 2, 6])
        year = col_y.selectbox("년도 선택", range(today.year - 1, today.year + 3), index=1)
        month = col_m.selectbox("월 선택", range(1, 13), index=today.month - 1)

        # 해당 월 프로젝트를 pid 포함 dict로 수집
        month_str = f"{year}-{month:02d}"
        month_proj_map = {}  # date_str -> [(pid, proj), ...]
        for pid, p in projects.items():
            dd = p.get('info', {}).get('delivery_date', '')
            if dd.startswith(month_str):
                if dd not in month_proj_map:
                    month_proj_map[dd] = []
                month_proj_map[dd].append((pid, p))

        cal = calendar.monthcalendar(year, month)
        html_cal = '<table style="width:100%; border-collapse:collapse; background:white; box-shadow:0 4px 6px rgba(0,0,0,0.05); border-radius:10px; overflow:hidden; table-layout:fixed;">'
        
        html_cal += '<tr>'
        for day_name in ["월", "화", "수", "목", "금", "토", "일"]:
            color = "#e74c3c" if day_name == "일" else "#3b82f6" if day_name == "토" else "#333"
            html_cal += f'<th style="border:1px solid #e2e8f0; padding:12px; text-align:center; background-color:#f8fafc; font-size:18px; color:{color};">{day_name}</th>'
        html_cal += '</tr>'
        
        for week in cal:
            html_cal += '<tr>'
            for idx, day in enumerate(week):
                if day == 0:
                    html_cal += '<td style="border:1px solid #e2e8f0; background-color:#f1f5f9; height:110px;"></td>'
                else:
                    date_str = f"{year}-{month:02d}-{day:02d}"
                    is_today = (date_str == today.strftime("%Y-%m-%d"))
                    bg_color = "#eff6ff" if is_today else "white"
                    day_color = "#e74c3c" if idx == 6 else "#3b82f6" if idx == 5 else "#333"
                    weight = "900" if is_today else "bold"
                    
                    day_projs = month_proj_map.get(date_str, [])
                    
                    cell_content = f'<div style="font-weight:{weight}; font-size:16px; margin-bottom:6px; color:{day_color};">{day}{" (오늘)" if is_today else ""}</div>'
                    
                    for pid_item, p in day_projs:
                        info = p.get("info", {})
                        is_deliv = info.get("is_delivered", False)
                        
                        try: diff = (datetime.strptime(date_str, "%Y-%m-%d").date() - today).days
                        except: diff = 99
                        
                        if is_deliv: badge_color = "#95a5a6"
                        elif diff < 0: badge_color = "#e74c3c"
                        elif diff <= 7: badge_color = "#e74c3c"
                        else: badge_color = "#3b82f6"
                        
                        cell_content += f'<div style="background:{badge_color}; color:white; padding:3px 6px; border-radius:4px; margin-bottom:3px; font-size:12px; font-weight:bold; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{info.get("equipment")}</div>'
                    
                    html_cal += f'<td style="border:1px solid #e2e8f0; padding:8px; height:110px; vertical-align:top; background-color:{bg_color};">{cell_content}</td>'
            html_cal += '</tr>'
        html_cal += '</table>'
        
        st.markdown(html_cal, unsafe_allow_html=True)

        # 프로젝트 선택 팝업: selectbox로 프로젝트 선택 → expander로 상세 팝업
        all_month_projs = [(pid, p) for projs in month_proj_map.values() for pid, p in projs]
        
        if all_month_projs:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("#### 📌 프로젝트 상세 보기")
            
            proj_options = {pid: f"{p.get('info',{}).get('equipment','')} ({p.get('info',{}).get('company','')}) - 납기: {p.get('info',{}).get('delivery_date','')}" for pid, p in all_month_projs}
            selected_cal_pid = st.selectbox("프로젝트 선택", list(proj_options.keys()), format_func=lambda x: proj_options[x], label_visibility="collapsed")
            
            if selected_cal_pid:
                sp = projects.get(selected_cal_pid, {})
                sp_info = sp.get("info", {})
                sp_checks = sp.get("checks", {})
                sp_pct = int(calc_progress(sp_checks) * 100)
                sp_score = calc_score(sp_checks)
                sp_deliv = sp_info.get("is_delivered", False)
                sp_dd = sp_info.get("delivery_date", "")
                
                try: sp_diff = (datetime.strptime(sp_dd, "%Y-%m-%d").date() - today).days
                except: sp_diff = 99
                
                if sp_deliv: sp_status = "✅ 납품완료"
                elif sp_diff < 0: sp_status = f"⚠️ 납기 {abs(sp_diff)}일 초과"
                elif sp_diff <= 7: sp_status = f"⏰ D-{sp_diff}"
                else: sp_status = f"D-{sp_diff}"
                
                opts_text = f"옵션: {', '.join(sp_info.get('frame_options', []))}" if sp_info.get('frame_options') else ""
                
                with st.expander(f"📋 {sp_info.get('equipment')} 상세 정보", expanded=True):
                    st.markdown(f"""
                    <div style="background:#f8f9fa; padding:16px; border-radius:10px; line-height:2;">
                        <b style="font-size:20px;">🔩 {sp_info.get('equipment')}</b><br>
                        <b>업체:</b> {sp_info.get('company')} &nbsp;│&nbsp;
                        <b>납품예정:</b> {sp_dd} &nbsp;│&nbsp;
                        <b>상태:</b> {sp_status}<br>
                        <b>진척률:</b> {sp_pct}% &nbsp;│&nbsp;
                        <b>검사점수:</b> {sp_score}/100<br>
                        {'<b>' + opts_text + '</b><br>' if opts_text else ''}
                        <b>외관:</b> {sp_info.get('exterior_spec','')} &nbsp;│&nbsp;
                        <b>내부:</b> {sp_info.get('interior_spec','')} &nbsp;│&nbsp;
                        <b>파트:</b> {sp_info.get('frame_parts','')}덩어리
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # 대분류별 진척
                    cat_progress = {}
                    for item in CHECKLIST_ITEMS:
                        cat = item["category"]
                        if cat not in cat_progress: cat_progress[cat] = {"total": 0, "done": 0}
                        cat_progress[cat]["total"] += 1
                        key = str(item["no"])
                        if key in sp_checks and sp_checks[key].get("status") == "확인":
                            cat_progress[cat]["done"] += 1
                    
                    cat_cols = st.columns(len(cat_progress))
                    for ci, (cat, data) in enumerate(cat_progress.items()):
                        with cat_cols[ci]:
                            cp = int(data["done"]/data["total"]*100) if data["total"]>0 else 0
                            st.metric(cat, f"{data['done']}/{data['total']}", f"{cp}%")
                    
                    if not sp_deliv:
                        st.markdown("")
                        btn_col1, btn_col2 = st.columns(2)
                        with btn_col1:
                            if st.button("📋 점검 페이지로 이동", key=f"cal_go_inspect_{selected_cal_pid}", use_container_width=True):
                                st.session_state.inspection_project = selected_cal_pid
                                st.query_params["nav"] = "inspect"
                                st.rerun()
                        with btn_col2:
                            if user["role"] == "admin":
                                if st.button("✅ 납품 완료 처리", key=f"cal_deliver_{selected_cal_pid}", use_container_width=True, type="primary"):
                                    st.session_state.projects[selected_cal_pid]["info"]["is_delivered"] = True
                                    save_to_sheets(st.session_state.projects)
                                    st.session_state.flash_msg = f"✅ [{sp_info.get('equipment')}] 납품 처리 완료!"
                                    st.rerun()

    # ═══════════════════════════════════════════════════
    # 프로젝트 등록 (중복 방지 & 팝업 시스템 완벽 탑재!)
    # ═══════════════════════════════════════════════════
    elif menu == "신규 등록" and user["role"] == "admin":
        st.markdown(f"""
            <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                {get_logo_html('34px')} 신규 프로젝트 등록
            </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        # 🔥 입력 폼: clear_on_submit=True 를 넣어서 제출 시 입력칸 자동 초기화!
        with st.form("new_project_form", clear_on_submit=True):
            company = st.radio("업체명 *", ["한울산업", "정한테크"], horizontal=True)
            
            c1, c2 = st.columns(2)
            equipment = c1.text_input("장비명 (설비명) *", placeholder="장비명을 입력하세요")
            order_date = c2.date_input("발주일")
            
            c3, c4 = st.columns(2)
            frame_parts = c3.number_input("Frame Part 수 (덩어리) *", min_value=1, value=1, step=1)
            delivery_date = c4.date_input("납품 예정일자")

            st.markdown("<br><b>프레임 옵션</b>", unsafe_allow_html=True)
            opt_cols = st.columns(3)
            with opt_cols[0]: opt_clean = st.checkbox("클린부스")
            with opt_cols[1]: opt_table = st.checkbox("테이블")
            with opt_cols[2]: opt_jig = st.checkbox("전도방지지그")
            frame_options = [opt for opt, checked in zip(["클린부스", "테이블", "전도방지지그"], [opt_clean, opt_table, opt_jig]) if checked]

            st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)
            
            spec_col1, spec_col2 = st.columns(2)
            with spec_col1: exterior_spec = st.radio("외관 사양", ["SUS", "도장"], horizontal=True)
            with spec_col2: interior_spec = st.radio("내부 사양", ["SUS", "도장"], horizontal=True)

            notes_top = st.text_area("특이사항", placeholder="추가 전달 사항을 입력하세요", height=80)
            
            st.markdown("<br>", unsafe_allow_html=True)
            submitted = st.form_submit_button("구글 시트에 프로젝트 등록", use_container_width=True)

            if submitted:
                if not equipment: st.error("장비명은 필수 입력입니다.")
                else:
                    try:
                        pid = f"{company}_{equipment}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                        st.session_state.projects[pid] = repair_project({
                            "info": {
                                "company": company, "equipment": equipment, "order_date": str(order_date),
                                "delivery_date": str(delivery_date), "frame_parts": frame_parts,
                                "frame_options": frame_options, "exterior_spec": exterior_spec,
                                "interior_spec": interior_spec, "notes_top": notes_top,
                                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                "delivery_delay_count": 0, "delay_request": {}, "is_delivered": False
                            },
                            "checks": {}, "special_notes": "", "history": []
                        })
                        save_to_sheets(st.session_state.projects)
                        # 저장 후 플래시 메시지 등록하고 새로고침! (그러면 화면 상단에 팝업처럼 나타남)
                        st.session_state.flash_msg = "✅ 신규 프로젝트가 성공적으로 등록되었습니다!"
                        st.rerun()
                    except Exception as e:
                        st.error(f"등록 중 오류 발생: {e}")

    # ═══════════════════════════════════════════════════
    # 점검
    # ═══════════════════════════════════════════════════
    elif menu == "점검":
        if not projects: st.info("점검할 프로젝트가 없습니다.")
        else:
            active_projects = {pid: p for pid, p in projects.items() if not p.get("info", {}).get("is_delivered", False)}

            if st.session_state.inspection_project is None:
                st.markdown(f"""
                    <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                        {get_logo_html('34px')} 진행중인 프로젝트 점검
                    </div>
                """, unsafe_allow_html=True)
                
                if not active_projects:
                    st.success("모든 프로젝트가 납품 완료되었습니다!")
                
                for comp_name in ["한울산업", "정한테크"]:
                    comp_projs = [(k, v) for k, v in active_projects.items() if v.get("info", {}).get("company") == comp_name]
                    if comp_projs:
                        st.markdown(f'<div style="background:#f8f9fa; border-radius:12px; padding:16px 20px 6px 20px; margin-bottom:8px;"><div style="font-size:22px; font-weight:bold; color:#2c3e50; padding-bottom:8px; border-bottom:3px solid #4A90D9;">{comp_name}</div></div>', unsafe_allow_html=True)
                        for pid, proj in comp_projs:
                            pct = int(calc_progress(proj.get("checks", {})) * 100)
                            dd = proj.get("info", {}).get("delivery_date", "")
                            bar_c = ('#27ae60' if pct>=80 else '#4A90D9' if pct>=50 else '#f39c12' if pct>0 else '#e74c3c')
                            
                            col1, col2 = st.columns([8, 2])
                            with col1:
                                kanban_html = (
                                    f'<div style="background:white; border-radius:10px; padding:16px; margin-bottom:10px; border-left:5px solid #4A90D9; box-shadow:0 1px 6px rgba(0,0,0,0.07);">'
                                    f'<b style="font-size:18px;">{proj["info"]["equipment"]}</b> (납기: {dd})'
                                    f'<div style="background:#e8ecf1;border-radius:6px;height:10px;margin-top:8px;">'
                                    f'<div style="background:{bar_c};height:100%;width:{max(pct,2)}%;border-radius:6px;"></div>'
                                    f'</div></div>'
                                )
                                st.markdown(kanban_html, unsafe_allow_html=True)
                            with col2:
                                st.markdown("<br>", unsafe_allow_html=True)
                                if st.button("점검 시작", key=f"ins_{pid}", use_container_width=True):
                                    st.session_state.inspection_project = pid
                                    st.rerun()
            else:
                pid = st.session_state.inspection_project
                proj = projects.get(pid)
                if not proj:
                    st.session_state.inspection_project = None
                    st.rerun()

                info = proj.get("info", {})
                checks = proj.get("checks", {})
                
                if st.button("프로젝트 목록으로 돌아가기"):
                    st.session_state.inspection_project = None
                    st.rerun()

                delay_req = info.get("delay_request", {})
                st.markdown("---")
                st.markdown("### 일정 관리 및 조율")
                
                if user["role"] == "vendor":
                    if delay_req and delay_req.get("status") == "pending":
                        st.info(f"관리자에게 납기 변경을 요청했습니다. (희망일: {delay_req.get('requested_date')} / 승인 대기중)")
                    else:
                        with st.expander("관리자에게 납기 변경 요청하기"):
                            with st.form(f"delay_form_{pid}", clear_on_submit=True):
                                new_date = st.date_input("변경 희망 납품일")
                                reason = st.text_input("납기 변경(지연) 사유")
                                if st.form_submit_button("요청 전송", use_container_width=True):
                                    proj["info"]["delay_request"] = {"requested_date": str(new_date), "reason": reason, "status": "pending"}
                                    save_to_sheets(st.session_state.projects)
                                    st.session_state.flash_msg = "납기 변경 요청이 전송되었습니다."
                                    st.rerun()

                if user["role"] == "admin":
                    if delay_req and delay_req.get("status") == "pending":
                        st.warning(f"**[{info.get('company')}] 납기 변경 요청** : 기존 {info.get('delivery_date')} -> **희망 {delay_req.get('requested_date')}** (사유: {delay_req.get('reason')})")
                        col_a, col_b, col_c = st.columns([2, 2, 6])
                        if col_a.button("요청 승인 (납기 적용)", use_container_width=True):
                            proj["info"]["delivery_date"] = delay_req["requested_date"]
                            proj["info"]["delivery_delay_count"] = info.get("delivery_delay_count", 0) + 1
                            proj["info"]["delay_request"]["status"] = "approved"
                            save_to_sheets(st.session_state.projects)
                            st.session_state.flash_msg = "✅ 납기 변경 요청이 승인되었습니다."
                            st.rerun()
                        if col_b.button("요청 반려", use_container_width=True):
                            proj["info"]["delay_request"]["status"] = "rejected"
                            save_to_sheets(st.session_state.projects)
                            st.session_state.flash_msg = "납기 변경 요청이 반려되었습니다."
                            st.rerun()
                    
                    with st.expander("관리자 권한: 납기일 임의 변경 (지연 누적 적용)"):
                        with st.form(f"admin_delay_form_{pid}", clear_on_submit=True):
                            try: default_d = datetime.strptime(info.get("delivery_date", ""), "%Y-%m-%d").date()
                            except: default_d = date.today()
                            
                            new_date_admin = st.date_input("새로운 납기일로 변경", value=default_d)
                            st.warning("납기일을 변경하면 해당 업체의 '납기 지연 횟수'가 1회 증가합니다.")
                            confirm_change = st.checkbox("위 내용을 확인하였으며 납기일을 변경합니다.")
                            
                            if st.form_submit_button("납기일 변경 적용", use_container_width=True):
                                if str(new_date_admin) == info.get("delivery_date", ""):
                                    st.error("기존 납기일과 동일합니다. 다른 날짜를 선택하세요.")
                                elif not confirm_change:
                                    st.error("변경 사항을 적용하시려면 확인 체크박스를 선택해주세요.")
                                else:
                                    proj["info"]["delivery_date"] = str(new_date_admin)
                                    proj["info"]["delivery_delay_count"] = info.get("delivery_delay_count", 0) + 1
                                    save_to_sheets(st.session_state.projects)
                                    st.session_state.flash_msg = "✅ 구글 시트에 납기일이 변경되었습니다!"
                                    st.rerun()

                st.markdown("---")
                opts_html = f'<span style="background:#eff6ff; color:#3b82f6; padding:4px 12px; border-radius:20px; font-size:14px; font-weight:700; border:1px solid #bfdbfe; margin-left:15px;">옵션: {", ".join(info.get("frame_options", []))}</span>' if info.get("frame_options") else ''

                st.markdown(f"""
                <div style="background:#eef2f7;padding:16px 22px;border-radius:10px;margin-bottom:24px;">
                    <div style="display:flex; align-items:center; margin-bottom:8px;">
                        <h3 style="margin:0;font-size:23px;">{info.get('equipment')}</h3>{opts_html}
                    </div>
                    <p style="margin:4px 0;font-size:17px;">
                        <b>업체:</b> {info.get('company')} │ <b>납품일:</b> <span style="color:#e74c3c;font-weight:bold;">{info.get('delivery_date')}</span> (납기변경 누적 {info.get('delivery_delay_count', 0)}회) │
                        <b>진척률:</b> {int(calc_progress(checks)*100)}%
                    </p>
                </div>
                """, unsafe_allow_html=True)

                check_date = st.date_input("당일 점검 일자 지정", value=date.today(), key="cd")
                updated_checks = copy.deepcopy(checks)
                current_cat = ""

                for item in CHECKLIST_ITEMS:
                    key = str(item["no"])
                    if item["category"] != current_cat:
                        current_cat = item["category"]
                        st.markdown(f'<div style="background:#4A90D9; color:white; padding:10px 18px; border-radius:8px; margin:22px 0 10px 0; font-weight:bold; font-size:19px;">{current_cat}</div>', unsafe_allow_html=True)

                    existing = checks.get(key, {})
                    c1, c2, c3 = st.columns([5, 2, 3])
                    c1.markdown(f"**{item['no']}.** {item['item']}")
                    status = c2.radio(f"s{key}", ["미점검", "확인", "미비"], index=["미점검", "확인", "미비"].index(existing.get("status", "미점검")), horizontal=True, label_visibility="collapsed")
                    memo = c3.text_input(f"m{key}", value=existing.get("memo", ""), placeholder="비고 입력", label_visibility="collapsed")
                    updated_checks[key] = {"status": status, "date": str(check_date) if status != "미점검" else existing.get("date", ""), "memo": memo}

                st.markdown("---")
                special_notes = st.text_area("현장 특이사항 기록", value=proj.get("special_notes", ""), height=120)

                if st.button("점검 결과 저장", type="primary", use_container_width=True):
                    proj["checks"] = updated_checks
                    proj["special_notes"] = special_notes
                    if "history" not in proj: proj["history"] = []
                    proj["history"].append({"date": str(check_date), "progress": int(calc_progress(updated_checks) * 100), "score": calc_score(updated_checks), "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M")})
                    save_to_sheets(st.session_state.projects)
                    st.success("✅ 구글 시트에 점검 결과가 저장되었습니다.")

                if proj.get("history"):
                    st.markdown("### 점검 히스토리")
                    st.dataframe(pd.DataFrame(proj["history"])[["date", "progress", "score"]].rename(columns={"date":"점검일", "progress":"진척률(%)", "score":"점수"}), use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════
    # 체크리스트 추출
    # ═══════════════════════════════════════════════════
    elif menu == "양식 추출":
        st.markdown(f"""
            <div style="font-size:34px; font-weight:800; color:#1e293b; margin-top:10px; margin-bottom:20px;">
                {get_logo_html('34px')} 엑셀 양식 일괄 추출
            </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        
        if not projects: st.info("추출할 프로젝트가 없습니다.")
        else:
            months = sorted(list(set(p.get('info', {}).get('delivery_date', '')[:7] for p in projects.values() if p.get('info', {}).get('delivery_date', ''))), reverse=True)
            col_filter, _ = st.columns([3, 7])
            selected_month = col_filter.selectbox("조회 기간 선택 (년/월)", ["전체"] + months)
            
            filtered_projs = {pid: p for pid, p in projects.items() if selected_month == "전체" or p.get('info', {}).get('delivery_date', '').startswith(selected_month)}
            
            if not filtered_projs:
                st.warning("해당 기간에 해당하는 프로젝트가 없습니다.")
            else:
                st.markdown("### 추출할 프로젝트를 선택하세요")
                
                df_data = []
                for pid, p in filtered_projs.items():
                    info = p.get('info', {})
                    df_data.append({
                        "선택": False,
                        "pid": pid,
                        "장비명": info.get('equipment', ''),
                        "업체": info.get('company', ''),
                        "납기일": info.get('delivery_date', ''),
                        "진척률": f"{int(calc_progress(p.get('checks', {}))*100)}%"
                    })
                df = pd.DataFrame(df_data)
                
                edited_df = st.data_editor(
                    df,
                    column_config={
                        "선택": st.column_config.CheckboxColumn("선택", default=False),
                        "pid": None
                    },
                    disabled=["장비명", "업체", "납기일", "진척률"],
                    hide_index=True,
                    use_container_width=True
                )
                
                selected_pids = edited_df[edited_df["선택"] == True]["pid"].tolist()
                
                if selected_pids:
                    today_str = date.today().strftime('%Y%m%d')
                    
                    st.success(f"총 {len(selected_pids)}개 선택됨")
                    st.markdown("")
                    
                    if len(selected_pids) == 1:
                        # 1개: 개별 엑셀 파일 다운로드
                        proj = projects.get(selected_pids[0])
                        if proj:
                            excel_data = generate_checklist_excel(proj)
                            filename = f"Frame_점검표_{proj.get('info', {}).get('company', '')}_{proj.get('info', {}).get('equipment', '')}_{today_str}.xlsx"
                            st.download_button(
                                label=f"📥 {proj.get('info', {}).get('equipment', '')} 점검표 다운로드",
                                data=excel_data.getvalue(),
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                                key="dl_single",
                                use_container_width=True
                            )
                    else:
                        # 2개 이상: ZIP 일괄 다운로드
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for pid in selected_pids:
                                proj = projects.get(pid)
                                if not proj: continue
                                excel_data = generate_checklist_excel(proj)
                                filename = f"Frame_점검표_{proj.get('info', {}).get('company', '')}_{proj.get('info', {}).get('equipment', '')}_{today_str}.xlsx"
                                zf.writestr(filename, excel_data.getvalue())
                        zip_buffer.seek(0)
                        
                        st.download_button(
                            label=f"📥 {len(selected_pids)}개 점검표 일괄 다운로드 (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Frame_점검표_일괄_{today_str}.zip",
                            mime="application/zip",
                            key="dl_zip_all",
                            use_container_width=True
                        )
                else:
                    st.info("위 표에서 다운로드할 프로젝트의 '선택' 칸을 체크해주세요.")
