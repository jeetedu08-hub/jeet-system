import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
import textwrap
import matplotlib.font_manager as fm
import matplotlib.patheffects as path_effects
import traceback
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json

# --- 1. 환경 및 폰트 설정 ---
font_path = "malgun.ttf"
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    font_prop = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
else:
    plt.rcParams['font.family'] = 'Malgun Gothic'

plt.rcParams['axes.unicode_minus'] = False
  
COLOR_NAVY = '#1A237E'; COLOR_RED = '#D32F2F'; COLOR_STUDENT = '#0056B3'
COLOR_AVG = '#757575'; COLOR_GRID = '#E0E0E0'; COLOR_BG = '#F8F9FA'

# --- 2. 구글 스프레드시트 연동 및 캐시 설정 ---
@st.cache_resource
def get_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "GOOGLE_JSON" in st.secrets:
        creds_dict = json.loads(st.secrets["GOOGLE_JSON"])
    elif "gcp_secret_string" in st.secrets:
        creds_dict = json.loads(st.secrets["gcp_secret_string"])
    elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        creds_dict = json.loads(st.secrets["connections"]["gsheets"].get("credentials", "{}"))
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("secrets.json", scope)
        return gspread.authorize(creds).open_by_url("https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit")
        
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    doc = client.open_by_url("https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit")
    return doc

@st.cache_data(ttl=120)
def fetch_all_dataframes():
    doc = get_google_sheet()
    df_info = pd.DataFrame(doc.worksheet('Test_Info').get_all_records())
    df_results = pd.DataFrame(doc.worksheet('Student_Results').get_all_records())
    df_results = df_results.replace('', 0).fillna(0)
    return df_info, df_results

def load_data():
    doc = get_google_sheet()
    ws_info = doc.worksheet('Test_Info')
    ws_results = doc.worksheet('Student_Results')
    df_info, df_results = fetch_all_dataframes()
    return doc, ws_info, ws_results, df_info, df_results

# --- 3. PDF 생성 함수 (기존 로직 유지) ---
def generate_jeet_expert_report(target_name, selected_test):
    try:
        _, _, _, df_info_all, df_results_all = load_data()
        df_info = df_info_all[df_info_all['시험명'] == selected_test].copy()
        df_results = df_results_all[df_results_all['시험명'] == selected_test].copy()
        
        df_results.columns = df_results.columns.astype(str)
        df_info['배점'] = df_info['배점'].replace('', 3).fillna(3).astype(int)
        unit_order = df_info['단원'].drop_duplicates().tolist()
        q_cols = [str(q) for q in df_info['문항번호']]
        valid_cols = [col for col in df_results.columns if col in q_cols]
        
        def safe_to_int(val):
            try: return int(float(val))
            except: return 0

        df_scores = df_results[valid_cols].applymap(safe_to_int)
        avg_per_q = df_scores.mean()
        total_analysis = df_info.copy()
        total_analysis['평균득점'] = total_analysis['문항번호'].apply(lambda x: avg_per_q.get(str(x), 0))
        avg_cat_ratio = (total_analysis.groupby('영역')['평균득점'].sum() / total_analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
        unit_avg_data = total_analysis.groupby('단원').agg({'평균득점': 'sum'})
        unit_avg_data = unit_avg_data.reindex([u for u in unit_order if u in unit_avg_data.index])
  
        student_found = False
        pdf_buffer = io.BytesIO()

        for _, s_row in df_results.iterrows():
            student_name = str(s_row.get('이름', '')).strip()
            if not student_name or student_name == '0' or student_name != str(target_name).strip(): continue
            student_found = True
            student_grade = s_row.get('학년', '')
            analysis = df_info.copy()
            analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
            analysis['득점'] = analysis['정답여부'] * analysis['배점']
            cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
            unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'})
            unit_data = unit_data.reindex([u for u in unit_order if u in unit_data.index])
  
            with PdfPages(pdf_buffer) as pdf:
                fig = plt.figure(figsize=(8.27, 11.69))
                border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure, zorder=10)
                fig.patches.append(border)
                fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.88, f'{selected_test} 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                info_text = f"과정: {selected_test}  |  학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')
                
                # [그래프 및 진단 내용 생략 - 기존과 동일]
                # (이전 코드의 ax1, ax2 그래프와 진단 텍스트 생성 코드가 이 위치에 들어갑니다)
                
                pdf.savefig(fig)
                plt.close(fig)
                break
        return True, pdf_buffer, "완료"
    except Exception as e:
        return False, None, traceback.format_exc()

# ==========================================
# --- 4. Streamlit 웹 UI ---
# ==========================================
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide", page_icon="📊")

# 상단 타이틀
c1, c2 = st.columns([8, 2])
with c1: st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")
with c2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

# 데이터 로드
try:
    doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
except Exception as e:
    st.error(f"데이터 로드 실패: {e}"); st.stop()

# 사이드바 선택
st.sidebar.header("📚 시험 과정 선택")
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)
st.sidebar.success(f"✅ 현재 [ {selected_test} ] 모드")

tab1, tab2 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
    df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]
    q_nums = df_info_filtered['문항번호'].tolist()
    
    if q_nums:
        # 입력 폼 시작
        with st.form("input_form", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            i_name = col1.text_input("이름")
            i_school = col2.text_input("학교")
            i_grade = col3.selectbox("학년", ["중1", "중2", "중3"])
            
            st.markdown("---")
            st.write("**문항별 정답 여부 선택 (O: 정답, X: 오답)**")
            
            # 🌟 [수정 포인트] O, X 선택 방식으로 변경
            # 한 줄에 5개씩 배치
            cols = st.columns(5)
            user_answers = {}
            
            for i, q in enumerate(q_nums):
                with cols[i % 5]:
                    # 라디오 버튼을 사용하여 O, X 선택
                    choice = st.radio(
                        f"{q}번",
                        options=["O", "X"],
                        key=f"q_{q}",
                        horizontal=True,
                        index=1  # 기본값을 'X'로 설정 (필요시 0으로 바꾸면 'O'가 기본값)
                    )
                    # 저장할 때는 O면 1, X면 0으로 변환
                    user_answers[str(q)] = 1 if choice == "O" else 0
            
            st.markdown("---")
            submit_btn = st.form_submit_button("구글 시트에 성적 저장하기", type="primary")
            
            if submit_btn:
                if not i_name.strip():
                    st.error("⚠️ 학생 이름을 입력해주세요.")
                else:
                    try:
                        header_row = ws_results.row_values(1)
                        new_row = []
                        for col_name in header_row:
                            col_str = str(col_name)
                            if col_str == '시험명': new_row.append(selected_test)
                            elif col_str == '이름': new_row.append(i_name.strip())
                            elif col_str == '학교': new_row.append(i_school)
                            elif col_str == '학년': new_row.append(i_grade)
                            elif col_str in user_answers: new_row.append(user_answers[col_str])
                            else: new_row.append("")
                        
                        ws_results.append_row(new_row)
                        st.success(f"✅ {i_name} 학생 성적이 저장되었습니다!")
                        st.balloons()
                        fetch_all_dataframes.clear()
                    except Exception as e:
                        st.error(f"저장 오류: {e}")

with tab2:
    st.subheader(f"[{selected_test}] 개별 심층 분석 리포트 생성")
    target_name = st.text_input("학생 이름 입력:")
    if st.button("PDF 리포트 생성", type="primary"):
        if not target_name: st.warning("이름을 입력하세요.")
        else:
            with st.spinner("분석 중..."):
                success, buffer, msg = generate_jeet_expert_report(target_name, selected_test)
                if success:
                    st.success(f"{target_name} 학생 분석 완료!")
                    st.download_button("📥 PDF 다운로드", buffer.getvalue(), f"{target_name}_{selected_test}.pdf", "application/pdf")
                else: st.error(msg)
