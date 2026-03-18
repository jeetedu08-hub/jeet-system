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

# --- 3. PDF 생성 함수 ---
def generate_jeet_expert_report(target_name, selected_test):
    try:
        _, _, _, df_info, df_results = load_data()
        
        df_info = df_info[df_info['시험명'] == selected_test]
        df_results = df_results[df_results['시험명'] == selected_test]
        
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

        for _, s_row in df_results.iterrows():
            student_name = str(s_row.get('이름', '')).strip()
            if not student_name or student_name == '0': continue
            
            if student_name != str(target_name).strip():
                continue
                
            student_found = True
            student_grade = s_row.get('학년', '')
            
            analysis = df_info.copy()
            analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
            analysis['득점'] = analysis['정답여부'] * analysis['배점']
            
            cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
            unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'})
            unit_data = unit_data.reindex([u for u in unit_order if u in unit_data.index])
  
            pdf_buffer = io.BytesIO()
            with PdfPages(pdf_buffer) as pdf:
                fig = plt.figure(figsize=(8.27, 11.69))
                
                border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure, zorder=10)
                fig.patches.append(border)

                # 🌟 [수정 핵심] 기존 설정을 유지하며 로고 y좌표만 위로 조정하여 제목과 분리
                if os.path.exists("logo.png"):
                    logo_img = plt.imread("logo.png")
                    # y좌표를 테두리 바로 아래(0.905)로 올리고 높이를 0.05로 줄여 공간 확보
                    logo_ax = fig.add_axes([0.80, 0.905, 0.15, 0.05], zorder=15)
                    logo_ax.imshow(logo_img)
                    logo_ax.axis('off')

                # 제목 텍스트는 기존 위치(0.88) 유지 -> 로고와 겹치지 않음
                fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.88, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                
                info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}  |  과정: {selected_test}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')
  
                ax1 = fig.add_axes([0.10, 0.52, 0.32, 0.22], polar=True)
                # ... (데이터 처리 및 차트 로직은 사용자님 코드와 동일)
                all_cats = cat_ratio.index.tolist()
                ordered_labels = ['계산력'] + [c for c in all_cats if c != '계산력'] if '계산력' in all_cats else all_cats
                s_ordered = cat_ratio.reindex(ordered_labels)
                a_ordered = avg_cat_ratio.reindex(ordered_labels)
                labels = s_ordered.index.tolist()
                s_vals = s_ordered.values.tolist() + [s_ordered.values[0]]
                a_vals = a_ordered.values.tolist() + [a_ordered.values[0]]
                angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
                
                ax1.set_theta_direction(-1) 
                ax1.set_theta_offset(np.pi/2.0)
                ax1.plot(angles, a_vals, color=COLOR_AVG, linewidth=1, linestyle='--', label='전체 평균')
                ax1.fill(angles, a_vals, color=COLOR_AVG, alpha=0.1)
                ax1.plot(angles, s_vals, color=COLOR_STUDENT, linewidth=2.5, label='학생 점수')
                ax1.set_ylim(0, 110) 
                ax1.set_xticks(angles[:-1])
                ax1.set_xticklabels([])
                ax1.set_yticklabels([]) 
  
                for i in range(len(labels)):
                    angle = angles[i]
                    label_text = labels[i]
                    if angle == 0: h_align = 'center'; v_align = 'bottom'; dist = 115
                    elif 0 < angle < np.pi: h_align = 'left'; v_align = 'center'; dist = 110
                    elif angle == np.pi: h_align = 'center'; v_align = 'top'; dist = 115
                    else: h_align = 'right'; v_align = 'center'; dist = 110
                    
                    ax1.text(angle, dist, label_text, fontsize=10, fontweight='bold', va=v_align, ha=h_align, color=COLOR_NAVY)
                    s_val = int(s_vals[i]); a_val = int(a_vals[i])
                    text_dist = s_val + 10 if s_val < 85 else s_val - 18
                    txt_s = ax1.text(angle, text_dist, f"{s_val}%", fontsize=9, fontweight='bold', color=COLOR_STUDENT, va='center', ha='right')
                    txt_a = ax1.text(angle, text_dist, f" ({a_val}%)", fontsize=9, fontweight='bold', color=COLOR_RED, va='center', ha='left')
                    for t in [txt_s, txt_a]: t.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
                
                ax1.legend(loc='upper right', bbox_to_anchor=(1.45, 1.15), fontsize=8, frameon=False)
                ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)
  
                ax2 = fig.add_axes([0.55, 0.52, 0.35, 0.20])
                x_pos = np.arange(len(unit_data))
                bars = ax2.bar(x_pos, unit_data['득점'], color=COLOR_STUDENT, alpha=0.8, width=0.5, zorder=3)
                ax2.scatter(x_pos, unit_avg_data['평균득점'], color=COLOR_RED, marker='_', s=1000, linewidth=3, zorder=4)
                ax2.set_xticks(x_pos)
                ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
                
                max_val = unit_data['배점'].max()
                max_val = 10 if pd.isna(max_val) or max_val == 0 else max_val
                ax2.set_ylim(0, max_val * 1.5)
                ax2.set_title("▶ 단원별 성취도", pad=15, fontsize=14, fontweight='bold', color=COLOR_NAVY)
                ax2.grid(axis='y', color=COLOR_GRID, linestyle='-', linewidth=0.5, zorder=0)
  
                for i, bar in enumerate(bars):
                    s_v = int(bar.get_height()); a_v = int(unit_avg_data['평균득점'].iloc[i])
                    ax2.text(bar.get_x() + bar.get_width()/2, s_v + 0.5, f"{s_v}", ha='right', va='bottom', fontsize=9, fontweight='bold', color=COLOR_STUDENT)
                    ax2.text(bar.get_x() + bar.get_width()/2, s_v + 0.5, f" ({a_v})", ha='left', va='bottom', fontsize=9, fontweight='bold', color=COLOR_RED)
  
                rect_diag = plt.Rectangle((0.08, 0.15), 0.84, 0.32, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                fig.text(0.11, 0.44, "▶ ", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                fig.text(0.13, 0.44, " JEET", fontsize=15, fontweight='bold', color=COLOR_RED)
                fig.text(0.185, 0.44, f"   중등 수학 교육원 {student_name} 학생 심층 분석", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                
                avg_val, total_avg_val = int(cat_ratio.mean()), int(avg_cat_ratio.mean())
                diff_val = avg_val - total_avg_val
                diff_cats = s_ordered - a_ordered
                best_cat = diff_cats.idxmax() if not diff_cats.empty else "종합"
                worst_cat = diff_cats.idxmin() if not diff_cats.empty else "종합"
                unit_diff = unit_data['득점'] - unit_avg_data['평균득점']
                worst_unit = unit_diff.idxmin() if not unit_diff.empty else "전반적인"
                
                if avg_val >= 90: eval_tier = "최상위권 수준의 탁월한 성취도"
                elif avg_val >= 75: eval_tier = "기본기가 탄탄한 우수한 성취도"
                elif avg_val >= 60: eval_tier = "핵심 개념에 대한 보완이 필요한 성취도"
                else: eval_tier = "전반적인 기초 학습 재점검이 요구되는 성취도"
                
                solution_dict = {
                    '계산력': "매일 일정한 분량의 연산 훈련을 병행하여 잦은 실수를 줄이고 풀이 속도를 높이는 연습을 강력히 권장합니다.",
                    '이해력': "개념서의 기본 원리와 공식 유도 과정을 스스로 백지에 설명해보는 '백지 복습법'을 통해 뼈대를 튼튼히 해야 합니다.",
                    '추론력': "조건이 복잡한 심화 문제를 단계별로 끊어 읽고, 출제자의 숨은 의도를 분석 및 도식화하는 훈련이 필요합니다.",
                    '문제해결력': "다양한 개념이 통합된 융합형 문제 위주로 다루며, 스스로 식을 세우고 끝까지 답을 도출하는 끈기를 길러야 합니다."
                }
                worst_solution = solution_dict.get(worst_cat, "개인별 맞춤 클리닉을 통해 취약 단원을 집중 공략할 것을 권장합니다.")

                diag_content = (
                    f"1. 종합 진단: {student_name} 학생은 전체 평균({total_avg_val}%) 대비 {abs(diff_val)}%p {'높은' if diff_val >= 0 else '낮은'} {avg_val}%의 성적을 기록하여 [{eval_tier}]를 보이고 있습니다.\n\n"
                    f"2. 강약점 분석: 영역별 진단 결과 '{best_cat}' 역량이 가장 돋보이나, 상대적으로 '{worst_cat}' 역량의 보완이 시급합니다. 특히 단원별 성취도에서 '{worst_unit}' 파트의 오답률이 높아 해당 부분의 개념 재점검이 우선되어야 합니다.\n\n"
                    f"3. JEET 맞춤 솔루션: 단기적으로는 '{worst_unit}' 단원의 오답 노트를 작성하고 유사 유형을 반복 훈련해야 합니다. 중장기적으로 '{worst_cat}'을(를) 끌어올리기 위해 {worst_solution}"
                )
                
                wrapped_lines = [textwrap.fill(p, width=54) for p in diag_content.split('\n\n')]
                fig.text(0.11, 0.41, "\n\n".join(wrapped_lines), fontsize=10.5, linespacing=1.8, va='top', ha='left', color='#333')
  
                line_footer = plt.Line2D([0.05, 0.95], [0.12, 0.12], color=COLOR_NAVY, linewidth=1, transform=fig.transFigure)
                fig.lines.append(line_footer)
                campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜명로 10")]
                for i, (name, addr) in enumerate(campuses):
                    fig.text([0.22, 0.50, 0.78][i], 0.08, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
                    fig.text([0.22, 0.50, 0.78][i], 0.05, addr, ha='center', fontsize=7.5, color='#555')
  
                pdf.savefig(fig)
                plt.close(fig)
            
        if not student_found:
            return False, None, f"선택한 과정({selected_test})에서 '{target_name}' 학생을 찾을 수 없습니다."
            
        return True, pdf_buffer, f"'{target_name}' 학생의 [{selected_test}] 심층 분석 리포트가 성공적으로 생성되었습니다!"
  
    except Exception as e:
        error_msg = traceback.format_exc()
        return False, None, f"오류가 발생했습니다:\n{error_msg}"

# ==========================================
# --- 4. Streamlit 웹 UI 구성 ---
# ==========================================
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide", page_icon="📊")

# (웹 상단 UI 및 로고 표시)
col1, col2 = st.columns([8, 2])
with col1:
    st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")
with col2:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_column_width=False, width=150)

# (구글 시트 로드)
try:
    doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
except Exception as e:
    st.error(f"구글 시트를 불러오는 데 실패했습니다. 에러 내용: {e}")
    st.stop()

# (사이드바 시험 과정 선택)
st.sidebar.header("📚 시험 과정 선택")
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)

df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]
st.sidebar.success(f"✅ 현재 [ {selected_test} ] 모드입니다.")

# (탭 구성)
tab1, tab2 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력"])

# --- [tab1: 기존 O/X 선택형 입력 방식 완벽 유지] ---
with tab1:
    st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
    st.info("문항별로 O(정답) 또는 X(오답)를 클릭해주세요.")
    
    question_numbers = df_info_filtered['문항번호'].tolist()

    if question_numbers:
        with st.form("data_input_form", clear_on_submit=True):
            ci1, ci2, ci3 = st.columns(3)
            with ci1: input_name = st.text_input("이름")
            with ci2: input_school = st.text_input("학교")
            with ci3: input_grade = st.selectbox("학년", ["중1", "중2", "중3"])
                
            st.markdown("---")
            st.write("**문항별 정오표 체크**")
            
            # 5열 그리드로 O/X 라디오 버튼 배치 (기존 설정 유지)
            cols = st.columns(5)
            answers = {}
            for i, q_num in enumerate(question_numbers):
                with cols[i % 5]:
                    choice = st.radio(
                        f"**{q_num}번**",
                        options=["O", "X"],
                        horizontal=True,
                        key=f"q_{q_num}"
                    )
                    answers[str(q_num)] = 1 if choice == "O" else 0
            
            st.markdown("---")
            submit_btn = st.form_submit_button("구글 시트에 성적 저장하기", type="primary")
            
            if submit_btn:
                # (구글 시트 저장 로직 - 기존 설정 유지)
                clean_name = input_name.strip()
                if not clean_name:
                    st.error("⚠️ 학생 이름을 입력해주세요.")
                else:
                    try:
                        header_row = ws_results.row_values(1)
                        new_row = []
                        for col_name in header_row:
                            col_str = str(col_name)
                            if col_str == '시험명': new_row.append(selected_test) 
                            elif col_str == '이름': new_row.append(clean_name)
                            elif col_str == '학교': new_row.append(input_school)
                            elif col_str == '학년': new_row.append(input_grade)
                            elif col_str in answers: new_row.append(answers[col_str])
                            else: new_row.append("")
                        
                        ws_results.append_row(new_row)
                        st.success(f"✅ 저장되었습니다!")
                        st.balloons()
                        st.cache_data.clear() # 캐시 초기화
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")

# --- [tab2: 기존 리포트 출력 방식 유지] ---
with tab2:
    st.subheader(f"[{selected_test}] 개별 심층 분석 리포트 생성")
    target_student = st.text_input("리포트를 출력할 학생 이름:", placeholder="예: 홍길동")
    
    if st.button("PDF 리포트 생성", type="primary"):
        clean_target_name = target_student.strip()
        if not clean_target_name:
            st.warning("⚠️ 학생 이름을 먼저 입력해주세요.")
        else:
            with st.spinner(f"[{selected_test}] 데이터를 분석하고 리포트를 그리는 중입니다..."):
                success, pdf_buffer, message = generate_jeet_expert_report(clean_target_name, selected_test)
                
                if success:
                    st.success(message)
                    st.download_button(
                        label="📥 PDF 다운로드",
                        data=pdf_buffer.getvalue(),
                        file_name=f"{clean_target_name}_{selected_test}_리포트.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.error(message)
