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

# --- 1. 환경 및 폰트 설정 (클라우드 호환) ---
font_path = "malgun.ttf"
if os.path.exists(font_path):
    font_prop = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
else:
    plt.rcParams['font.family'] = 'Malgun Gothic'

plt.rcParams['axes.unicode_minus'] = False
  
COLOR_NAVY = '#1A237E'; COLOR_RED = '#D32F2F'; COLOR_STUDENT = '#0056B3'
COLOR_AVG = '#757575'; COLOR_GRID = '#E0E0E0'; COLOR_BG = '#F8F9FA'

import json

# --- 2. 구글 스프레드시트 연동 함수 (클라우드 비밀 금고 지원) ---
@st.cache_resource
def get_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    # 클라우드 비밀 금고에 열쇠가 있으면 사용하고, 없으면 내 컴퓨터의 열쇠 사용
    if "gcp_secret_string" in st.secrets:
        creds_dict = json.loads(st.secrets["gcp_secret_string"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("secrets.json", scope)
        
    client = gspread.authorize(creds)
    doc = client.open("성적표")
    return doc

# --- 3. PDF 생성 함수 (전문가 분석 + 겹침 완벽 방지 + 간격 최적화) ---
def generate_jeet_expert_report(target_name):
    try:
        _, _, _, df_info, df_results = load_data()
        
        # 구글 시트의 모든 열 이름을 강제로 문자로 통일
        df_results.columns = df_results.columns.astype(str)
        
        df_info['배점'] = df_info['배점'].replace('', 3).fillna(3).astype(int)
        unit_order = ['유리수와 순환소수', '식의 계산', '일차부등식', '연립방정식', '일차함수']
  
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
  
            # 클라우드 다운로드를 위한 메모리 버퍼
            pdf_buffer = io.BytesIO()
            with PdfPages(pdf_buffer) as pdf:
                fig = plt.figure(figsize=(8.27, 11.69))
                
                # 배경 및 타이틀
                border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure, zorder=10)
                fig.patches.append(border)
                fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.88, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')
  
                # --- [차트 1] 방사형 차트 ---
                ax1 = fig.add_axes([0.10, 0.52, 0.32, 0.22], polar=True)
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
  
                # 🚨 [거리 조절 완벽 적용] 라벨 및 수치 자동 정렬
                for i in range(len(labels)):
                    angle = angles[i]
                    label_text = labels[i]
                    
                    if angle == 0:  
                        h_align = 'center'; v_align = 'bottom'; dist = 115
                    elif 0 < angle < np.pi:  
                        h_align = 'left'; v_align = 'center'; dist = 110
                    elif angle == np.pi:  
                        h_align = 'center'; v_align = 'top'; dist = 115
                    else:  
                        h_align = 'right'; v_align = 'center'; dist = 110
                    
                    ax1.text(angle, dist, label_text, fontsize=10, fontweight='bold', va=v_align, ha=h_align, color=COLOR_NAVY)
                    
                    s_val = int(s_vals[i]); a_val = int(a_vals[i])
                    text_dist = s_val + 10 if s_val < 85 else s_val - 18
                    
                    txt_s = ax1.text(angle, text_dist, f"{s_val}%", fontsize=9, fontweight='bold', color=COLOR_STUDENT, va='center', ha='right')
                    txt_a = ax1.text(angle, text_dist, f" ({a_val}%)", fontsize=9, fontweight='bold', color=COLOR_RED, va='center', ha='left')
                    for t in [txt_s, txt_a]: t.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
                
                ax1.legend(loc='upper right', bbox_to_anchor=(1.45, 1.15), fontsize=8, frameon=False)
                ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)
  
                # --- [차트 2] 바 차트 ---
                ax2 = fig.add_axes([0.55, 0.52, 0.35, 0.20])
                x_pos = np.arange(len(unit_data))
                bars = ax2.bar(x_pos, unit_data['득점'], color=COLOR_STUDENT, alpha=0.8, width=0.5, zorder=3)
                ax2.scatter(x_pos, unit_avg_data['평균득점'], color=COLOR_RED, marker='_', s=1000, linewidth=3, zorder=4)
                ax2.set_xticks(x_pos)
                ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
                ax2.set_ylim(0, unit_data['배점'].max() * 1.5)
                ax2.set_title("▶ 단원별 성취도", pad=15, fontsize=14, fontweight='bold', color=COLOR_NAVY)
                ax2.grid(axis='y', color=COLOR_GRID, linestyle='-', linewidth=0.5, zorder=0)
  
                for i, bar in enumerate(bars):
                    s_v = int(bar.get_height()); a_v = int(unit_avg_data['평균득점'].iloc[i])
                    ax2.text(bar.get_x() + bar.get_width()/2, s_v + 0.5, f"{s_v}", ha='right', va='bottom', fontsize=9, fontweight='bold', color=COLOR_STUDENT)
                    ax2.text(bar.get_x() + bar.get_width()/2, s_v + 0.5, f" ({a_v})", ha='left', va='bottom', fontsize=9, fontweight='bold', color=COLOR_RED)
  
                # --- [분석 섹션: 전문가 수준 자동 분석] ---
                rect_diag = plt.Rectangle((0.08, 0.15), 0.84, 0.32, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                fig.text(0.11, 0.44, "▶ ", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                fig.text(0.13, 0.44, " JEET", fontsize=15, fontweight='bold', color=COLOR_RED)
                fig.text(0.185, 0.44, f"   중등 수학 교육원 {student_name} 학생 심층 분석", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                
                avg_val, total_avg_val = int(cat_ratio.mean()), int(avg_cat_ratio.mean())
                diff_val = avg_val - total_avg_val
                
                # 강/약점 자동 추출
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
  
                # --- [푸터] ---
                line_footer = plt.Line2D([0.05, 0.95], [0.12, 0.12], color=COLOR_NAVY, linewidth=1, transform=fig.transFigure)
                fig.lines.append(line_footer)
                campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜명로 10")]
                for i, (name, addr) in enumerate(campuses):
                    fig.text([0.22, 0.50, 0.78][i], 0.08, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
                    fig.text([0.22, 0.50, 0.78][i], 0.05, addr, ha='center', fontsize=7.5, color='#555')
  
                pdf.savefig(fig)
                plt.close(fig)
            
        if not student_found:
            return False, None, f"구글 시트에서 '{target_name}' 학생을 찾을 수 없습니다."
            
        return True, pdf_buffer, f"'{target_name}' 학생의 심층 분석 리포트가 성공적으로 생성되었습니다!"
  
    except Exception as e:
        error_msg = traceback.format_exc()
        return False, None, f"오류가 발생했습니다:\n{error_msg}"

# ==========================================
# --- 4. Streamlit 웹 UI 구성 ---
# ==========================================
st.set_page_config(page_title="JEET 통합 관리 시스템 (Cloud)", layout="wide", page_icon="☁️")
st.title("☁️ JEET 통합 관리 시스템 (Google 연동형)")

tab1, tab2 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력"])

with tab1:
    st.subheader("신규 학생 성적 입력")
    st.info("입력하신 데이터는 구글 스프레드시트에 실시간으로 저장됩니다.")
    
    try:
        _, _, ws_results, df_info, df_results = load_data()
        question_numbers = df_info['문항번호'].tolist()
    except Exception as e:
        st.error(f"구글 시트를 불러오는 데 실패했습니다. 열쇠 파일(secrets.json)과 공유 설정을 확인하세요.\n에러 내용: {e}")
        question_numbers = []

    if question_numbers:
        with st.form("data_input_form", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1: input_name = st.text_input("이름")
            with col2: input_school = st.text_input("학교")
            with col3: input_grade = st.selectbox("학년", ["중1", "중2", "중3"])
                
            st.markdown("---")
            st.write("**문항별 정답 입력 (1: 정답, 0: 오답)**")
            
            cols = st.columns(5)
            answers = {}
            for i, q_num in enumerate(question_numbers):
                with cols[i % 5]:
                    answers[str(q_num)] = st.number_input(f"{q_num}번 문항", min_value=0, max_value=1, step=1)
                    
            submit_btn = st.form_submit_button("구글 시트에 성적 저장하기", type="primary")
            
            if submit_btn:
                clean_name = input_name.strip()
                if not clean_name:
                    st.error("⚠️ 학생 이름을 입력해주세요.")
                else:
                    try:
                        header_row = ws_results.row_values(1)
                        new_row = []
                        for col_name in header_row:
                            col_str = str(col_name)
                            if col_str == '이름': new_row.append(clean_name)
                            elif col_str == '학교': new_row.append(input_school)
                            elif col_str == '학년': new_row.append(input_grade)
                            elif col_str in answers: new_row.append(answers[col_str])
                            else: new_row.append("")
                        
                        ws_results.append_row(new_row)
                        st.success(f"✅ '{clean_name}' 학생의 성적이 구글 시트에 실시간으로 저장되었습니다!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"저장 중 오류가 발생했습니다: {e}")

with tab2:
    st.subheader("개별 심층 분석 리포트 생성")
    target_student = st.text_input("리포트를 출력할 학생 이름:", placeholder="예: 홍길동")
    
    if st.button("PDF 리포트 생성", type="primary"):
        clean_target_name = target_student.strip()
        if not clean_target_name:
            st.warning("⚠️ 학생 이름을 먼저 입력해주세요.")
        else:
            with st.spinner("구글 시트에서 데이터를 분석하고 리포트를 그리는 중입니다..."):
                success, pdf_buffer, message = generate_jeet_expert_report(clean_target_name)
                
                if success:
                    st.success(message)
                    st.download_button(
                        label="📥 PDF 리포트 파일 다운로드",
                        data=pdf_buffer.getvalue(),
                        file_name=f"{clean_target_name}_JEET_심층분석리포트.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.error(message)
