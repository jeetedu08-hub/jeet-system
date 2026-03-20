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
        total_analysis['영역'] = total_analysis['영역'].str.replace('문제해결력', '문제\n해결력')
        
        avg_cat_ratio = (total_analysis.groupby('영역')['평균득점'].sum() / total_analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
        unit_avg_data = total_analysis.groupby('단원').agg({'평균득점': 'sum'})
        unit_avg_data = unit_avg_data.reindex([u for u in unit_order if u in unit_avg_data.index])
  
        student_found = False

        for _, s_row in df_results.iterrows():
            student_name = str(s_row.get('이름', '')).strip()
            if not student_name or student_name == '0': continue
            if student_name != str(target_name).strip(): continue
                
            student_found = True
            student_grade = s_row.get('학년', '')
            
            analysis = df_info.copy()
            analysis['영역'] = analysis['영역'].str.replace('문제해결력', '문제\n해결력')
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

                if os.path.exists("logo.png"):
                    logo_img = plt.imread("logo.png")
                    logo_ax = fig.add_axes([0.80, 0.915, 0.15, 0.045], zorder=15)
                    logo_ax.imshow(logo_img)
                    logo_ax.axis('off')

                txt_jeet = fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                txt_title = fig.text(0.33, 0.88, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                
                info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}  |  과정: {selected_test}"
                txt_info = fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')

                txt_jeet.set_path_effects([path_effects.withStroke(linewidth=2, foreground=COLOR_RED)])
                txt_title.set_path_effects([path_effects.withStroke(linewidth=1.5, foreground=COLOR_NAVY)])
                txt_info.set_path_effects([path_effects.withStroke(linewidth=1, foreground='#222')])
  
                # --- 방사형 그래프 ---
                ax1 = fig.add_axes([0.10, 0.52, 0.32, 0.22], polar=True)
                all_cats = cat_ratio.index.tolist()
                ordered_labels = ['계산력'] + [c for c in all_cats if c != '계산력'] if '계산력' in all_cats else all_cats
                s_ordered = cat_ratio.reindex(ordered_labels)
                a_ordered = avg_cat_ratio.reindex(ordered_labels)
                labels = s_ordered.index.tolist()
                s_vals = s_ordered.values.tolist() + [s_ordered.values[0]]
                a_vals = a_ordered.values.tolist() + [a_ordered.values[0]]
                angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
                
                ax1.set_theta_direction(-1); ax1.set_theta_offset(np.pi/2.0)
                ax1.plot(angles, a_vals, color=COLOR_AVG, linewidth=1, linestyle='--', label='전체 평균')
                ax1.fill(angles, a_vals, color=COLOR_AVG, alpha=0.1)
                ax1.plot(angles, s_vals, color=COLOR_STUDENT, linewidth=2.5, label='학생 점수')
                ax1.set_ylim(0, 110); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([]); ax1.set_yticklabels([]) 
                
                for i in range(len(labels)):
                    angle = angles[i]; label_text = labels[i]
                    if angle == 0: ha, va, dist = 'center', 'bottom', 115
                    elif 0 < angle < np.pi: ha, va, dist = 'left', 'center', 110
                    elif angle == np.pi: ha, va, dist = 'center', 'top', 115
                    else: ha, va, dist = 'right', 'center', 110
                    
                    if '문제\n해결력' in label_text:
                        dist += 10
                        ha = 'left' if 0 < angle < np.pi else 'right'
                    
                    ax1.text(angle, dist, label_text, fontsize=10, fontweight='bold', va=va, ha=ha, color=COLOR_NAVY)
                    s_v, a_v = int(s_vals[i]), int(a_vals[i])
                    if '문제\n해결력' in label_text:
                        td = s_v - 15 if s_v > 30 else s_v + 15
                    else:
                        td = s_v + 10 if s_v < 85 else s_v - 18
                        
                    txt_s = ax1.text(angle, td, f"{s_v}%", fontsize=9, fontweight='bold', color=COLOR_STUDENT, va='center', ha='right')
                    txt_a = ax1.text(angle, td, f" ({a_v}%)", fontsize=9, fontweight='bold', color=COLOR_RED, va='center', ha='left')
                    for t in [txt_s, txt_a]: t.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
                
                ax1.legend(loc='upper right', bbox_to_anchor=(1.45, 1.15), fontsize=8, frameon=False)
                title1 = ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)
                title1.set_path_effects([path_effects.withStroke(linewidth=1, foreground=COLOR_NAVY)])
                fig.text(0.26, 0.49, "(파란색: 학생 성취율 / 빨간색: 전체 평균 성취율)", ha='center', fontsize=9, color='#555')
  
                # --- 단원별 성취도 그래프 (수정본 유지) ---
                ax2 = fig.add_axes([0.55, 0.54, 0.35, 0.18]) 
                x_pos = np.arange(len(unit_data))
                bar_width = 0.35 
                
                s_pct = (unit_data['득점'] / unit_data['배점'] * 100).fillna(0)
                a_pct = (unit_avg_data['평균득점'] / unit_data['배점'] * 100).fillna(0)
                
                ax2.bar(x_pos - bar_width/2, s_pct, color=COLOR_STUDENT, alpha=0.9, width=bar_width, zorder=3)
                ax2.bar(x_pos + bar_width/2, a_pct, color=COLOR_RED, alpha=0.8, width=bar_width, zorder=3)
                
                ax2.tick_params(axis='x', which='both', length=0) 
                ax2.set_xticks(x_pos)
                ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
                
                ax2.set_ylim(0, 110) 
                title2 = ax2.set_title("▶ 단원별 성취도 (%)", pad=25, fontsize=14, fontweight='bold', color=COLOR_NAVY)
                title2.set_path_effects([path_effects.withStroke(linewidth=1, foreground=COLOR_NAVY)])
                ax2.grid(axis='y', color=COLOR_GRID, linestyle='-', linewidth=0.5, zorder=0)
                
                fig.text(0.725, 0.485, "(파란색: 학생 성취율 / 빨간색: 전체 평균 성취율)", ha='center', fontsize=8.5, color='#555')
                
                for i in range(len(x_pos)):
                    s_val = int(s_pct.iloc[i])
                    a_val = int(a_pct.iloc[i])
                    s_y_pos = s_val - 3 if s_val > 15 else s_val + 3
                    s_color = 'white' if s_val > 15 else COLOR_STUDENT
                    t2 = ax2.text(x_pos[i] - bar_width/2, s_y_pos, f"{s_val}%", ha='center', va='top' if s_val > 15 else 'bottom', fontsize=7.5, fontweight='bold', color=s_color)
                    
                    a_y_pos = a_val - 3 if a_val > 15 else a_val + 3
                    a_color = 'white' if a_val > 15 else COLOR_RED
                    t3 = ax2.text(x_pos[i] + bar_width/2, a_y_pos, f"{a_val}%", ha='center', va='top' if a_val > 15 else 'bottom', fontsize=7.5, fontweight='bold', color=a_color)
                    
                    for t in [t2, t3]: 
                        if t.get_color() == 'white': t.set_path_effects([path_effects.withStroke(linewidth=1, foreground='#333')])
                        else: t.set_path_effects([path_effects.withStroke(linewidth=2, foreground='white')])
  
                # --- 4. 하단 심층 분석 박스 (요청 로직 반영) ---
                rect_diag = plt.Rectangle((0.08, 0.10), 0.84, 0.37, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                
                t_p1 = fig.text(0.11, 0.47, "▶ ", fontsize=14, fontweight='bold', color=COLOR_NAVY)
                t_p2 = fig.text(0.13, 0.47, " JEET", fontsize=14, fontweight='bold', color=COLOR_RED)
                t_p3 = fig.text(0.185, 0.47, f"   중등 수학 교육원 {student_name} 학생 심층 분석", fontsize=14, fontweight='bold', color=COLOR_NAVY)
                for t_obj in [t_p1, t_p2, t_p3]: t_obj.set_path_effects([path_effects.withStroke(linewidth=1, foreground=t_obj.get_color())])
                
                # --- 데이터 추출 (로직용) ---
                calc = cat_ratio.get('계산력', 0)
                solve = cat_ratio.get('문제\n해결력', 0)
                think = cat_ratio.get('이해력', 0)
                infer = cat_ratio.get('추론력', 0)
                
                u_res = (unit_data['득점'] / unit_data['배점'] * 100).fillna(0)
                best_unit = u_res.idxmax() if not u_res.empty else "종합"
                worst_unit = u_res.idxmin() if not u_res.empty else "종합"
                
                # 난이도별 분석용 데이터
                diff_data = analysis.groupby('난이도').agg({'득점':'sum', '배점':'sum'})
                concept = (diff_data.loc['개념', '득점'] / diff_data.loc['개념', '배점'] * 100) if '개념' in diff_data.index else 0
                apply = (diff_data.loc['응용', '득점'] / diff_data.loc['응용', '배점'] * 100) if '응용' in diff_data.index else 0
                deep = (diff_data.loc['심화', '득점'] / diff_data.loc['심화', '배점'] * 100) if '심화' in diff_data.index else 0

                # 1. 종합진단
                avg_val, total_avg_val = int(cat_ratio.mean()), int(avg_cat_ratio.mean())
                if avg_val >= 90: eval_tier = "심화 개념까지 완벽히 소화하는 탁월한 성취도"
                elif avg_val >= 75: eval_tier = "성실한 학습 태도가 돋보이는 우수한 성취도"
                elif avg_val >= 60: eval_tier = "개념을 정립하며 꾸준히 도약 중인 성취도"
                else: eval_tier = "기초를 다지며 가능성을 키워가는 단계의 성취도"
                
                # 2. 영역별 분석
                diag_area = ""
                if solve >= 75: diag_area += f"{student_name} 학생은 탁월한 문제해결 능력을 갖춘 학생입니다. 습득한 개념을 실전 문제에 효율적으로 투영하는 감각이 매우 훌륭합니다. "
                else: diag_area += f"개념의 실전 적용 단계에서 세심한 접근이 필요해 보입니다. 발문의 핵심 조건을 구조화하는 습관을 들인다면 큰 성장이 기대됩니다. "
                if calc >= 75: diag_area += f"기초 계산 숙련도가 안정적이어서 실수 없는 풀이가 가능합니다. "
                elif calc <= 40: diag_area += f"다만, 숙련도 영역인 계산력에서 다소 아쉬움이 관찰됩니다. 반복 연산 연습을 통해 계산 시간을 단축하고 정확도를 높이는 과정이 병행되어야 합니다. "
                if think >= 50: diag_area += f"사고력 지표는 안정적인 흐름을 보이며 어려운 문제에 도전할 수 있는 기본 체력을 보여주고 있습니다. "
                else: diag_area += f"사고력 영역은 현재 성장하는 단계에 있으며, 꾸준한 유형 분석을 통해 사고의 폭을 점진적으로 넓혀가는 것을 추천합니다. "
                if infer >= 75: diag_area += f"특히 추론 능력이 매우 뛰어나 수학적 규칙성을 발견하고 이를 논리적으로 확장하는 과정이 대단히 인상적입니다."
                else: diag_area += f"추론 영역은 다양한 유형의 도식화 연습을 통해 논리적 연결 고리를 찾아내는 훈련을 지속한다면 충분히 보완 가능합니다."

                # 3. 단원별 분석
                diag_unit = ""
                if u_res.get(best_unit) >= 75: diag_unit += f"'{best_unit}' 단원에서 보여준 고도의 집중력과 완벽한 성취도는 매우 고무적입니다. "
                if u_res.get(worst_unit) <= 45: diag_unit += f"상대적으로 '{worst_unit}' 단원은 개념의 계통성 있는 이해가 더 필요한 지점입니다. 지트(JEET) 정밀 분석 시스템을 통해 약점 단원의 오답 원인을 철저히 해소하겠습니다."
                else: diag_unit += f"전 단원에서 고른 학습 밸런스를 보여주고 있습니다."

                # 4. 난이도별 분석
                diag_diff = ""
                if concept >= 75 and apply >= 70: diag_diff += f"개념과 응용 수준의 성취도가 대단히 견고하여 수학적 기초 공사가 매우 잘 되어 있습니다. "
                elif concept <= 50: diag_diff += f"개념의 완전한 숙지가 선행된다면, 이미 보유한 응용 잠재력이 더욱 빛을 발할 수 있을 것으로 판단됩니다. "
                if deep >= 75: diag_diff += f"고난도 심화 문항까지 정밀하게 해결해낸 점은 {student_name} 학생이 가진 깊이 있는 사고 능력을 입증합니다."
                else: diag_diff += f"심화 문항에 대한 도전 경험을 쌓아가고 있으며, 숙련도가 보완된다면 향후 더 높은 고득점 진입이 확실시됩니다."

                # 5. 솔루션
                sol_text = f"단기적으로는 취약 유형 오답 노트를 작성하며 '{worst_unit}' 단원의 결손을 보완해야 합니다. JEET만의 맞춤 클리닉을 통해 상위권 도약을 위한 정밀 지도를 이어가겠습니다."

                # --- 텍스트 렌더링 ---
                sections = [
                    ("1. 종합 진단", f"{student_name} 학생은 전체 평균({total_avg_val}%) 대비 성취도 {avg_val}%를 기록하며, 현재 [{eval_tier}]를 보여주고 있습니다."),
                    ("[영역별 분석]", diag_area),
                    ("[단원별 분석]", diag_unit),
                    ("[난이도별 분석]", diag_diff),
                    ("[JEET 맞춤 솔루션]", sol_text)
                ]

                curr_y = 0.445
                for subtitle, content in sections:
                    stxt = fig.text(0.11, curr_y, subtitle, fontsize=9.5, fontweight='bold', color='#222')
                    stxt.set_path_effects([path_effects.withStroke(linewidth=0.5, foreground='#222')])
                    wrapped_content = textwrap.fill(content, width=65)
                    # 행간 조정을 위해 \n 기준으로 출력
                    fig.text(0.11, curr_y - 0.012, wrapped_content, fontsize=8.8, linespacing=1.4, va='top', color='#333')
                    curr_y -= 0.07 # 섹션 간격

                line_footer = plt.Line2D([0.05, 0.95], [0.09, 0.09], color=COLOR_NAVY, linewidth=1, transform=fig.transFigure)
                fig.lines.append(line_footer)
                campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜명로 10")]
                for i, (name, addr) in enumerate(campuses):
                    fig.text([0.22, 0.50, 0.78][i], 0.06, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
                    fig.text([0.22, 0.50, 0.78][i], 0.035, addr, ha='center', fontsize=7.5, color='#555')
                
                pdf.savefig(fig); plt.close(fig)
            
        if not student_found: return False, None, "학생을 찾을 수 없습니다."
        return True, pdf_buffer, "리포트 생성 완료!"
    except Exception as e: return False, None, f"오류 발생: {traceback.format_exc()}"

# --- 4. Streamlit 웹 UI 구성 ---
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide", page_icon="📊")
col1, col2 = st.columns([8, 2])
with col1: st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")
with col2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

try:
    doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
except Exception as e:
    st.error(f"구글 시트 로드 실패: {e}"); st.stop()

st.sidebar.header("📚 시험 과정 선택")
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)
df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]

tab1, tab2 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
    question_numbers = df_info_filtered['문항번호'].tolist()
    if question_numbers:
        with st.form("data_input_form", clear_on_submit=True):
            ci1, ci2, ci3 = st.columns(3)
            with ci1: input_name = st.text_input("이름")
            with ci2: input_school = st.text_input("학교")
            with ci3: input_grade = st.selectbox("학년", ["중1", "중2", "중3"])
            st.markdown("---")
            answers = {}
            for i in range(0, len(question_numbers), 5):
                cols = st.columns(5)
                for j, q_num in enumerate(question_numbers[i:i+5]):
                    with cols[j]:
                        choice = st.radio(f"**{q_num}번**", options=["O", "X"], horizontal=True, key=f"q_{q_num}")
                        answers[str(q_num)] = 1 if choice == "O" else 0
            if st.form_submit_button("구글 시트에 성적 저장하기", type="primary"):
                clean_name = input_name.strip()
                if not clean_name: st.error("⚠️ 이름을 입력해주세요.")
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
                        ws_results.append_row(new_row); st.success("성적이 저장되었습니다!"); st.cache_data.clear()
                    except Exception as e: st.error(f"저장 중 오류: {e}")

with tab2:
    st.subheader(f"[{selected_test}] 개별 심층 분석 리포트 생성")
    target_student = st.text_input("리포트를 출력할 학생 이름:", placeholder="예: 홍길동")
    if st.button("PDF 리포트 생성", type="primary"):
        with st.spinner("리포트 그리는 중..."):
            success, buf, msg = generate_jeet_expert_report(target_student.strip(), selected_test)
            if success:
                st.success(msg)
                st.download_button("📥 PDF 다운로드", buf.getvalue(), f"{target_student}_리포트.pdf", "application/pdf")
            else: st.error(msg)
