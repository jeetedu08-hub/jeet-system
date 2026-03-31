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
import io
import json
from supabase import create_client, Client

# --- 1. 환경 및 폰트 설정 ---
font_path = "malgun.ttf"
try:
    if os.path.exists(font_path):
        fm.fontManager.addfont(font_path)
        font_prop = fm.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    else:
        plt.rcParams['font.family'] = 'Malgun Gothic'
except Exception as e:
    plt.rcParams['font.family'] = 'sans-serif'

plt.rcParams['axes.unicode_minus'] = False

COLOR_NAVY = '#1A237E'
COLOR_RED = '#D32F2F'
COLOR_STUDENT = '#0056B3' 
COLOR_UNIT = '#00796B'    
COLOR_AVG = '#757575'
COLOR_GRID = '#E0E0E0'
COLOR_BG = '#F8F9FA'

# --- 2. Supabase 연동 설정 ---
@st.cache_resource
def get_supabase_client() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

# ★ 수정된 부분: 이 함수가 누락되어 에러가 났었습니다. ★
@st.cache_data(ttl=120)
def fetch_all_dataframes():
    supabase = get_supabase_client()
    
    info_res = supabase.table("Test_Info").select("*").execute()
    df_info = pd.DataFrame(info_res.data)
    
    results_res = supabase.table("Student_Results").select("*").execute()
    df_results = pd.DataFrame(results_res.data)
    
    if not df_results.empty:
        # 이름이 없거나 '0.0', '0'인 유령 데이터 필터링
        df_results['이름'] = df_results['이름'].astype(str).str.strip()
        df_results = df_results[
            (df_results['이름'] != '') & 
            (df_results['이름'] != '0') & 
            (df_results['이름'] != '0.0') & 
            (df_results['이름'] != 'nan') &
            (df_results['이름'].notna())
        ]
        # 반 정보 청소
        df_results['반'] = df_results['반'].astype(str).str.replace('0.0', '').replace('0', '').replace('nan', '').str.strip()
    
    df_results = df_results.replace('', 0).fillna(0)
    return df_info, df_results

def load_data():
    df_info, df_results = fetch_all_dataframes()
    return None, None, None, df_info, df_results


# --- 3. 공통 그래프 그리기 함수 (기존 로직 100% 동일) ---
def draw_report_figure(fig, s_row, student_name, student_grade, selected_test, cat_ratio, avg_cat_ratio, unit_data, unit_avg_data, unit_order):
    border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure, zorder=10)
    fig.patches.append(border)

    if os.path.exists("logo.png"):
        logo_img = plt.imread("logo.png")
        logo_ax = fig.add_axes([0.80, 0.915, 0.15, 0.045], zorder=15)
        logo_ax.imshow(logo_img)
        logo_ax.axis('off')

    txt_jeet = fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
    txt_title = fig.text(0.33, 0.88, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
    
    student_class = str(s_row.get('반', '')).strip()
    class_text = f"{student_class}  |  " if student_class and student_class != '0' else ""
    info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  {class_text}이름: {student_name}  |  과정: {selected_test}"
    txt_info = fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')

    txt_jeet.set_path_effects([path_effects.withStroke(linewidth=2, foreground=COLOR_RED)])
    txt_title.set_path_effects([path_effects.withStroke(linewidth=1.5, foreground=COLOR_NAVY)])
    txt_info.set_path_effects([path_effects.withStroke(linewidth=1, foreground='#222')])

    # --- 방사형 그래프 ---
    ax1 = fig.add_axes([0.10, 0.52, 0.32, 0.22], polar=True)
    all_cats = cat_ratio.index.tolist()
    ordered_labels = ['계산력'] + [c for c in all_cats if c != '계산력'] if '계산력' in all_cats else all_cats
    s_ordered = cat_ratio.reindex(ordered_labels)
    labels = s_ordered.index.tolist()
    s_vals = s_ordered.values.tolist() + [s_ordered.values[0]]
    angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
    
    max_s_val = max(s_vals) if len(s_vals) > 0 else 0
    ax1_limit = max(45, min(110, max_s_val + (max_s_val * 0.25) + 10)) 
    
    ax1.set_theta_direction(-1); ax1.set_theta_offset(np.pi/2.0)
    ax1.plot(angles, s_vals, color=COLOR_RED, linewidth=2.5, label='학생 점수')
    ax1.set_ylim(0, ax1_limit); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([]); ax1.set_yticklabels([]) 
    
    for i in range(len(labels)):
        angle = angles[i]; label_text = labels[i]
        dist_tb = ax1_limit * 1.05  
        dist_lr = ax1_limit * 1.02  
        ha, va, dist = ('center', 'bottom', dist_tb) if angle == 0 else ('left', 'center', dist_lr) if 0 < angle < np.pi else ('center', 'top', dist_tb) if angle == np.pi else ('right', 'center', dist_lr)
        if '문제\n해결력' in label_text: 
            dist += (ax1_limit * 0.08)
            ha = 'left' if 0 < angle < np.pi else 'right'
        ax1.text(angle, dist, label_text, fontsize=10, fontweight='bold', va=va, ha=ha, color=COLOR_NAVY)
        s_v = int(s_vals[i])
        td = s_v - (ax1_limit * 0.12) if s_v > ax1_limit * 0.85 else s_v + (ax1_limit * 0.12)
        txt_s = ax1.text(angle, td, f"{s_v}%", fontsize=9, fontweight='bold', color=COLOR_RED, va='center', ha='center')
        txt_s.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
    
    title1 = ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)
    title1.set_path_effects([path_effects.withStroke(linewidth=1, foreground=COLOR_NAVY)])

    # --- 단원별 성취도 그래프 ---
    ax2 = fig.add_axes([0.55, 0.54, 0.35, 0.18]) 
    x_pos = np.arange(len(unit_data))
    bar_width = 0.45 
    s_pct = (unit_data['득점'] / unit_data['배점'] * 100).fillna(0)
    max_b_val = s_pct.max() if not s_pct.empty else 0
    ax2_limit = max(40, min(110, max_b_val + (max_b_val * 0.25) + 15)) 
    
    ax2.bar(x_pos, s_pct, color=COLOR_STUDENT, alpha=0.9, width=bar_width, zorder=3)
    ax2.set_xticks(x_pos); ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
    ax2.tick_params(axis='x', which='both', length=0) 
    ax2.set_ylim(0, ax2_limit); ax2.grid(axis='y', color=COLOR_GRID, linestyle='-', linewidth=0.5, zorder=0)
    title2 = ax2.set_title("▶ 단원별 성취도 (%)", pad=25, fontsize=14, fontweight='bold', color=COLOR_NAVY)
    title2.set_path_effects([path_effects.withStroke(linewidth=1, foreground=COLOR_NAVY)])
    
    for i in range(len(x_pos)):
        val = int(s_pct.iloc[i])
        pos = x_pos[i]
        y_p = val + (ax2_limit * 0.04)
        t = ax2.text(pos, y_p, f"{val}%", ha='center', va='bottom', fontsize=7.5, fontweight='bold', color=COLOR_STUDENT)
        t.set_path_effects([path_effects.withStroke(linewidth=2, foreground='white')])

    # --- 하단 심층 분석 박스 ---
    rect_diag = plt.Rectangle((0.08, 0.12), 0.84, 0.35, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
    fig.patches.append(rect_diag)
    
    t_p1 = fig.text(0.11, 0.44, "▶ ", fontsize=13, fontweight='bold', color=COLOR_NAVY)
    t_p2 = fig.text(0.13, 0.44, " JEET", fontsize=13, fontweight='bold', color=COLOR_RED)
    t_p3 = fig.text(0.185, 0.44, f" 중등 수학 교육원 {student_name} 학생 심층 분석", fontsize=13, fontweight='bold', color=COLOR_NAVY)
    for t_obj in [t_p1, t_p2, t_p3]: t_obj.set_path_effects([path_effects.withStroke(linewidth=1, foreground=t_obj.get_color())])
    
    u_res = (unit_data['득점'] / unit_data['배점'] * 100).fillna(0)
    avg_val = int(cat_ratio.mean())
    
    if avg_val >= 80: eval_tier = "심화 개념까지 완벽히 소화하며 탁월한 수학적 직관력을 보여주는 최상위 수준의 성취도"
    elif avg_val >= 60: eval_tier = "안정적인 기본기를 바탕으로 성실한 학습 태도가 돋보이는 우수한 성취도"
    elif avg_val >= 20: eval_tier = "핵심 개념을 정립해 나가며 꾸준한 성장이 기대되는 도약 단계의 성취도"
    else: eval_tier = "수학적 기초 체력을 다지며 자신감을 키워가야 하는 잠재력 발현 단계의 성취도"
        
    diag_total = f"{student_name} 학생은 성취도 {avg_val}%를 기록하며, 현재 [{eval_tier}]를 보여주고 있습니다."

    c_best = cat_ratio[cat_ratio >= 80].index.str.replace('\n', '').tolist()
    c_good = cat_ratio[(cat_ratio >= 50) & (cat_ratio < 80)].index.str.replace('\n', '').tolist()
    c_weak = cat_ratio[(cat_ratio >= 20) & (cat_ratio < 50)].index.str.replace('\n', '').tolist()
    c_warn = cat_ratio[cat_ratio < 20].index.str.replace('\n', '').tolist()

    diag_combined = ""
    if c_best: diag_combined += f"특히 {', '.join([f'[{c}]' for c in c_best])} 영역에서 높은 이해도와 응용력을 보이며 탁월한 강점을 나타내고 있습니다. 해당 역영에서는 지속적으로 성취도를 유지할 수 있도록 연습이 필요합니다. "
    if c_good: diag_combined += f"{', '.join([f'[{c}]' for c in c_good])} 영역 역시 양호한 정답률을 유지하며 탄탄한 기본기를 증명했지만 실수를 줄이는 연습과 심화적인 부분의 학습이 필요합니다. "
    if c_weak: diag_combined += f"{', '.join([f'[{c}]' for c in c_weak])} 영역은 복합 개념 적용에 있어 다소 아쉬움이 남습니다. 해당 영역에 해당하는 개념, 응용 문제들을 해결하면서 정밀한 보완이 필요합니다. "
    if c_warn: diag_combined += f"무엇보다 {', '.join([f'[{c}]' for c in c_warn])} 영역은 기본적인 개념부터 다시 집중적으로 연습하여 성취도를 끌어올릴 수 있는 재학습이 필요해 보입니다. "

    g_best = u_res[u_res >= 80].index.tolist()
    g_good = u_res[(u_res >= 50) & (u_res < 80)].index.tolist()
    g_weak = u_res[(u_res >= 20) & (u_res < 50)].index.tolist()
    g_warn = u_res[u_res < 20].index.tolist()

    if g_best: diag_combined += f"세부 단원별로는 {', '.join([f'<{u}>' for u in g_best])} 단원의 완성도가 매우 훌륭합니다. 이 단원에서는 신유형 위주의 학습을 하며 성취도를 유지할 필요가 있습니다. "
    if g_good: diag_combined += f"{', '.join([f'<{u}>' for u in g_good])} 단원은 안정적인 궤도에 올랐으나 심화적인 부분에서 아쉬운 부분이 있으니 심화 위주의 학습을 하면서 성취도를 끌어올릴 필요가 있습니다. "
    if g_weak: diag_combined += f"{', '.join([f'<{u}>' for u in g_weak])} 단원은 오답이 발생하는 취약 유형에 대한 재학습을 하면서 오답을 줄이고, 난이도가 있는 문제들을 해결하면서 유형에 대한 연습이 필요합니다. "
    if g_warn: diag_combined += f"{', '.join([f'<{u}>' for u in g_warn])} 단원은 기초부터 다시 다지면서 쉬운 유형들에 대한 연습을 하면서 틀린 문제에 대한 맞춤형 클리닉 진행이 필요합니다."
    if not (g_best or g_good or g_weak or g_warn): diag_combined += "전반적인 단원 성취도가 고르게 나타나고 있습니다."

    weak_list = u_res[u_res < 40].index.tolist()
    if weak_list:
        sol_text = f"{student_name} 학생은 {', '.join([f'<{u}>' for u in weak_list])} 단원에 대한 철저한 오답 분석이 최우선 과제입니다. JEET만의 맞춤 솔루션인 JEET CARE+와 JDM(JEET DAILY MAKE UP) 시스템을 적극 활용하여 발견된 취약점을 빈틈없이 메워 나가며 다음 단계로의 도약을 준비하겠습니다." 
    else:
        sol_text = f"모든 단원에서 고르고 우수한 성취를 보이고 있는 만큼, 현재의 좋은 흐름을 유지하는 것이 중요합니다. 상위권 도약을 위한 고난도 심화 문항 도전과 실전 감각 유지를 목표로 JEET의 커리큘럼을 따라 한 단계 더 성장할 수 있도록 지도하겠습니다."

    sections = [("[종합 진단]", diag_total), ("[영역별&단원별 분석]", diag_combined), ("[JEET 맞춤 솔루션]", sol_text)]
    curr_y = 0.415 
    for subtitle, content in sections:
        stxt = fig.text(0.11, curr_y, subtitle, fontsize=9.5, fontweight='bold', color='#222')
        stxt.set_path_effects([path_effects.withStroke(linewidth=0.5, foreground='#222')])
        wrapped_content = textwrap.fill(content, width=65)
        ctxt = fig.text(0.11, curr_y - 0.015, wrapped_content, fontsize=8.2, linespacing=1.6, va='top', color='#333')
        num_lines = len(wrapped_content.split('\n'))
        curr_y -= (0.045 + (num_lines * 0.013))

    line_footer = plt.Line2D([0.05, 0.95], [0.10, 0.10], color=COLOR_NAVY, linewidth=1, transform=fig.transFigure); fig.lines.append(line_footer)
    campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜명로 10")]
    for i, (name, addr) in enumerate(campuses):
        fig.text([0.22, 0.50, 0.78][i], 0.07, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
        fig.text([0.22, 0.50, 0.78][i], 0.045, addr, ha='center', fontsize=7.5, color='#555')


# --- 4. 데이터 처리 함수 ---
def prepare_report_data(selected_test):
    _, _, _, df_info, df_results = load_data()
    df_info = df_info[df_info['시험명'] == selected_test]
    df_results = df_results[df_results['시험명'] == selected_test]
    df_results.columns = df_results.columns.astype(str)
    df_info['배점'] = pd.to_numeric(df_info['배점'], errors='coerce').fillna(3).astype(int)
    
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
    
    return df_info, df_results, avg_cat_ratio, unit_avg_data, unit_order, safe_to_int

def generate_jeet_expert_report(target_name, selected_test):
    try:
        df_info, df_results, avg_cat_ratio, unit_avg_data, unit_order, safe_to_int = prepare_report_data(selected_test)
        student_found = False
        pdf_buffer = io.BytesIO()
        with PdfPages(pdf_buffer) as pdf:
            for _, s_row in df_results.iterrows():
                student_name = str(s_row.get('이름', '')).strip()
                if not student_name or student_name == '0' or student_name != str(target_name).strip(): continue
                student_found = True
                student_grade = s_row.get('학년', '')
                analysis = df_info.copy()
                analysis['영역'] = analysis['영역'].str.replace('문제해결력', '문제\n해결력')
                analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
                analysis['득점'] = analysis['정답여부'] * analysis['배점']
                cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
                unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'})
                unit_data = unit_data.reindex([u for u in unit_order if u in unit_data.index])

                fig = plt.figure(figsize=(8.27, 11.69))
                draw_report_figure(fig, s_row, student_name, student_grade, selected_test, cat_ratio, avg_cat_ratio, unit_data, unit_avg_data, unit_order)
                pdf.savefig(fig); plt.close(fig)
        if not student_found: return False, None, "학생을 찾을 수 없습니다."
        return True, pdf_buffer, "리포트 생성 완료!"
    except Exception: return False, None, f"오류 발생: {traceback.format_exc()}"

def generate_batch_report(target_class, selected_test, selected_students=None):
    try:
        df_info, df_results, avg_cat_ratio, unit_avg_data, unit_order, safe_to_int = prepare_report_data(selected_test)
        class_students = df_results[df_results['반'].astype(str).str.strip() == str(target_class).strip()]
        if selected_students is not None:
            cleaned_selected = [str(s).strip() for s in selected_students]
            class_students = class_students[class_students['이름'].astype(str).str.strip().isin(cleaned_selected)]
        if class_students.empty: return False, None, "선택된 학생 데이터가 없습니다."
        pdf_buffer = io.BytesIO()
        with PdfPages(pdf_buffer) as pdf:
            for _, s_row in class_students.iterrows():
                student_name = str(s_row.get('이름', '')).strip()
                if not student_name or student_name == '0': continue
                student_grade = s_row.get('학년', '')
                analysis = df_info.copy()
                analysis['영역'] = analysis['영역'].str.replace('문제해결력', '문제\n해결력')
                analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
                analysis['득점'] = analysis['정답여부'] * analysis['배점']
                cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
                unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'})
                unit_data = unit_data.reindex([u for u in unit_order if u in unit_data.index])
                fig = plt.figure(figsize=(8.27, 11.69))
                draw_report_figure(fig, s_row, student_name, student_grade, selected_test, cat_ratio, avg_cat_ratio, unit_data, unit_avg_data, unit_order)
                pdf.savefig(fig); plt.close(fig)
        return True, pdf_buffer, f"일괄 생성 완료!"
    except Exception: return False, None, f"오류 발생: {traceback.format_exc()}"


# --- 5. Streamlit 웹 UI ---
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide", page_icon="📊")
col1, col2 = st.columns([8, 2])
with col1: st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")
with col2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

try:
    _, _, _, df_info_all, df_results_all = load_data()
except Exception as e:
    st.error(f"데이터베이스 로드 실패: {e}"); st.stop()

st.sidebar.header("📚 시험 과정 선택")
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("과정 선택:", test_list)
df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]

tab1, tab2, tab3 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력", "📚 일괄 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
    q_nums = df_info_filtered['문항번호'].tolist()
    if q_nums:
        with st.form("input_form", clear_on_submit=True):
            ci1, ci2, ci3, ci4 = st.columns(4)
            in_name = ci1.text_input("이름")
            in_class = ci2.text_input("반")
            in_school = ci3.text_input("학교")
            in_grade = ci4.selectbox("학년", ["중1", "중2", "중3"])
            st.markdown("---")
            answers = {}
            for i in range(0, len(q_nums), 5):
                cols = st.columns(5)
                for j, q in enumerate(q_nums[i:i+5]):
                    with cols[j]:
                        choice = st.radio(f"**{q}번**", ["O", "X"], key=f"q_{q}", horizontal=True)
                        answers[str(q)] = 1 if choice == "O" else 0
            if st.form_submit_button("Supabase에 저장하기", type="primary"):
                if not in_name.strip(): st.error("이름을 입력하세요.")
                else:
                    try:
                        new_row = {"시험명": selected_test, "이름": in_name.strip(), "반": in_class, "학교": in_school, "학년": in_grade}
                        new_row.update(answers)
                        get_supabase_client().table("Student_Results").insert(new_row).execute()
                        st.success("저장 완료!"); st.cache_data.clear()
                    except Exception as e: st.error(f"저장 오류: {e}")

with tab2:
    st.subheader("개별 리포트 생성")
    target = st.text_input("리포트 출력 학생 이름:")
    if st.button("PDF 리포트 생성", type="primary"):
        with st.spinner("생성 중..."):
            success, buf, msg = generate_jeet_expert_report(target.strip(), selected_test)
            if success: st.download_button("📥 다운로드", buf.getvalue(), f"{target}_리포트.pdf", "application/pdf")
            else: st.error(msg)

with tab3:
    st.subheader("반별 일괄 출력")
    if '반' in df_results_all.columns:
        # --- [이 부분을 수정합니다] ---
        current_test_results = df_results_all[df_results_all['시험명'] == selected_test]
        all_classes = current_test_results['반'].astype(str).str.strip().unique().tolist()
        class_list = sorted([c for c in all_classes if c and c != ''])
        # -----------------------------
        
        if class_list:
            target_c = st.selectbox("반 선택:", class_list)
            students = df_results_all[(df_results_all['시험명'] == selected_test) & (df_results_all['반'].astype(str).str.strip() == target_c)]['이름'].astype(str).str.strip().tolist()
            students = sorted([s for s in students if s and s not in ['0', 'nan']])
            sel_s = st.multiselect("출력 학생 선택:", options=students, default=students)
            if st.button("일괄 PDF 생성", type="primary"):
                with st.spinner("일괄 생성 중..."):
                    success, buf, msg = generate_batch_report(target_c, selected_test, sel_s)
                    if success: st.download_button("📥 일괄 다운로드", buf.getvalue(), f"{target_c}_일괄리포트.pdf", "application/pdf")
                    else: st.error(msg)
