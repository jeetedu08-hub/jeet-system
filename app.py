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
COLOR_STUDENT = '#0056B3' # 파란색
COLOR_GRID = '#E0E0E0'
COLOR_BG = '#F8F9FA'

# --- 2. Supabase 연동 설정 ---
@st.cache_resource
def init_connection():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_connection()

@st.cache_data(ttl=120)
def fetch_all_dataframes():
    info_res = supabase.table('Test_Info').select("*").execute()
    df_info = pd.DataFrame(info_res.data)
    results_res = supabase.table('Student_Results').select("*").execute()
    df_results = pd.DataFrame(results_res.data)
    df_results = df_results.replace('', 0).fillna(0)
    return df_info, df_results

def load_data():
    return fetch_all_dataframes()

# --- 3. 공통 그래프 및 ✨심층 분석 박스✨ 그리기 함수 (로직 완전 복구) ---
def draw_report_figure(fig, s_row, student_name, student_grade, selected_test, cat_ratio, avg_cat_ratio, unit_data, unit_avg_data, unit_order):
    border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure, zorder=10)
    fig.patches.append(border)

    if os.path.exists("logo.png"):
        logo_img = plt.imread("logo.png")
        logo_ax = fig.add_axes([0.80, 0.915, 0.15, 0.045], zorder=15)
        logo_ax.imshow(logo_img); logo_ax.axis('off')

    fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right', path_effects=[path_effects.withStroke(linewidth=2, foreground=COLOR_RED)])
    fig.text(0.33, 0.88, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left', path_effects=[path_effects.withStroke(linewidth=1.5, foreground=COLOR_NAVY)])
    
    student_class = str(s_row.get('반', '')).strip()
    class_text = f"{student_class}  |  " if student_class and student_class != '0' else ""
    info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  {class_text}이름: {student_name}  |  과정: {selected_test}"
    fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')

    # --- 방사형 그래프 (빨간색) ---
    ax1 = fig.add_axes([0.10, 0.52, 0.32, 0.22], polar=True)
    all_cats = cat_ratio.index.tolist()
    ordered_labels = ['계산력'] + [c for c in all_cats if c != '계산력'] if '계산력' in all_cats else all_cats
    s_ordered = cat_ratio.reindex(ordered_labels)
    labels, s_vals = s_ordered.index.tolist(), s_ordered.values.tolist() + [s_ordered.values[0]]
    angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
    ax1_limit = max(45, min(110, max(s_vals) + 20))
    ax1.set_theta_direction(-1); ax1.set_theta_offset(np.pi/2.0)
    ax1.plot(angles, s_vals, color=COLOR_RED, linewidth=2.5)
    ax1.set_ylim(0, ax1_limit); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([]); ax1.set_yticklabels([]) 
    for i, label in enumerate(labels):
        angle, dist = angles[i], ax1_limit * 1.05
        ax1.text(angle, dist, label, fontsize=10, fontweight='bold', ha='center', color=COLOR_NAVY)
        ax1.text(angle, s_vals[i], f"{int(s_vals[i])}%", fontsize=9, fontweight='bold', color=COLOR_RED, ha='center', va='center', path_effects=[path_effects.withStroke(linewidth=3, foreground='white')])
    ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)

    # --- 단원별 성취도 그래프 (파란색) ---
    ax2 = fig.add_axes([0.55, 0.54, 0.35, 0.18])
    s_pct = (unit_data['득점'] / unit_data['배점'] * 100).fillna(0)
    ax2.bar(range(len(s_pct)), s_pct, color=COLOR_STUDENT, width=0.45, zorder=3)
    ax2.set_xticks(range(len(s_pct))); ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in s_pct.index], fontsize=8, fontweight='bold')
    ax2.set_ylim(0, max(40, min(110, s_pct.max() + 20))); ax2.grid(axis='y', color=COLOR_GRID, zorder=0)
    ax2.set_title("▶ 단원별 성취도 (%)", pad=25, fontsize=14, fontweight='bold', color=COLOR_NAVY)

    # --- ✨ 하단 심층 분석 박스 (원장님 오리지널 로직 100% 복구) ✨ ---
    fig.patches.append(plt.Rectangle((0.08, 0.12), 0.84, 0.35, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure))
    fig.text(0.11, 0.44, "▶ ", fontsize=13, fontweight='bold', color=COLOR_NAVY)
    fig.text(0.13, 0.44, " JEET", fontsize=13, fontweight='bold', color=COLOR_RED)
    fig.text(0.185, 0.44, f" 중등 수학 교육원 {student_name} 학생 심층 분석", fontsize=13, fontweight='bold', color=COLOR_NAVY)
    
    u_res = (unit_data['득점'] / unit_data['배점'] * 100).fillna(0)
    avg_val = int(cat_ratio.mean())
    
    # 1. 총평 로직
    if avg_val >= 80: eval_tier = "심화 개념까지 완벽히 소화하며 탁월한 수학적 직관력을 보여주는 최상위 수준의 성취도"
    elif avg_val >= 60: eval_tier = "안정적인 기본기를 바탕으로 성실한 학습 태도가 돋보이는 우수한 성취도"
    elif avg_val >= 20: eval_tier = "핵심 개념을 정립해 나가며 꾸준한 성장이 기대되는 도약 단계의 성취도"
    else: eval_tier = "수학적 기초 체력을 다지며 자신감을 키워가야 하는 잠재력 발현 단계의 성취도"
    diag_total = f"{student_name} 학생은 성취도 {avg_val}%를 기록하며, 현재 [{eval_tier}]를 보여주고 있습니다."

    # 2. 영역별 & 단원별 상세 분석 로직
    c_best = cat_ratio[cat_ratio >= 80].index.str.replace('\n', '').tolist()
    c_good = cat_ratio[(cat_ratio >= 50) & (cat_ratio < 80)].index.str.replace('\n', '').tolist()
    c_weak = cat_ratio[cat_ratio < 50].index.str.replace('\n', '').tolist()
    
    diag_combined = ""
    if c_best: diag_combined += f"특히 {', '.join([f'[{c}]' for c in c_best])} 영역에서 높은 이해도와 응용력을 보이며 탁월한 강점을 나타내고 있습니다. "
    if c_good: diag_combined += f"{', '.join([f'[{c}]' for c in c_good])} 영역 역시 양호한 정답률을 유지하며 탄탄한 기본기를 증명했습니다. "
    if c_weak: diag_combined += f"{', '.join([f'[{c}]' for c in c_weak])} 영역은 복합 개념 적용에 있어 다소 아쉬움이 남으므로 정밀한 보완이 필요합니다. "

    u_best = u_res[u_res >= 80].index.tolist()
    u_weak = u_res[u_res < 40].index.tolist()
    if u_best: diag_combined += f"세부 단원별로는 {', '.join([f'<{u}>' for u in u_best])} 단원의 완성도가 매우 훌륭합니다. "
    if u_weak: diag_combined += f"다만 {', '.join([f'<{u}>' for u in u_weak])} 단원은 오답 유형에 대한 재학습이 필요해 보입니다."

    # 3. 솔루션 로직
    if u_weak: sol_text = f"{student_name} 학생은 취약 단원에 대한 철저한 오답 분석이 최우선 과제입니다. JEET만의 맞춤 솔루션인 JEET CARE+와 JDM 시스템을 적극 활용하여 발견된 취약점을 빈틈없이 메워 나가겠습니다."
    else: sol_text = f"모든 단원에서 고른 성취를 보이고 있는 만큼, 상위권 도약을 위한 고난도 심화 문항 도전과 실전 감각 유지를 목표로 JEET의 커리큘럼에 맞춰 지도하겠습니다."

    sections = [("[종합 진단]", diag_total), ("[영역별&단원별 분석]", diag_combined), ("[JEET 맞춤 솔루션]", sol_text)]
    curr_y = 0.415
    for subtitle, content in sections:
        fig.text(0.11, curr_y, subtitle, fontsize=9.5, fontweight='bold', color='#222')
        wrapped = textwrap.fill(content, width=65)
        fig.text(0.11, curr_y - 0.015, wrapped, fontsize=8.2, linespacing=1.6, va='top', color='#333')
        curr_y -= (0.045 + (len(wrapped.split('\n')) * 0.013))

    campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜명로 10")]
    for i, (name, addr) in enumerate(campuses):
        fig.text([0.22, 0.50, 0.78][i], 0.07, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
        fig.text([0.22, 0.50, 0.78][i], 0.045, addr, ha='center', fontsize=7.5, color='#555')

# --- 4. 데이터 처리 로직 (100% 복원) ---
def prepare_report_data(selected_test):
    df_info, df_results = load_data()
    df_info = df_info[df_info['시험명'] == selected_test]
    df_results = df_results[df_results['시험명'] == selected_test]
    df_results.columns = df_results.columns.astype(str)
    df_info['배점'] = pd.to_numeric(df_info['배점'], errors='coerce').fillna(3).astype(int)
    def safe_to_int(val):
        try: return int(float(val))
        except: return 0
    unit_order = df_info['단원'].drop_duplicates().tolist()
    q_cols = [str(q) for q in df_info['문항번호']]
    df_scores = df_results[q_cols].applymap(safe_to_int)
    avg_per_q = df_scores.mean()
    total_analysis = df_info.copy()
    total_analysis['평균득점'] = total_analysis['문항번호'].apply(lambda x: avg_per_q.get(str(x), 0))
    total_analysis['영역'] = total_analysis['영역'].str.replace('문제해결력', '문제\n해결력')
    avg_cat_ratio = (total_analysis.groupby('영역')['평균득점'].sum() / total_analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
    return df_info, df_results, avg_cat_ratio, unit_order, safe_to_int

def generate_jeet_expert_report(target_name, selected_test):
    try:
        df_info, df_results, avg_cat_ratio, unit_order, safe_to_int = prepare_report_data(selected_test)
        student_found = False
        pdf_buffer = io.BytesIO()
        with PdfPages(pdf_buffer) as pdf:
            for _, s_row in df_results.iterrows():
                if str(s_row.get('이름', '')).strip() != str(target_name).strip(): continue
                student_found = True
                analysis = df_info.copy()
                analysis['득점'] = [safe_to_int(s_row.get(str(q), 0)) * b for q, b in zip(analysis['문항번호'], analysis['배점'])]
                cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
                unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'}).reindex(unit_order)
                fig = plt.figure(figsize=(8.27, 11.69))
                draw_report_figure(fig, s_row, str(target_name), s_row.get('학년',''), selected_test, cat_ratio, avg_cat_ratio, unit_data, None, unit_order)
                pdf.savefig(fig); plt.close(fig)
        return (True, pdf_buffer, "리포트 생성 완료!") if student_found else (False, None, "학생을 찾을 수 없습니다.")
    except Exception: return False, None, traceback.format_exc()

def generate_batch_report(target_class, selected_test, selected_students=None):
    try:
        df_info, df_results, avg_cat_ratio, unit_order, safe_to_int = prepare_report_data(selected_test)
        class_students = df_results[df_results['반'].astype(str).str.strip() == str(target_class).strip()]
        if selected_students:
            class_students = class_students[class_students['이름'].astype(str).str.strip().isin([s.strip() for s in selected_students])]
        pdf_buffer = io.BytesIO()
        with PdfPages(pdf_buffer) as pdf:
            for _, s_row in class_students.iterrows():
                analysis = df_info.copy()
                analysis['득점'] = [safe_to_int(s_row.get(str(q), 0)) * b for q, b in zip(analysis['문항번호'], analysis['배점'])]
                cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
                unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'}).reindex(unit_order)
                fig = plt.figure(figsize=(8.27, 11.69))
                draw_report_figure(fig, s_row, str(s_row['이름']), s_row.get('학년',''), selected_test, cat_ratio, avg_cat_ratio, unit_data, None, unit_order)
                pdf.savefig(fig); plt.close(fig)
        return True, pdf_buffer, f"총 {len(class_students)}명의 리포트 일괄 생성 완료!"
    except Exception: return False, None, traceback.format_exc()

# --- 5. Streamlit 웹 UI ---
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide", page_icon="📊")
col1, col2 = st.columns([8, 2])
with col1: st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")
with col2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

try:
    df_info_all, df_results_all = load_data()
    test_list = df_info_all['시험명'].dropna().unique().tolist()
    selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)
    df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]
    
    tab1, tab2, tab3 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력", "🗂 일괄 리포트 출력"])

    with tab1:
        st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
        with st.form("input_form", clear_on_submit=True):
            ci1, ci2, ci3, ci4 = st.columns(4)
            name, u_class, school, grade = ci1.text_input("이름"), ci2.text_input("반"), ci3.text_input("학교"), ci4.selectbox("학년", ["중1","중2","중3"])
            st.markdown("---")
            answers = {str(q): (1 if st.radio(f"**{q}번**", ["O","X"], horizontal=True, key=f"q{q}") == "O" else 0) for q in df_info_filtered['문항번호']}
            if st.form_submit_button("데이터베이스에 성적 저장하기", type="primary"):
                submit_data = {"시험명": selected_test, "이름": name.strip(), "반": u_class.strip(), "학교": school.strip(), "학년": grade, **answers}
                supabase.table('Student_Results').insert(submit_data).execute()
                st.success("데이터베이스에 저장되었습니다!"); st.cache_data.clear()

    with tab2:
        st.subheader(f"[{selected_test}] 개별 심층 분석 리포트 생성")
        target = st.text_input("리포트를 출력할 학생 이름:", placeholder="예: 홍길동")
        if st.button("개별 PDF 리포트 생성", type="primary"):
            success, buf, msg = generate_jeet_expert_report(target, selected_test)
            if success: st.success(msg); st.download_button("📥 PDF 다운로드", buf.getvalue(), f"{target}_리포트.pdf", "application/pdf")
            else: st.error(msg)

    with tab3:
        st.subheader(f"[{selected_test}] 반별 전체 심층 분석 일괄 출력")
        if '반' in df_results_all.columns:
            classes = sorted([c for c in df_results_all['반'].astype(str).str.strip().unique().tolist() if c and c not in ['0', 'nan']])
            target_class = st.selectbox("📌 출력할 반을 선택하세요:", classes)
            students = sorted([s for s in df_results_all[(df_results_all['시험명'] == selected_test) & (df_results_all['반'].astype(str).str.strip() == target_class)]['이름'].astype(str).str.strip().tolist() if s and s not in ['0', 'nan']])
            selected_students = st.multiselect("👇 출력할 학생을 선택하세요:", options=students, default=students)
            if st.button("반 전체/선택 일괄 PDF 생성", type="primary"):
                success, buf, msg = generate_batch_report(target_class, selected_test, selected_students)
                if success: st.success(msg); st.download_button("📥 일괄 PDF 다운로드", buf.getvalue(), f"{target_class}_일괄_리포트.pdf", "application/pdf")
                else: st.error(msg)
except Exception as e:
    st.error(f"데이터베이스 로드 중 오류 발생: {e}")
