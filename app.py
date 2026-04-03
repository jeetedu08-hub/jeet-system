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
from supabase import create_client, Client

# --- 1. 환경 및 폰트 설정 (에러 방지 안전장치 추가) ---
font_path = "malgun.ttf"
try:
    if os.path.exists(font_path):
        fm.fontManager.addfont(font_path)
        font_prop = fm.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    else:
        plt.rcParams['font.family'] = 'Malgun Gothic'
except Exception as e:
    print(f"폰트 로딩 실패 (기본 폰트로 대체됨): {e}")
    plt.rcParams['font.family'] = 'sans-serif'

plt.rcParams['axes.unicode_minus'] = False

COLOR_NAVY = '#1A237E'
COLOR_RED = '#D32F2F'
COLOR_STUDENT = '#0056B3' # 파란색
COLOR_UNIT = '#00796B'    # 청록색 (현재 미사용)
COLOR_AVG = '#757575'
COLOR_GRID = '#E0E0E0'
COLOR_BG = '#F8F9FA'

# --- 2. Supabase 연동 및 캐시 설정 ---
@st.cache_resource
def init_supabase() -> Client:
    url: str = st.secrets["SUPABASE_URL"]
    key: str = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

@st.cache_data(ttl=120)
def fetch_all_dataframes():
    supabase = init_supabase()
    # test_info 가져오기
    info_res = supabase.table("test_info").select("*").execute()
    df_info = pd.DataFrame(info_res.data)
    
    # student_results 가져오기
    results_res = supabase.table("student_results").select("*").execute()
    df_results = pd.DataFrame(results_res.data)
    
    if not df_results.empty:
        df_results = df_results.fillna(0)
        df_results.columns = df_results.columns.astype(str)
        
    return df_info, df_results


# --- 3. 공통 그래프 그리기 함수 (개별/일괄 출력에서 모두 사용) ---
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
    
    # 소속 반 정보 가져오기 (없으면 빈 칸)
    student_class = str(s_row.get('반', '')).strip()
    class_text = f"{student_class}  |  " if student_class and student_class != '0' else ""
    info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  {class_text}이름: {student_name}  |  과정: {selected_test}"
    txt_info = fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')

    txt_jeet.set_path_effects([path_effects.withStroke(linewidth=2, foreground=COLOR_RED)])
    txt_title.set_path_effects([path_effects.withStroke(linewidth=1.5, foreground=COLOR_NAVY)])
    txt_info.set_path_effects([path_effects.withStroke(linewidth=1, foreground='#222')])

    # --- 방사형 그래프 (크기 자동 조절 + 빨간색으로 변경) ---
    ax1 = fig.add_axes([0.10, 0.52, 0.32, 0.22], polar=True)
    all_cats = cat_ratio.index.tolist()
    ordered_labels = ['계산력'] + [c for c in all_cats if c != '계산력'] if '계산력' in all_cats else all_cats
    s_ordered = cat_ratio.reindex(ordered_labels)
    labels = s_ordered.index.tolist()
    s_vals = s_ordered.values.tolist() + [s_ordered.values[0]]
    angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
    
    # 동적 스케일링: 학생의 최고 점수를 기준으로 한도를 정함
    max_s_val = max(s_vals) if len(s_vals) > 0 else 0
    ax1_limit = max(45, min(110, max_s_val + (max_s_val * 0.25) + 10)) # 최소 45%, 상단 여유 확보
    
    ax1.set_theta_direction(-1); ax1.set_theta_offset(np.pi/2.0)
    ax1.plot(angles, s_vals, color=COLOR_RED, linewidth=2.5, label='학생 점수')
    ax1.set_ylim(0, ax1_limit); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([]); ax1.set_yticklabels([]) 
    
    for i in range(len(labels)):
        angle = angles[i]; label_text = labels[i]
        
        # 글자 위치를 그래프 바깥선에 예쁘게 밀착되도록 간격 대폭 축소
        dist_tb = ax1_limit * 1.05  # 위, 아래 글씨 간격
        dist_lr = ax1_limit * 1.02  # 양옆 글씨 간격
        
        ha, va, dist = ('center', 'bottom', dist_tb) if angle == 0 else ('left', 'center', dist_lr) if 0 < angle < np.pi else ('center', 'top', dist_tb) if angle == np.pi else ('right', 'center', dist_lr)
        
        # 두 줄짜리 텍스트(문제해결력)일 경우에만 살짝 더 여백을 줌
        if '문제\n해결력' in label_text: 
            dist += (ax1_limit * 0.08)
            ha = 'left' if 0 < angle < np.pi else 'right'
            
        ax1.text(angle, dist, label_text, fontsize=10, fontweight='bold', va=va, ha=ha, color=COLOR_NAVY)
        
        s_v = int(s_vals[i])
        # 점수 텍스트는 겹치지 않게 바깥쪽/안쪽으로 동적 띄어쓰기
        td = s_v - (ax1_limit * 0.12) if s_v > ax1_limit * 0.85 else s_v + (ax1_limit * 0.12)
        txt_s = ax1.text(angle, td, f"{s_v}%", fontsize=9, fontweight='bold', color=COLOR_RED, va='center', ha='center')
        txt_s.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
    
    title1 = ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)
    title1.set_path_effects([path_effects.withStroke(linewidth=1, foreground=COLOR_NAVY)])

    # --- 단원별 성취도 그래프 (크기 자동 조절 + 파란색으로 변경) ---
    ax2 = fig.add_axes([0.55, 0.54, 0.35, 0.18]) 
    x_pos = np.arange(len(unit_data))
    bar_width = 0.45 
    s_pct = (unit_data['득점'] / unit_data['배점'] * 100).fillna(0)
    
    # 막대 그래프 동적 스케일링: 눈금이 너무 높아서 막대가 바닥에 붙는 현상 해결
    max_b_val = s_pct.max() if not s_pct.empty else 0
    ax2_limit = max(40, min(110, max_b_val + (max_b_val * 0.25) + 15)) # 최고점에 비례해 유동적 여백 제공
    
    ax2.bar(x_pos, s_pct, color=COLOR_STUDENT, alpha=0.9, width=bar_width, zorder=3)
    
    ax2.set_xticks(x_pos); ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
    ax2.tick_params(axis='x', which='both', length=0) 
    ax2.set_ylim(0, ax2_limit); ax2.grid(axis='y', color=COLOR_GRID, linestyle='-', linewidth=0.5, zorder=0)
    title2 = ax2.set_title("▶ 단원별 성취도 (%)", pad=25, fontsize=14, fontweight='bold', color=COLOR_NAVY)
    title2.set_path_effects([path_effects.withStroke(linewidth=1, foreground=COLOR_NAVY)])
    
    for i in range(len(x_pos)):
        val = int(s_pct.iloc[i])
        pos = x_pos[i]
        
        # 막대 위쪽에 일정한 비율로 텍스트 예쁘게 배치
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
    
    # 1. 총평
    if avg_val >= 80:
        eval_tier = "심화 개념의 완벽한 체득과 날카로운 수학적 직관력을 겸비한 최상위권 수준의 성취를 보여줍니다. 고난도 문항 해결 능력이 탁월하며 자기 주도적 심화 학습이 충분히 가능한 상태입니다."
    elif avg_val >= 60:
        eval_tier = "탄탄한 기본기를 바탕으로 성실하게 학습의 밀도를 높여가는 우수한 성취를 보여줍니다. 안정적인 개념 활용 능력을 갖추고 있어, 향후 고난도 유형에 대한 도전이 성장의 핵심이 될 것입니다."
    elif avg_val >= 20:
        eval_tier = "핵심 개념을 정교하게 다듬어가는 과정에 있으며, 학습 잠재력이 점진적으로 발현되는 도약 단계의 성취를 보여줍니다. 꾸준한 학습 태도를 유지한다면 성취도 향상이 기대되는 구간입니다."
    else:
        eval_tier = "수학적 기초 체력을 보강하며 자신감을 쌓아가는 기틀 마련 단계의 성취를 보여줍니다. 학습적 결손을 메우고 성공적인 문제 풀이 경험을 축적하여 학습 동기를 부여하는 데 집중하고 있습니다."
        
    diag_total = f"{student_name} 학생은 성취도 {avg_val}%를 기록하며, 현재 [{eval_tier}]를 보여주고 있습니다."

    # 2. 영역별 분석
    c_best = cat_ratio[cat_ratio >= 80].index.str.replace('\n', '').tolist()
    c_good = cat_ratio[(cat_ratio >= 50) & (cat_ratio < 80)].index.str.replace('\n', '').tolist()
    c_weak = cat_ratio[(cat_ratio >= 20) & (cat_ratio < 50)].index.str.replace('\n', '').tolist()
    c_warn = cat_ratio[cat_ratio < 20].index.str.replace('\n', '').tolist()

    diag_combined = ""
    if c_best: diag_combined += f"특히 {', '.join([f'[{c}]' for c in c_best])} 영역에서 정교한 논리 구조와 높은 응용력을 보이며 압도적인 강점을 나타냅니다. 현재의 감각을 유지하며 최고 난도 문항을 통해 사고의 확장을 이어가야 합니다. "
    if c_good: diag_combined += f"{', '.join([f'[{c}]' for c in c_good])} 영역 또한 안정적인 정답률로 견고한 기본기를 증명하였습니다. 다만, 문항의 복합도가 높아질 때 발생하는 오답을 줄이기 위해 심화 유형에 대한 반복 훈련이 병행되어야 합니다. "
    if c_weak: diag_combined += f"{', '.join([f'[{c}]' for c in c_weak])} 영역은 복합 개념을 문제에 투영하는 과정에서 다소 고전하는 모습이 보입니다. 단편적인 문제 풀이보다는 개념 간의 유기적 연결성을 이해하고 유사 유형을 집중 분석하는 보완 작업이 필요합니다. "
    if c_warn: diag_combined += f"무엇보다 {', '.join([f'[{c}]' for c in c_warn])} 영역은 단원 간 연계성이 높은 만큼, 기초 개념의 재정립과 필수 유형에 대한 집중 학습이 필요합니다. 단계별 맞춤 문항을 통해 이해도를 근본적으로 끌어올려 실력을 끌어올릴 필요가 있습니다. "

    # 3. 단원별 분석
    g_best = u_res[u_res >= 80].index.tolist()
    g_good = u_res[(u_res >= 50) & (u_res < 80)].index.tolist()
    g_weak = u_res[(u_res >= 20) & (u_res < 50)].index.tolist()
    g_warn = u_res[u_res < 20].index.tolist()

    if g_best: diag_combined += f"{', '.join([f'<{u}>' for u in g_best])} 단원은 기본 및 응용 단계를 넘어 심화 단계까지 완벽히 소화하고 있습니다. 이제는 일반적인 유형 학습보다는 사고의 폭을 넓힐 수 있는 '킬러 문항' 중심의 도전적 학습이 필요한 단계입니다. "
    if g_good: diag_combined += f"{', '.join([f'<{u}>' for u in g_good])} 단원은 필수 유형들은 막힘없이 해결하고 있습니다. 이제는 단편적인 문제 풀이에서 벗어나, 개념을 다각도로 비트는 응용 문항에 대한 적응력을 키워야 할 때입니다. 심화 문항 도전 횟수를 늘려 사고의 유연성을 기른다면 보다 높은 수학 실력을 갖출 수 있을 것으로 판단됩니다. "
    if g_weak: diag_combined += f"{', '.join([f'<{u}>' for u in g_weak])} 단원은 개념의 조각들은 인지하고 있으나 실전 적용에서 병목 현상이 관찰됩니다. 난이도 높은 문제를 해결하기보다 핵심 유형의 반복 숙달이 필요하고, 성공적인 문제 풀이 경험을 축적하여 해당 단원의 수학적 자신감을 높인다면 한단계 더 발전할 수 있을 것으로 판단됩니다. "
    if g_warn: diag_combined += f"{', '.join([f'<{u}>' for u in g_warn])} 단원은 현재는 잠재력이 발현되기 전의 응축 단계입니다. 수학적 기초 체력이 부족할 뿐, 적절한 자극과 맞춤형 관리가 병행된다면 충분히 반등할 수 있는 가능성을 가지고 있습니다. 아이가 포기하지 않도록 가정에서 따뜻한 격려 부탁드립니다. "
    if not (g_best or g_good or g_weak or g_warn):
        diag_combined += "전반적인 단원별 성취도가 매우 균형 있게 나타나고 있습니다. 어느 한 쪽으로 치우치지 않는 고른 학습 균형이 큰 강점입니다."

    # 4. 향후 솔루션
    weak_list = u_res[u_res < 40].index.tolist()
    if weak_list:
        sol_text = f"{student_name} 학생은 {', '.join([f'<{u}>' for u in weak_list])} 단원에 대한 철저한 오답 분석이 최우선 과제입니다. JEET만의 맞춤 솔루션인 JEET CARE+와 JDM(JEET DAILY MAKE UP) 시스템을 적극 활용하여 발견된 취약점을 빈틈없이 메워 나가며 다음 단계로의 도약을 준비하겠습니다. " 
    else:
        sol_text = f"모든 단원에서 고르고 우수한 성취를 보이고 있는 만큼, 현재의 좋은 흐름을 유지하는 것이 중요합니다. 상위권 도약을 위한 고난도 심화 문항 도전과 실전 감각 유지를 목표로 JEET의 커리큘럼을 따라 한 단계 더 성장할 수 있도록 지도하겠습니다. "

    sections = [("[종합 진단]", diag_total), ("[영역별&단원별 분석]", diag_combined), ("[JEET 맞춤 솔루션]", sol_text)]

    # ✨ 텍스트 레이아웃 출력부 동적 튜닝 ✨ 
    # 1. 전체 글자 수 계산
    total_chars = sum(len(content) for _, content in sections)

    # 2. 글자 수에 따른 동적 레이아웃 변수 할당 (글이 길어지면 폰트를 줄이고 한 줄에 들어가는 글자 수를 늘림)
    if total_chars > 800:
        wrap_width = 82        # 한 줄당 글자 수
        main_fs = 6.8          # 본문 폰트 크기
        sub_fs = 8.5           # 소제목 폰트 크기
        y_offset = 0.012       # 소제목과 본문 사이 간격
        y_gap = 0.035          # 섹션 간 기본 간격
        line_height = 0.010    # 줄당 차지하는 높이
    elif total_chars > 600:
        wrap_width = 74
        main_fs = 7.5
        sub_fs = 9.0
        y_offset = 0.014
        y_gap = 0.040
        line_height = 0.0115
    else:
        wrap_width = 65
        main_fs = 8.2
        sub_fs = 9.5
        y_offset = 0.015
        y_gap = 0.045
        line_height = 0.013

    curr_y = 0.415 
    for subtitle, content in sections:
        # 1. 소제목 출력 (동적 폰트 적용)
        stxt = fig.text(0.11, curr_y, subtitle, fontsize=sub_fs, fontweight='bold', color='#222')
        stxt.set_path_effects([path_effects.withStroke(linewidth=0.5, foreground='#222')])
        
        # 본문 자동 줄바꿈 (동적 너비 적용)
        wrapped_content = textwrap.fill(content, width=wrap_width)
        
        # 2. 본문 출력 (동적 폰트 적용)
        ctxt = fig.text(0.11, curr_y - y_offset, wrapped_content, fontsize=main_fs, linespacing=1.6, va='top', color='#333')
        
        # 3. 다음 섹션 시작 위치 계산 (동적 높이 적용)
        num_lines = len(wrapped_content.split('\n'))
        curr_y -= (y_gap + (num_lines * line_height))

    curr_y = 0.415 
    for subtitle, content in sections:
        # 1. 소제목 출력 (동적 폰트 적용)
        stxt = fig.text(0.11, curr_y, subtitle, fontsize=sub_fs, fontweight='bold', color='#222')
        stxt.set_path_effects([path_effects.withStroke(linewidth=0.5, foreground='#222')])
        
        # 본문 자동 줄바꿈 (동적 너비 적용)
        wrapped_content = textwrap.fill(content, width=wrap_width)
        
        # 2. 본문 출력 (동적 폰트 적용)
        ctxt = fig.text(0.11, curr_y - y_offset, wrapped_content, fontsize=main_fs, linespacing=1.6, va='top', color='#333')
        
        # 3. 다음 섹션 시작 위치 계산 (동적 높이 적용)
        num_lines = len(wrapped_content.split('\n'))
        curr_y -= (y_gap + (num_lines * line_height))

    line_footer = plt.Line2D([0.05, 0.95], [0.10, 0.10], color=COLOR_NAVY, linewidth=1, transform=fig.transFigure); fig.lines.append(line_footer)
    campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜명로 10")]
    for i, (name, addr) in enumerate(campuses):
        fig.text([0.22, 0.50, 0.78][i], 0.07, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
        fig.text([0.22, 0.50, 0.78][i], 0.045, addr, ha='center', fontsize=7.5, color='#555')


# --- 4. 개별/일괄 데이터 처리 함수 ---
def prepare_report_data(selected_test):
    df_info, df_results = fetch_all_dataframes()
    
    df_info = df_info[df_info['시험명'] == selected_test]
    df_results = df_results[df_results['시험명'] == selected_test]
    df_info['배점'] = df_info['배점'].replace('', 3).fillna(3).astype(int)
    
    unit_order = df_info['단원'].drop_duplicates().tolist()
    q_cols = [str(q) for q in df_info['문항번호']]
    valid_cols = [col for col in df_results.columns if col in q_cols]
    
    def safe_to_int(val):
        try: return int(float(val))
        except: return 0

    df_scores = df_results[valid_cols].map(safe_to_int) if not df_results.empty else pd.DataFrame()
    avg_per_q = df_scores.mean() if not df_scores.empty else pd.Series()
    
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
    except Exception as e: return False, None, f"오류 발생: {traceback.format_exc()}"

def generate_batch_report(target_class, selected_test, selected_students=None):
    try:
        df_info, df_results, avg_cat_ratio, unit_avg_data, unit_order, safe_to_int = prepare_report_data(selected_test)
        
        class_students = df_results[df_results['반'].astype(str).str.strip() == str(target_class).strip()]
        
        if selected_students is not None:
            cleaned_selected = [str(s).strip() for s in selected_students]
            class_students = class_students[class_students['이름'].astype(str).str.strip().isin(cleaned_selected)]
        
        if class_students.empty:
            return False, None, f"선택된 학생 데이터가 없습니다. (이름 공백/오타 확인 필요)"
            
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
            
        return True, pdf_buffer, f"'{target_class}' 반 총 {len(class_students)}명의 리포트 일괄 생성 완료!"
    except Exception as e: return False, None, f"오류 발생: {traceback.format_exc()}"


# --- 5. Stream 추이 UI 구성 ---
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide", page_icon="📊")

if st.sidebar.button("🔄 데이터베이스 새로고침", use_container_width=True):
    st.cache_data.clear()
    st.success("데이터를 새로 불러왔습니다!")
    st.rerun()

st.sidebar.markdown("---") 

col1, col2 = st.columns([8, 2])
with col1: st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")
with col2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

try:
    df_info_all, df_results_all = fetch_all_dataframes()
except Exception as e:
    st.error(f"데이터베이스 로드 실패: {e}"); st.stop()

st.sidebar.header("📚 시험 과정 선택")
if not df_info_all.empty and '시험명' in df_info_all.columns:
    test_list = df_info_all['시험명'].dropna().unique().tolist()
else:
    test_list = []
    
if not test_list:
    st.warning("데이터베이스에 시험 정보가 없습니다.")
    st.stop()

selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)
df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]

tab1, tab2, tab3 = st.tabs(["✍️ 성적 입력", "📊 개별 리포트 출력", "📚 반별 일괄 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
    question_numbers = df_info_filtered['문항번호'].tolist()
    if question_numbers:
        with st.form("data_input_form", clear_on_submit=True):
            ci1, ci2, ci3, ci4 = st.columns(4)
            with ci1: input_name = st.text_input("이름")
            with ci2: input_class = st.text_input("반 (예: A반)")
            with ci3: input_school = st.text_input("학교")
            with ci4: input_grade = st.selectbox("학년", ["중1", "중2", "중3"])
            st.markdown("---")
            answers = {}
            for i in range(0, len(question_numbers), 5):
                cols = st.columns(5)
                for j, q_num in enumerate(question_numbers[i:i+5]):
                    with cols[j]:
                        choice = st.radio(f"**{q_num}번**", options=["O", "X"], horizontal=True, key=f"q_{q_num}")
                        answers[str(q_num)] = 1 if choice == "O" else 0
            
            if st.form_submit_button("DB에 성적 저장하기", type="primary"):
                clean_name = input_name.strip()
                if not clean_name: st.error("⚠ 이름을 입력해주세요.")
                else:
                    try:
                        # Supabase 테이블에 Insert할 딕셔너리 구성
                        new_record = {
                            "시험명": selected_test,
                            "이름": clean_name,
                            "반": input_class,
                            "학교": input_school,
                            "학년": input_grade
                        }
                        # 각 문항의 정답 여부를 레코드에 추가
                        for q_num, ans in answers.items():
                            new_record[str(q_num)] = ans

                        supabase = init_supabase()
                        supabase.table("student_results").insert(new_record).execute()
                        
                        st.success("성적이 DB에 안전하게 저장되었습니다!")
                        st.cache_data.clear() 
                    except Exception as e: st.error(f"저장 중 오류 발생: {e}")

with tab2:
    st.subheader(f"[{selected_test}] 개별 심층 분석 리포트 생성")
    target_student = st.text_input("리포트를 출력할 학생 이름:", placeholder="예: 홍길동")
    if st.button("개별 PDF 리포트 생성", type="primary"):
        with st.spinner("리포트 생성 중..."):
            success, buf, msg = generate_jeet_expert_report(target_student.strip(), selected_test)
            if success:
                st.success(msg)
                st.download_button("📥 PDF 다운로드", buf.getvalue(), f"{target_student}_리포트.pdf", "application/pdf")
            else: st.error(msg)

with tab3:
    st.subheader(f"[{selected_test}] 반별 전체 심층 분석 일괄 출력")
    
    if '반' in df_results_all.columns:
        all_classes = df_results_all['반'].astype(str).str.strip().unique().tolist()
        class_list = sorted([c for c in all_classes if c and c != '0' and c != 'nan'])
        
        if class_list:
            target_class = st.selectbox("📌 출력할 반을 선택하세요:", class_list)
            
            students_in_class = df_results_all[
                (df_results_all['시험명'] == selected_test) & 
                (df_results_all['반'].astype(str).str.strip() == target_class)
            ]['이름'].astype(str).str.strip().tolist()
            
            students_in_class = sorted([s for s in students_in_class if s and s != '0' and s != 'nan'])
            
            if students_in_class:
                selected_students = st.multiselect(
                    "👇 출력할 학생을 선택하세요 (제외할 학생의 'X'를 누르세요):", 
                    options=students_in_class, 
                    default=students_in_class
                )
            else:
                st.warning(f"⚠ 현재 선택하신 '{selected_test}' 과정에 '{target_class}' 학생 데이터가 없습니다.")
                selected_students = []
        else:
            st.info("DB에 아직 입력된 '반' 데이터가 없습니다.")
            target_class = st.text_input("출력할 반 이름 직접 입력:", placeholder="예: S반")
            selected_students = None
    else:
        st.warning("⚠ DB에 '반' 컬럼이 없어 수동으로 입력해야 합니다.")
        target_class = st.text_input("출력할 반 이름 직접 입력:", placeholder="예: S반")
        selected_students = None

    if st.button("반 전체/선택 일괄 PDF 생성", type="primary"):
        if not target_class.strip():
            st.error("반 이름을 입력하거나 선택해주세요.")
        elif selected_students is not None and len(selected_students) == 0:
            st.error("출력할 학생을 최소 1명 이상 선택해주세요.")
        else:
            with st.spinner(f"리포트를 하나로 모으는 중입니다. 잠시만 기다려주세요..."):
                success, buf, msg = generate_batch_report(target_class, selected_test, selected_students)
                if success:
                    st.success(msg)
                    st.download_button("📥 일괄 PDF 다운로드", buf.getvalue(), f"{target_class}_선택_리포트.pdf", "application/pdf")
                else: 
                    st.error(msg)
