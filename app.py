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
            if student_name != str(target_name).strip(): continue
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

                # 🌟 [수정 핵심] 로고를 천장(0.91)에 붙이고 제목(0.86)을 내려서 겹침 완전 해결
                if os.path.exists("logo.png"):
                    logo_img = plt.imread("logo.png")
                    logo_ax = fig.add_axes([0.78, 0.91, 0.15, 0.05], zorder=15)
                    logo_ax.imshow(logo_img)
                    logo_ax.axis('off')

                fig.text(0.31, 0.86, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.86, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                
                info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}  |  과정: {selected_test}"
                fig.text(0.5, 0.81, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')
  
                # 이후 차트 로직 (동일)
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
                ax1.plot(angles, a_vals, color=COLOR_AVG, linewidth=1, linestyle='--')
                ax1.fill(angles, a_vals, color=COLOR_AVG, alpha=0.1)
                ax1.plot(angles, s_vals, color=COLOR_STUDENT, linewidth=2.5)
                ax1.set_ylim(0, 110); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([]); ax1.set_yticklabels([])
                for i in range(len(labels)):
                    angle = angles[i]; label_text = labels[i]
                    if angle == 0: ha, va, d = 'center', 'bottom', 115
                    elif 0 < angle < np.pi: ha, va, d = 'left', 'center', 110
                    elif angle == np.pi: ha, va, d = 'center', 'top', 115
                    else: ha, va, d = 'right', 'center', 110
                    ax1.text(angle, d, label_text, fontsize=10, fontweight='bold', va=va, ha=ha, color=COLOR_NAVY)
                    s_v, a_v = int(s_vals[i]), int(a_vals[i])
                    td = s_v + 10 if s_v < 85 else s_v - 18
                    txt_s = ax1.text(angle, td, f"{s_v}%", fontsize=9, fontweight='bold', color=COLOR_STUDENT, va='center', ha='right')
                    txt_a = ax1.text(angle, td, f" ({a_v}%)", fontsize=9, fontweight='bold', color=COLOR_RED, va='center', ha='left')
                    for t in [txt_s, txt_a]: t.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
                ax2 = fig.add_axes([0.55, 0.52, 0.35, 0.20])
                x_pos = np.arange(len(unit_data))
                bars = ax2.bar(x_pos, unit_data['득점'], color=COLOR_STUDENT, alpha=0.8, width=0.5, zorder=3)
                ax2.scatter(x_pos, unit_avg_data['평균득점'], color=COLOR_RED, marker='_', s=1000, linewidth=3, zorder=4)
                ax2.set_xticks(x_pos); ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
                ax2.set_ylim(0, (unit_data['배점'].max() or 10) * 1.5)
                ax2.grid(axis='y', color=COLOR_GRID, linestyle='-', linewidth=0.5, zorder=0)
                for i, bar in enumerate(bars):
                    sv, av = int(bar.get_height()), int(unit_avg_data['평균득점'].iloc[i])
                    ax2.text(bar.get_x() + bar.get_width()/2, sv + 0.5, f"{sv}", ha='right', va='bottom', fontsize=9, color=COLOR_STUDENT)
                    ax2.text(bar.get_x() + bar.get_width()/2, sv + 0.5, f" ({av})", ha='left', va='bottom', fontsize=9, color=COLOR_RED)
                rect_diag = plt.Rectangle((0.08, 0.15), 0.84, 0.30, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                fig.text(0.11, 0.42, "▶ JEET 중등 수학 교육원 분석", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                avg_val, total_avg_val = int(cat_ratio.mean()), int(avg_cat_ratio.mean())
                diag_content = f"1. 종합 진단: 전체 평균({total_avg_val}%) 대비 {avg_val}% 기록\n2. 취약 분석: 맞춤형 클리닉 필요\n3. 솔루션: 오답 노트 중심 반복 학습 권장"
                fig.text(0.11, 0.38, diag_content, fontsize=10.5, linespacing=1.8, va='top', ha='left', color='#333')
                pdf.savefig(fig); plt.close(fig)
        return True, pdf_buffer, "성공적으로 생성되었습니다!"
    except Exception as e:
        return False, None, f"오류 발생: {traceback.format_exc()}"

# --- 4. Streamlit 웹 UI (O/X 입력 설정 유지) ---
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide")
col1, col2 = st.columns([8, 2])
with col1: st.title("📊 JEET 성적 관리 시스템")
with col2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("시험 과정 선택:", test_list)
df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]

tab1, tab2 = st.tabs(["📝 성적 입력", "📑 리포트 출력"])

with tab1:
    with st.form("input_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        n, s, g = c1.text_input("이름"), c2.text_input("학교"), c3.selectbox("학년", ["중1", "중2", "중3"])
        q_nums = df_info_filtered['문항번호'].tolist()
        cols = st.columns(5); answers = {}
        for i, q in enumerate(q_nums):
            with cols[i % 5]:
                choice = st.radio(f"**{q}번**", ["O", "X"], horizontal=True, key=f"q_{q}")
                answers[str(q)] = 1 if choice == "O" else 0
        if st.form_submit_button("저장"):
            row = [selected_test, n, s, g] + [answers.get(str(c), "") for c in ws_results.row_values(1)[4:]]
            ws_results.append_row(row); st.success("저장 완료!"); st.cache_data.clear()

with tab2:
    target = st.text_input("학생 이름:")
    if st.button("PDF 생성"):
        success, buf, msg = generate_jeet_expert_report(target, selected_test)
        if success: st.download_button("📥 다운로드", buf.getvalue(), f"{target}_리포트.pdf", "application/pdf")
