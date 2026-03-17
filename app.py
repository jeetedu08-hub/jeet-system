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

# --- 2. 구글 스프레드시트 연동 함수 ---
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

# --- 3. PDF 생성 함수 (단원 자동 인식 + 로고 위치 정밀 수정본) ---
def generate_jeet_expert_report(target_name, selected_test):
    try:
        _, _, _, df_info, df_results = load_data()
        
        # 🌟 시험명 필터링
        df_info = df_info[df_info['시험명'] == selected_test]
        df_results = df_results[df_results['시험명'] == selected_test]
        
        df_results.columns = df_results.columns.astype(str)
        df_info['배점'] = df_info['배점'].replace('', 3).fillna(3).astype(int)
        
        # 🚨 [여기서 시험지 단원을 자동으로 읽어옵니다!]
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
                
                # 🌟 [로고 정밀 위치] 테두리 바로 아래
                if os.path.exists("logo.png"):
                    logo_img = plt.imread("logo.png")
                    logo_ax = fig.add_axes([0.80, 0.87, 0.15, 0.08], zorder=15)
                    logo_ax.imshow(logo_img)
                    logo_ax.axis('off')

                # 제목 위치 0.92
                fig.text(0.31, 0.92, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.92, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                
                # 학생 정보 0.88 (과정/학교/학년/이름 전체 표기)
                info_text = f"과정: {selected_test}  |  학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}"
                fig.text(0.5, 0.88, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')
  
                ax1 = fig.add_axes([0.10, 0.58, 0.32, 0.22], polar=True)
                # ... (차트 그리는 부분은 동일)
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
                ax1.plot(angles, s_vals, color=COLOR_STUDENT, linewidth=2.5)
                ax1.set_ylim(0, 110); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([])
  
                for i in range(len(labels)):
                    angle = angles[i]
                    ax1.text(angle, 115, labels[i], fontsize=10, fontweight='bold', color=COLOR_NAVY, ha='center')
                    s_val = int(s_vals[i]); a_val = int(a_vals[i])
                    txt_s = ax1.text(angle, s_val-10, f"{s_val}%", fontsize=9, fontweight='bold', color=COLOR_STUDENT, ha='right')
                    txt_a = ax1.text(angle, s_val-10, f"({a_val}%)", fontsize=9, fontweight='bold', color=COLOR_RED, ha='left')
                    for t in [txt_s, txt_a]: t.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])

                ax2 = fig.add_axes([0.55, 0.58, 0.35, 0.20])
                x_pos = np.arange(len(unit_data))
                ax2.bar(x_pos, unit_data['득점'], color=COLOR_STUDENT, alpha=0.8, width=0.5)
                ax2.scatter(x_pos, unit_avg_data['평균득점'], color=COLOR_RED, marker='_', s=1000, linewidth=3)
                ax2.set_xticks(x_pos)
                ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
                
                # 🚨 NaN 에러 방지용 리미트
                max_val = unit_data['배점'].max()
                ax2.set_ylim(0, (10 if pd.isna(max_val) or max_val == 0 else max_val) * 1.5)
  
                rect_diag = plt.Rectangle((0.08, 0.15), 0.84, 0.32, fill=True, facecolor=COLOR_BG, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                fig.text(0.13, 0.44, " JEET", fontsize=15, fontweight='bold', color=COLOR_RED)
                fig.text(0.185, 0.44, f"   중등 수학 교육원 {student_name} 학생 심층 분석", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                
                avg_val, total_avg_val = int(cat_ratio.mean()), int(avg_cat_ratio.mean())
                diag_content = f"1. 종합 진단: {student_name} 학생은 전체 평균 대비 {avg_val}%의 성취도를 보이고 있습니다.\n\n2. 강약점 분석: 단원별 분석 결과 시트에 입력된 정보를 바탕으로 보완이 필요합니다.\n\n3. JEET 맞춤 솔루션: 오답 노트를 통해 취약 부분을 집중 관리하시기 바랍니다."
                wrapped_lines = [textwrap.fill(p, width=54) for p in diag_content.split('\n\n')]
                fig.text(0.11, 0.41, "\n\n".join(wrapped_lines), fontsize=10.5, linespacing=1.8, va='top')
  
                pdf.savefig(fig); plt.close(fig)
            
        if not student_found: return False, None, f"학생을 찾을 수 없습니다."
        return True, pdf_buffer, f"리포트 생성 완료!"
  
    except Exception as e:
        return False, None, f"오류 발생:\n{traceback.format_exc()}"

# --- 4. 웹 UI 구성 ---
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide")
st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")

try:
    doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
except Exception:
    st.error("데이터 로드 실패")
    st.stop()

st.sidebar.header("📚 시험 과정 선택")
if st.sidebar.button("🔄 데이터 최신화", type="primary"):
    fetch_all_dataframes.clear(); st.rerun()

test_list = [str(t) for t in df_info_all['시험명'].unique() if str(t).strip() != '']
selected_test = st.sidebar.selectbox("시험 선택:", test_list)

tab1, tab2 = st.tabs(["📝 성적 입력", "📑 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 성적 입력")
    df_f = df_info_all[df_info_all['시험명'] == selected_test]
    q_nums = df_f['문항번호'].tolist()
    with st.form("input_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        i_n = c1.text_input("이름"); i_s = c2.text_input("학교"); i_g = c3.selectbox("학년", ["중1", "중2", "중3"])
        cols = st.columns(5); ans = {}
        for i, q in enumerate(q_nums):
            ans[str(q)] = cols[i%5].number_input(f"{q}번", 0, 1, 0)
        if st.form_submit_button("저장"):
            row = []
            for h in ws_results.row_values(1):
                if h == '시험명': row.append(selected_test)
                elif h == '이름': row.append(i_n)
                elif h == '학교': row.append(i_s)
                elif h == '학년': row.append(i_g)
                elif h in ans: row.append(ans[h])
                else: row.append("")
            ws_results.append_row(row); st.success("저장 완료!"); fetch_all_dataframes.clear()

with tab2:
    st.subheader(f"[{selected_test}] 리포트 출력")
    target = st.text_input("학생 이름")
    if st.button("리포트 생성"):
        success, buf, msg = generate_jeet_expert_report(target, selected_test)
        if success:
            st.download_button("📥 다운로드", buf.getvalue(), f"{target}_{selected_test}.pdf", "application/pdf")
        else: st.error(msg)
