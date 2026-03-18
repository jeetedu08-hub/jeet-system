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

# --- 2. 구글 스프레드시트 연동 ---
@st.cache_resource
def get_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # Streamlit Secrets 우선 확인
    if "GOOGLE_JSON" in st.secrets:
        creds_dict = json.loads(st.secrets["GOOGLE_JSON"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    elif os.path.exists("secrets.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("secrets.json", scope)
    else:
        return None
    
    client = gspread.authorize(creds)
    return client.open_by_url("https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit")

@st.cache_data(ttl=60)
def fetch_data():
    doc = get_google_sheet()
    if doc is None: return None, None
    df_info = pd.DataFrame(doc.worksheet('Test_Info').get_all_records())
    df_results = pd.DataFrame(doc.worksheet('Student_Results').get_all_records())
    return df_info, df_results

# --- 3. 리포트 생성 함수 (심층 분석 및 그래프 포함) ---
def generate_report(target_name, selected_test):
    try:
        df_info_all, df_results_all = fetch_data()
        df_info = df_info_all[df_info_all['시험명'] == selected_test].copy()
        df_results = df_results_all[df_results_all['시험명'] == selected_test].copy()
        
        df_results = df_results.replace('', 0).fillna(0)
        df_info['배점'] = pd.to_numeric(df_info['배점'], errors='coerce').fillna(3).astype(int)
        unit_order = df_info['단원'].drop_duplicates().tolist()
        
        q_cols = [str(q) for q in df_info['문항번호']]
        df_scores = df_results[[c for c in df_results.columns if str(c) in q_cols]].apply(pd.to_numeric)
        avg_per_q = df_scores.mean()
        
        total_analysis = df_info.copy()
        total_analysis['평균득점'] = total_analysis['문항번호'].apply(lambda x: avg_per_q.get(str(x), 0))
        avg_cat_ratio = (total_analysis.groupby('영역')['평균득점'].sum() / total_analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
        unit_avg_data = total_analysis.groupby('단원')['평균득점'].sum().reindex(unit_order).fillna(0)

        pdf_buffer = io.BytesIO()
        student_found = False

        for _, s_row in df_results.iterrows():
            if str(s_row.get('이름', '')).strip() != str(target_name).strip(): continue
            student_found = True
            
            analysis = df_info.copy()
            analysis['정답여부'] = [int(float(s_row.get(str(q), 0))) for q in analysis['문항번호']]
            analysis['득점'] = analysis['정답여부'] * analysis['배점']
            cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
            unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'}).reindex(unit_order).fillna(0)

            with PdfPages(pdf_buffer) as pdf:
                fig = plt.figure(figsize=(8.27, 11.69))
                plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5, transform=fig.transFigure)
                fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.88, f'{selected_test} 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY)
                
                info_text = f"과정: {selected_test}  |  학교: {s_row.get('학교', '0')}  |  이름: {target_name}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold')

                # 1. 방사형 차트
                ax1 = fig.add_axes([0.1, 0.52, 0.32, 0.22], polar=True)
                labels = cat_ratio.index.tolist()
                angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
                s_vals = cat_ratio.values.tolist() + [cat_ratio.values[0]]
                a_vals = avg_cat_ratio.reindex(labels).values.tolist() + [avg_cat_ratio.iloc[0]]
                
                ax1.set_theta_direction(-1); ax1.set_theta_offset(np.pi/2.0)
                ax1.plot(angles, a_vals, color=COLOR_AVG, linewidth=1, linestyle='--')
                ax1.plot(angles, s_vals, color=COLOR_STUDENT, linewidth=2.5)
                ax1.set_xticks(angles[:-1]); ax1.set_xticklabels(labels, fontsize=10, fontweight='bold')
                ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)

                # 2. 막대 차트
                ax2 = fig.add_axes([0.55, 0.52, 0.35, 0.20])
                x = np.arange(len(unit_order))
                ax2.bar(x, unit_data['득점'], color=COLOR_STUDENT, alpha=0.7, width=0.4)
                ax2.scatter(x, unit_avg_data, color=COLOR_RED, marker='_', s=500, linewidth=3)
                ax2.set_xticks(x); ax2.set_xticklabels([textwrap.fill(l, 6) for l in unit_order], fontsize=8)
                ax2.set_title("▶ 단원별 성취도", pad=15, fontsize=14, fontweight='bold', color=COLOR_NAVY)

                # 3. 심층 분석 섹션
                rect = plt.Rectangle((0.08, 0.15), 0.84, 0.32, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect)
                
                avg_val = int(cat_ratio.mean())
                total_avg = int(avg_cat_ratio.mean())
                diag_text = (
                    f"1. 종합 진단: 전체 평균({total_avg}%) 대비 {avg_val}%의 성취도를 보입니다.\n\n"
                    f"2. 강약점 분석: '{cat_ratio.idxmax()}' 영역이 우수하며 보완이 필요합니다.\n\n"
                    f"3. JEET 맞춤 솔루션: 오답 노트를 통한 개념 재정립을 권장합니다."
                )
                fig.text(0.11, 0.41, diag_text, fontsize=11, linespacing=2, va='top')
                fig.text(0.5, 0.08, "죽전 캠퍼스: 263-8003 | 기흥구 죽현로 29", ha='center', fontsize=10, fontweight='bold')

                pdf.savefig(fig); plt.close(fig)
                break
        return student_found, pdf_buffer, "성공"
    except:
        return False, None, traceback.format_exc()

# --- 4. 메인 화면 ---
st.set_page_config(page_title="JEET 관리 시스템", layout="wide")
df_info_all, df_results_all = fetch_data()

if df_info_all is None:
    st.error("⚠️ 구글 시트 연결 중... Streamlit Secrets 설정을 완료해 주세요.")
    st.stop()

test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("시험 선택", test_list)

tab1, tab2 = st.tabs(["📝 성적 입력", "📑 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 성적 입력")
    q_nums = df_info_all[df_info_all['시험명'] == selected_test]['문항번호'].tolist()
    with st.form("input_form"):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input("이름")
        school = c2.text_input("학교")
        grade = c3.selectbox("학년", ["중1", "중2", "중3"])
        
        st.write("---")
        st.write("📊 문항 정답 클릭 (**O: 정답, X: 오답**)")
        ans = {}
        for i in range(0, len(q_nums), 5):
            cols = st.columns(5)
            for j in range(5):
                idx = i + j
                if idx < len(q_nums):
                    q = q_nums[idx]
                    with cols[j]:
                        choice = st.radio(f"{q}번", ["X", "O"], horizontal=True, key=f"q_{q}")
                        ans[str(q)] = 1 if choice == "O" else 0
        
        if st.form_submit_button("저장"):
            doc = get_google_sheet()
            ws = doc.worksheet('Student_Results')
            new_row = [selected_test, name, school, grade] + [ans[str(q)] for q in q_nums]
            ws.append_row(new_row)
            st.success("✅ 저장 완료!"); st.cache_data.clear()

with tab2:
    target = st.text_input("학생 이름 입력")
    if st.button("리포트 생성", type="primary"):
        found, buf, msg = generate_report(target, selected_test)
        if found: st.download_button("📥 PDF 다운로드", buf.getvalue(), f"{target}.pdf")
        else: st.error(msg)
