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
    if "GOOGLE_JSON" in st.secrets:
        creds_dict = json.loads(st.secrets["GOOGLE_JSON"])
    elif "gcp_secret_string" in st.secrets:
        creds_dict = json.loads(st.secrets["gcp_secret_string"])
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("secrets.json", scope)
        return gspread.authorize(creds).open_by_url("https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit")
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds).open_by_url("https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit")

@st.cache_data(ttl=60)
def fetch_all_dataframes():
    doc = get_google_sheet()
    df_info = pd.DataFrame(doc.worksheet('Test_Info').get_all_records())
    df_results = pd.DataFrame(doc.worksheet('Student_Results').get_all_records())
    return df_info, df_results

# --- 3. PDF 리포트 생성 함수 (심층 분석 전체 포함) ---
def generate_jeet_expert_report(target_name, selected_test):
    try:
        df_info_all, df_results_all = fetch_all_dataframes()
        df_info = df_info_all[df_info_all['시험명'] == selected_test].copy()
        df_results = df_results_all[df_results_all['시험명'] == selected_test].copy()
        
        df_results = df_results.replace('', 0).fillna(0)
        df_results.columns = df_results.columns.astype(str)
        df_info['배점'] = pd.to_numeric(df_info['배점'], errors='coerce').fillna(3).astype(int)
        unit_order = df_info['단원'].drop_duplicates().tolist()
        
        def safe_to_int(val):
            try: return int(float(val))
            except: return 0

        # 전체 평균 계산
        q_cols = [str(q) for q in df_info['문항번호']]
        valid_cols = [col for col in df_results.columns if col in q_cols]
        df_scores = df_results[valid_cols].applymap(safe_to_int)
        avg_per_q = df_scores.mean()
        
        total_analysis = df_info.copy()
        total_analysis['평균득점'] = total_analysis['문항번호'].apply(lambda x: avg_per_q.get(str(x), 0))
        avg_cat_ratio = (total_analysis.groupby('영역')['평균득점'].sum() / total_analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
        unit_avg_data = total_analysis.groupby('단원').agg({'평균득점': 'sum'}).reindex(unit_order).fillna(0)

        pdf_buffer = io.BytesIO()
        student_found = False

        for _, s_row in df_results.iterrows():
            if str(s_row.get('이름', '')).strip() != str(target_name).strip(): continue
            student_found = True
            
            # 학생 개인 분석
            analysis = df_info.copy()
            analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
            analysis['득점'] = analysis['정답여부'] * analysis['배점']
            
            cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
            unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'}).reindex(unit_order).fillna(0)

            with PdfPages(pdf_buffer) as pdf:
                fig = plt.figure(figsize=(8.27, 11.69))
                border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure)
                fig.patches.append(border)
                
                # 헤더
                fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.88, f'{selected_test} 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY)
                info_text = f"과정: {selected_test}  |  학교: {s_row.get('학교', '0')}  |  학년: {s_row.get('학년', '')}  |  이름: {target_name}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold')

                # 1. 방사형 차트
                ax1 = fig.add_axes([0.10, 0.52, 0.32, 0.22], polar=True)
                labels = cat_ratio.index.tolist()
                s_vals = cat_ratio.values.tolist() + [cat_ratio.values[0]]
                a_vals = avg_cat_ratio.reindex(labels).values.tolist() + [avg_cat_ratio.iloc[0]]
                angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
                
                ax1.set_theta_direction(-1); ax1.set_theta_offset(np.pi/2.0)
                ax1.plot(angles, a_vals, color=COLOR_AVG, linewidth=1, linestyle='--')
                ax1.fill(angles, a_vals, color=COLOR_AVG, alpha=0.1)
                ax1.plot(angles, s_vals, color=COLOR_STUDENT, linewidth=2.5)
                ax1.set_xticks(angles[:-1]); ax1.set_xticklabels(labels, fontsize=10, fontweight='bold', color=COLOR_NAVY)
                ax1.set_ylim(0, 110)
                ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)

                # 2. 막대 차트 (단원별 성취도)
                ax2 = fig.add_axes([0.55, 0.52, 0.35, 0.20])
                x = np.arange(len(unit_order))
                ax2.bar(x, unit_data['득점'], color=COLOR_STUDENT, alpha=0.7, width=0.4)
                ax2.scatter(x, unit_avg_data['평균득점'], color=COLOR_RED, marker='_', s=500, linewidth=3)
                ax2.set_xticks(x)
                ax2.set_xticklabels([textwrap.fill(l, 6) for l in unit_order], fontsize=8)
                ax2.set_ylim(0, unit_data['배점'].max() + 1)
                ax2.set_title("▶ 단원별 성취도", pad=15, fontsize=14, fontweight='bold', color=COLOR_NAVY)
                for i, row in enumerate(unit_data.itertuples()):
                    ax2.text(i, 0.5, f"{int(row.득점)}({int(row.배점)})", ha='center', fontsize=9, fontweight='bold')

                # 3. 심층 분석 섹션 (이미지와 동일한 로직)
                fig.text(0.13, 0.44, f"JEET  {selected_test} {target_name} 학생 심층 분석", fontsize=16, fontweight='bold', color=COLOR_NAVY)
                rect = plt.Rectangle((0.08, 0.15), 0.84, 0.32, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect)
                
                avg_val = int(cat_ratio.mean())
                total_avg_val = int(avg_cat_ratio.mean())
                diff = avg_val - total_avg_val
                
                if avg_val >= 90: eval_tier = "최상위권 수준의 탁월한 성취도"
                elif avg_val >= 70: eval_tier = "기본기가 탄탄한 우수한 성취도"
                else: eval_tier = "전반적인 기초 학습 재점검이 요구되는 성취도"
                
                worst_cat = (cat_ratio - avg_cat_ratio).idxmin()
                sol_dict = {'계산력': "연산 훈련 병행", '문제해결력': "심화 문제 단계별 분석", '이해력': "백지 복습법"}
                
                diag_text = (
                    f"1. 종합 진단: {target_name} 학생은 전체 평균({total_avg_val}%) 대비 {abs(diff)}%p {'높은' if diff >= 0 else '낮은'} {avg_val}%의 성적을 기록하여 [{eval_tier}]를 보이고 있습니다.\n\n"
                    f"2. 강약점 분석: 영역별 진단 결과 '{cat_ratio.idxmax()}' 역량이 가장 돋보이나, 상대적으로 '{worst_cat}' 역량의 보완이 시급합니다. 특히 오답 재점검이 우선되어야 합니다.\n\n"
                    f"3. JEET 맞춤 솔루션: 단기적으로는 틀린 문항의 오답 노트를 작성하고, 중장기적으로 '{worst_cat}' 역량 강화를 위해 {sol_dict.get(worst_cat, '취약 부분 집중 공략')}을 권장합니다."
                )
                fig.text(0.11, 0.41, diag_text, fontsize=11, linespacing=2, va='top')
                fig.text(0.5, 0.08, "죽전 캠퍼스: 263-8003 | 기흥구 죽현로 29", ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)

                pdf.savefig(fig); plt.close(fig)
                break
        
        return student_found, pdf_buffer, "성공"
    except Exception as e:
        return False, None, traceback.format_exc()

# --- 4. Streamlit UI ---
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide")
df_info_all, df_results_all = fetch_all_dataframes()
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("시험 선택", test_list)

tab1, tab2 = st.tabs(["📝 성적 입력", "📑 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 성적 입력")
    q_nums = df_info_all[df_info_all['시험명'] == selected_test]['문항번호'].tolist()
    
    with st.form("input_form"):
        c1, c2, c3 = st.columns(3)
        i_name = c1.text_input("이름")
        i_school = c2.text_input("학교")
        i_grade = c3.selectbox("학년", ["중1", "중2", "중3"])
        
        st.write("---")
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
            ws_results = doc.worksheet('Student_Results')
            new_row = [selected_test, i_name, i_school, i_grade] + [ans[str(q)] for q in q_nums]
            ws_results.append_row(new_row)
            st.success("저장되었습니다!"); st.cache_data.clear()

with tab2:
    st.subheader("리포트 생성")
    target_name = st.text_input("학생 이름")
    if st.button("PDF 리포트 생성", type="primary"):
        success, buffer, msg = generate_jeet_expert_report(target_name, selected_test)
        if success:
            st.download_button("📥 PDF 다운로드", buffer.getvalue(), f"{target_name}.pdf")
        else: st.error(msg)
