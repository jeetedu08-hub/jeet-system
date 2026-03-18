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
        return gspread.authorize(creds).open_by_url("https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit")
        
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

# --- 3. PDF 생성 함수 (심층 분석 및 동적 제목 포함) ---
def generate_jeet_expert_report(target_name, selected_test):
    try:
        _, _, _, df_info_all, df_results_all = load_data()
        
        # 선택된 시험에 맞춰 데이터 필터링
        df_info = df_info_all[df_info_all['시험명'] == selected_test].copy()
        df_results = df_results_all[df_results_all['시험명'] == selected_test].copy()
        
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
        pdf_buffer = io.BytesIO()

        for _, s_row in df_results.iterrows():
            student_name = str(s_row.get('이름', '')).strip()
            if not student_name or student_name == '0' or student_name != str(target_name).strip():
                continue
                
            student_found = True
            student_grade = s_row.get('학년', '')
            
            analysis = df_info.copy()
            analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
            analysis['득점'] = analysis['정답여부'] * analysis['배점']
            
            cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
            unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'})
            unit_data = unit_data.reindex([u for u in unit_order if u in unit_data.index])
  
            with PdfPages(pdf_buffer) as pdf:
                fig = plt.figure(figsize=(8.27, 11.69))
                
                # 1. 테두리 및 제목
                border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure, zorder=10)
                fig.patches.append(border)
                fig.text(0.31, 0.88, 'JEET', fontsize=42, fontweight='bold', color=COLOR_RED, ha='right')
                fig.text(0.33, 0.88, f'{selected_test} 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                
                info_text = f"과정: {selected_test}  |  학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')
  
                # 2. 방사형 차트 (영역별 역량)
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
                ax1.set_ylim(0, 110); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([])
                
                for i in range(len(labels)):
                    angle = angles[i]; label_text = labels[i]
                    ax1.text(angle, 120, label_text, fontsize=10, fontweight='bold', va='center', ha='center', color=COLOR_NAVY)
                    s_val = int(s_vals[i]); a_val = int(a_vals[i])
                    txt_s = ax1.text(angle, s_val+10, f"{s_val}%", fontsize=9, fontweight='bold', color=COLOR_STUDENT, ha='right')
                    txt_a = ax1.text(angle, s_val+10, f"({a_val}%)", fontsize=9, fontweight='bold', color=COLOR_RED, ha='left')
                    for t in [txt_s, txt_a]: t.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
                
                ax1.set_title("▶ 영역별 핵심 역량 지표 (%)", pad=30, fontsize=14, fontweight='bold', color=COLOR_NAVY)
  
                # 3. 막대 차트 (단원별 성취도)
                ax2 = fig.add_axes([0.55, 0.52, 0.35, 0.20])
                x_pos = np.arange(len(unit_data))
                bars = ax2.bar(x_pos, unit_data['득점'], color=COLOR_STUDENT, alpha=0.8, width=0.5, zorder=3)
                ax2.scatter(x_pos, unit_avg_data['평균득점'], color=COLOR_RED, marker='_', s=1000, linewidth=3, zorder=4)
                ax2.set_xticks(x_pos)
                ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold')
                ax2.set_title("▶ 단원별 성취도", pad=15, fontsize=14, fontweight='bold', color=COLOR_NAVY)
                
                for i, bar in enumerate(bars):
                    s_v = int(bar.get_height()); a_v = int(unit_avg_data['평균득점'].iloc[i])
                    ax2.text(bar.get_x() + bar.get_width()/2, s_v + 0.5, f"{s_v}({a_v})", ha='center', va='bottom', fontsize=9, fontweight='bold')
  
                # 4. 심층 분석 텍스트 섹션
                rect_diag = plt.Rectangle((0.08, 0.15), 0.84, 0.32, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                fig.text(0.13, 0.44, " JEET", fontsize=15, fontweight='bold', color=COLOR_RED)
                fig.text(0.185, 0.44, f"   {selected_test} {student_name} 학생 심층 분석", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                
                # 진단 로직
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
                
                solution_dict = {'계산력': "연산 훈련 병행", '이해력': "백지 복습법 권장", '추론력': "심화 문제 단계별 분석", '문제해결력': "융합형 문제 훈련"}
                worst_solution = solution_dict.get(worst_cat, "취약 단원 집중 공략 권장")

                diag_content = (
                    f"1. 종합 진단: {student_name} 학생은 전체 평균({total_avg_val}%) 대비 {abs(diff_val)}%p {'높은' if diff_val >= 0 else '낮은'} {avg_val}%의 성적을 기록하여 [{eval_tier}]를 보이고 있습니다.\n\n"
                    f"2. 강약점 분석: 영역별 진단 결과 '{best_cat}' 역량이 가장 돋보이나, 상대적으로 '{worst_cat}' 역량의 보완이 시급합니다. 특히 '{worst_unit}' 파트의 오답 재점검이 우선되어야 합니다.\n\n"
                    f"3. JEET 맞춤 솔루션: 단기적으로는 '{worst_unit}' 단원의 오답 노트를 작성하고, 중장기적으로 '{worst_cat}' 역량 강화를 위해 {worst_solution}을 권장합니다."
                )
                
                wrapped_lines = [textwrap.fill(p, width=54) for p in diag_content.split('\n\n')]
                fig.text(0.11, 0.41, "\n\n".join(wrapped_lines), fontsize=10.5, linespacing=1.8, va='top', color='#333')
  
                # 푸터 (캠퍼스 정보)
                fig.text(0.5, 0.08, "죽전 캠퍼스: 263-8003 | 기흥구 죽현로 29", ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
  
                pdf.savefig(fig)
                plt.close(fig)
                break
            
        return True, pdf_buffer, "성공"
    except Exception as e:
        return False, None, traceback.format_exc()

# ==========================================
# --- 4. Streamlit 웹 UI ---
# ==========================================
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide", page_icon="📊")

# 상단 레이아웃 (로고 및 타이틀)
c1, c2 = st.columns([8, 2])
with c1: st.title("📊 JEET 죽전캠퍼스 성적 통합 관리 시스템")
with c2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

# 데이터 로드
try:
    doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
except Exception as e:
    st.error(f"데이터 로드 실패: {e}"); st.stop()

# --- [중요] 사이드바 선택 ---
st.sidebar.header("📚 시험 과정 선택")
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)
st.sidebar.success(f"✅ 현재 [ {selected_test} ] 모드")

# 메인 탭
tab1, tab2 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력"])

with tab1:
    # 제목 동적 변경
    st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
    # ... (입력 폼 생략 방지 - 위 코드와 동일하게 작동)
    df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]
    q_nums = df_info_filtered['문항번호'].tolist()
    if q_nums:
        with st.form("input_form"):
            col1, col2, col3 = st.columns(3)
            i_name = col1.text_input("이름")
            i_school = col2.text_input("학교")
            i_grade = col3.selectbox("학년", ["중1", "중2", "중3"])
            
            st.write("문항별 정답(1:정답, 0:오답)")
            cols = st.columns(5); ans = {}
            for i, q in enumerate(q_nums):
                with cols[i % 5]: ans[str(q)] = st.number_input(f"{q}번", 0, 1, 0)
            
            if st.form_submit_button("저장"):
                # 저장 로직 (selected_test 포함)
                new_row = [selected_test, i_name, i_school, i_grade] + list(ans.values())
                ws_results.append_row(new_row)
                st.success("저장 완료"); fetch_all_dataframes.clear()

with tab2:
    # 🌟 [요청하신 부분] 사이드바 선택에 따라 이 제목이 바뀝니다!
    st.subheader(f"[{selected_test}] 개별 심층 분석 리포트 생성")
    
    target_name = st.text_input("학생 이름 입력:")
    if st.button("PDF 리포트 생성", type="primary"):
        if not target_name: st.warning("이름을 입력하세요.")
        else:
            with st.spinner("분석 중..."):
                success, buffer, msg = generate_jeet_expert_report(target_name, selected_test)
                if success:
                    st.success(f"{target_name} 학생 분석 완료!")
                    st.download_button("📥 PDF 다운로드", buffer.getvalue(), f"{target_name}_{selected_test}.pdf", "application/pdf")
                else: st.error(msg)
