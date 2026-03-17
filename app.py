import streamlit as st
import pandas as pd
import os
from streamlit_gsheets import GSheetsConnection

# 1. 페이지 기본 설정 (가장 위에 있어야 함)
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide")

st.title("JEET 죽전캠퍼스 성적 통합 관리 시스템 📊")

# ==========================================
# 🚨 [보안 출입증 장착] 비밀 금고에서 출입증 꺼내오기
# ==========================================
if "GOOGLE_JSON" in st.secrets:
    with open("google_creds.json", "w") as f:
        f.write(st.secrets["GOOGLE_JSON"])
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "google_creds.json"
else:
    st.error("스트림릿 Secrets에 출입증(GOOGLE_JSON)이 없습니다! 설정을 확인해주세요.")
    st.stop()

# 원장님의 구글 시트 주소
SHEET_URL = "https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit"

# 2. 구글 시트 데이터 불러오기 (출입증 들고 당당하게 입장!)
conn = st.connection("gsheets", type=GSheetsConnection)
df_info = conn.read(spreadsheet=SHEET_URL, worksheet="Test_Info")
df_results = conn.read(spreadsheet=SHEET_URL, worksheet="Student_Results")

# ==========================================
# 🌟 [핵심 기능] 사이드바: 시험 과정 선택
# ==========================================
st.sidebar.header("📚 시험 과정 선택")

if '시험명' not in df_info.columns:
    st.error("구글 시트 Test_Info 시트에 '시험명' 열이 없습니다! A열에 추가해주세요.")
    st.stop()

test_list = df_info['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)

df_info_filtered = df_info[df_info['시험명'] == selected_test]
df_results_filtered = df_results[df_results['시험명'] == selected_test]

st.sidebar.success(f"✅ 현재 [ {selected_test} ] 모드입니다.")
st.sidebar.markdown("---")

# ==========================================
# 🌟 화면 탭 구성
# ==========================================
tab1, tab2 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 학생 성적 입력")
    st.write("학생 정보와 문항별 O/X/△ 결과를 입력해주세요.")
    
    with st.form("score_input_form", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            student_name = st.text_input("학생 이름")
        with col2:
            school = st.text_input("학교")
        with col3:
            grade = st.text_input("학년")
            
        st.markdown("---")
        
        question_count = len(df_info_filtered)
        scores = []
        cols = st.columns(5)
        
        for i in range(question_count):
            with cols[i % 5]:
                q_num = df_info_filtered.iloc[i]['문항번호'] if '문항번호' in df_info_filtered.columns else i+1
                score = st.text_input(f"{q_num}번 문항", key=f"q_{i}")
                scores.append(score)
                
        submitted = st.form_submit_button("성적 구글시트에 저장하기")
        
        if submitted:
            if not student_name:
                st.warning("학생 이름을 입력해주세요!")
            else:
                new_row = [selected_test, student_name, school, grade] + scores
                
                columns = df_results.columns.tolist()
                while len(new_row) < len(columns):
                    new_row.append("")
                
                new_df = pd.DataFrame([new_row], columns=columns)
                updated_df = pd.concat([df_results, new_df], ignore_index=True)
                
                # 💡 입력(쓰기) 권한을 가지고 구글 시트에 성적 업데이트!
                conn.update(spreadsheet=SHEET_URL, worksheet="Student_Results", data=updated_df)
                
                st.success(f"🎉 {student_name} 학생의 [{selected_test}] 성적이 성공적으로 저장되었습니다!")
                st.cache_data.clear()

with tab2:
    st.subheader(f"[{selected_test}] 학생별 분석 리포트 출력")
    
    student_list = df_results_filtered['이름'].dropna().unique().tolist()
    
    if not student_list:
        st.info(f"아직 {selected_test} 성적이 입력된 학생이 없습니다.")
    else:
        search_name = st.selectbox("리포트를 확인할 학생을 선택하세요:", ["선택하세요"] + student_list)
        
        if search_name != "선택하세요":
            student_data = df_results_filtered[df_results_filtered['이름'] == search_name].iloc[0]
            
            st.markdown("---")
            st.markdown(f"## 📄 {search_name} 학생 종합 분석 리포트")
            st.write(f"**진행 과정:** {selected_test}")
            st.write(f"**학교/학년:** {student_data['학교']} / {student_data['학년']}")
            
            st.info("💡 모바일 또는 PC에서 이 화면을 캡처하거나, PC 브라우저 우클릭 -> [인쇄] -> [PDF로 저장]을 누르시면 학부모님 전송용 리포트가 완성됩니다!")
            
            st.write("### 📝 문항별 결과")
            
            score_data = student_data.iloc[4:] 
            q_numbers = df_info_filtered['문항번호'].tolist() if '문항번호' in df_info_filtered.columns else [str(i) for i in range(1, len(score_data)+1)]
            
            result_df = pd.DataFrame({
                "문항 번호": q_numbers[:len(score_data)],
                "학생 결과": score_data.values
            })
            st.dataframe(result_df.T)
