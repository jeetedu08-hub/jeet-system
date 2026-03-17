import streamlit as st
import pandas as pd
import json
import gspread

# 1. 페이지 기본 설정
st.set_page_config(page_title="JEET 통합 관리 시스템", layout="wide")
st.title("JEET 죽전캠퍼스 성적 통합 관리 시스템 📊")

SHEET_URL = "https://docs.google.com/spreadsheets/d/1bYv3ff5xwzd4DS3EZUC9Xj6GSpeVmijobbW0svKpqXU/edit"

# ==========================================
# 🚨 철통보안 출입증 확인 (원장님이 저장하신 Secrets 위치 자동 추적!)
# ==========================================
def get_credentials():
    if "GOOGLE_JSON" in st.secrets:
        return st.secrets["GOOGLE_JSON"]
    elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        return st.secrets["connections"]["gsheets"].get("credentials", "")
    return None

# ==========================================
# 🚀 핵심 엔진(gspread)으로 데이터 불러오기 (캐시 적용으로 초고속!)
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    raw_json = get_credentials()
    if not raw_json:
        st.error("비밀 금고(Secrets)에 출입증 세팅이 안 되어 있습니다.")
        st.stop()
        
    # 출입증 들고 구글 시트 당당하게 입장
    creds_dict = json.loads(raw_json)
    gc = gspread.service_account_from_dict(creds_dict)
    sh = gc.open_by_url(SHEET_URL)
    
    # 시트 2개 열기
    df_info = pd.DataFrame(sh.worksheet("Test_Info").get_all_records())
    df_results = pd.DataFrame(sh.worksheet("Student_Results").get_all_records())
    
    return df_info, df_results

# 데이터 세팅
try:
    df_info, df_results = load_data()
except Exception as e:
    st.error(f"구글 시트 연결 중 오류가 발생했습니다: {e}")
    st.stop()

# ==========================================
# 🌟 사이드바: 시험 과정 선택
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
                q_num = str(df_info_filtered.iloc[i]['문항번호']) if '문항번호' in df_info_filtered.columns else str(i+1)
                score = st.text_input(f"{q_num}번 문항", key=f"q_{i}")
                scores.append(score)
                
        submitted = st.form_submit_button("성적 구글시트에 저장하기")
        
        if submitted:
            if not student_name:
                st.warning("학생 이름을 입력해주세요!")
            else:
                new_row = [selected_test, student_name, school, grade] + scores
                
                # 빈칸 개수 맞추기
                columns = df_results.columns.tolist()
                while len(new_row) < len(columns):
                    new_row.append("")
                
                # 💡 gspread 엔진을 이용해 맨 아랫줄에 바로 점수 꽂아넣기!
                raw_json = get_credentials()
                creds_dict = json.loads(raw_json)
                gc = gspread.service_account_from_dict(creds_dict)
                sh = gc.open_by_url(SHEET_URL)
                ws_results = sh.worksheet("Student_Results")
                
                ws_results.append_row(new_row)
                
                st.success(f"🎉 {student_name} 학생의 [{selected_test}] 성적이 성공적으로 저장되었습니다!")
                load_data.clear() # 다음 검색 시 새 데이터가 반영되도록 캐시 지우기

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
