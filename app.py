# 💡 [tab4] 영역을 이 코드로 통째로 덮어씌워 주세요.
with tab4:
    st.subheader("📚 분기별 재원생 성적 데이터 엑셀 내보내기")
    st.markdown("대시보드 상단의 시험 과정과 관계없이, 선택하신 **분기**에서 구분이 **'재원생'**인 모든 학생들의 성적 데이터를 통합하여 엑셀로 추출합니다.")
    
    if not df_results_all.empty and '분기' in df_results_all.columns:
        quarter_options = sorted(df_results_all['분기'].dropna().astype(str).unique().tolist())
    else:
        quarter_options = ["2026년 1분기", "2026년 2분기", "2026년 3분기", "2026년 4분기"]
        
    excel_quarter = st.selectbox("📥 내보낼 분기를 선택하세요:", quarter_options, key="excel_quarter_select")
    
    if st.button("📊 해당 분기 재원생 통합 엑셀 파일 생성하기", type="primary"):
        # 1. [수정] 시험명 필터를 제거하고 '분기'와 '재원생' 조건만으로 필터링합니다.
        filtered_df = df_results_all[
            (df_results_all['분기'].astype(str).str.strip() == excel_quarter.strip()) & 
            (df_results_all['구분'].astype(str).str.strip() == '재원생')
        ].copy()
        
        if filtered_df.empty:
            st.warning(f"⚠ 선택하신 [{excel_quarter}] 분기에는 '재원생' 데이터가 존재하지 않습니다.")
        else:
            # 2. [수정] 해당 분기 내의 학생들이 치른 모든 시험지의 실제 문항 번호를 test_info에서 수집합니다.
            # (여러 시험지가 섞여 있어도 문항 수가 유연하게 35번 등 실제 시험 범위까지만 나오도록 처리)
            distinct_tests = filtered_df['시험명'].dropna().unique().tolist()
            df_info_filtered = df_info_all[df_info_all['시험명'].isin(distinct_tests)]
            
            def clean_info_q(q):
                nums = re.findall(r'\d+', str(q).split('.')[0])
                return nums[0] if nums else str(q).strip()
                
            if not df_info_filtered.empty:
                actual_q_cols = df_info_filtered['문항번호'].apply(clean_info_q).unique().tolist()
                actual_q_cols = sorted(actual_q_cols, key=lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else x)
            else:
                # 만약 test_info 매핑이 안 된다면 데이터프레임 내의 1~35번 내외의 문항을 동적으로 추적
                actual_q_cols = [col for col in filtered_df.columns if col.isdigit() and int(col) <= 50]
                actual_q_cols = sorted(actual_q_cols, key=lambda x: int(x))
                
            if '반' in filtered_df.columns:
                filtered_df = filtered_df.sort_values(by=['반', '이름'] if '이름' in filtered_df.columns else ['반'])
                
            st.success(f"🎯 [{excel_quarter}]의 모든 시험 과정에서 총 {len(filtered_df)}명의 재원생 데이터가 통합 확인되었습니다.")
            
            # 3. 스타일에 맞게 엑셀 파일 생성 (불필요한 36~50번 컬럼 자동 차단)
            with st.spinner("통합 엑셀 시트 스타일 마스터링 중..."):
                excel_file = export_excel_styled(filtered_df, excel_quarter, actual_q_cols)
                
            safe_filename = f"JEET_통합_{excel_quarter}_재원생성적.xlsx".replace(" ", "_")
            
            st.download_button(
                label="📥 깔끔한 통합 엑셀 파일(.xlsx) 다운로드",
                data=excel_file.getvalue(),
                file_name=safe_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
