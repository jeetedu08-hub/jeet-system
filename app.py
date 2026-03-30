# --- 반별 일괄 리포트 생성 함수 추가 (선택 필터링 및 공백 제거 강력 대응) ---
def generate_batch_report(target_class, selected_test, selected_students=None):
    try:
        df_info, df_results, avg_cat_ratio, unit_avg_data, unit_order, safe_to_int = prepare_report_data(selected_test)
        
        # 1차 필터링: 입력받은 반과 일치하는 학생 (양옆 공백 완전 제거 후 비교)
        class_students = df_results[df_results['반'].astype(str).str.strip() == str(target_class).strip()]
        
        # 2차 필터링: 선택된 학생 명단이 있다면 해당 학생들만 남김 (이름의 양옆 공백도 완전 제거 후 비교)
        if selected_students is not None:
            # 선택된 학생 명단 공백 제거
            cleaned_selected = [str(s).strip() for s in selected_students]
            # 시트의 이름 데이터도 공백 제거 후 비교
            class_students = class_students[class_students['이름'].astype(str).str.strip().isin(cleaned_selected)]
        
        if class_students.empty:
            return False, None, f"선택된 학생 데이터가 없습니다. (이름 공백/오타 확인 필요)"
            
        pdf_buffer = io.BytesIO()
        
        with PdfPages(pdf_buffer) as pdf:
            for _, s_row in class_students.iterrows():
                # 여기서도 이름 공백 제거
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
