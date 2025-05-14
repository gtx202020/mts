import sqlite3
import pandas as pd
import sys
import os

# --- 설정 변수 ---
db_filename = 'info.sqlite' # SQLite DB 파일 (df_filtered를 가져오기 위함)
table_name = 'list'         # DB 테이블 이름

# 실제 컬럼명으로 수정해야 합니다.
column_b_name = '컬럼B'
column_c_name = '컬럼C'
# !!! 중요: '컬럼D'의 실제 이름을 아래 변수에 지정해주세요. !!!
column_d_name = '컬럼D'  # 예: 'D_Data_Column' 등 실제 엑셀/DB의 컬럼명

# 검색 및 대체 문자열 정의
val_ly = 'LY'
val_lz = 'LZ'
replace_ly_with = 'LH' # LY를 LH로 대체
replace_lz_with = 'VO' # LZ를 VO로 대체
# -----------------

# 1. 출력 Excel 파일 이름 설정
try:
    script_basename = os.path.basename(sys.argv[0])
    script_name_without_ext = os.path.splitext(script_basename)[0]
    # 원본 스크립트 이름에 "_processed"를 추가하여 구분
    excel_filename = f"{script_name_without_ext}_reordered.xlsx"
except Exception:
    excel_filename = "output_reordered.xlsx"
    print(f"스크립트 이름을 감지할 수 없어 기본 파일명 '{excel_filename}'을 사용합니다.")

df_filtered = pd.DataFrame() # 초기화

# --- 이전 단계에서 생성된 df_filtered를 가져오는 부분 ---
# 이 코드가 독립적으로 실행될 수 있도록, df_filtered를 DB에서 다시 생성하는 로직을 포함합니다.
# 만약 다른 스크립트에서 df_filtered를 이미 DataFrame 객체로 가지고 있다면,
# 이 DB 조회 부분은 해당 DataFrame을 사용하는 코드로 대체되어야 합니다.
try:
    conn = sqlite3.connect(db_filename)
    cursor = conn.cursor()

    # df_filtered 생성 쿼리: 컬럼B 또는 컬럼C에 'LY' 또는 'LZ'가 포함된 행
    initial_filter_query = f"""
    SELECT *
    FROM "{table_name}"
    WHERE
        ("{column_b_name}" LIKE ? OR "{column_b_name}" LIKE ?)
        OR
        ("{column_c_name}" LIKE ? OR "{column_c_name}" LIKE ?)
    """
    initial_params = (
        f'%{val_ly}%', f'%{val_lz}%',
        f'%{val_ly}%', f'%{val_lz}%'
    )
    cursor.execute(initial_filter_query, initial_params)
    filtered_rows_from_db = cursor.fetchall()

    if filtered_rows_from_db:
        column_names_from_db = [description[0] for description in cursor.description]
        df_filtered = pd.DataFrame(filtered_rows_from_db, columns=column_names_from_db)
        print(f"총 {len(df_filtered)}개의 행으로 df_filtered가 생성되었습니다.")
    else:
        print("초기 필터링 조건에 맞는 데이터가 DB에 없습니다. 이후 처리를 진행할 수 없습니다.")
        # df_filtered는 빈 상태로 유지됩니다.

except sqlite3.Error as e:
    print(f"SQLite 오류 발생: {e}")
except FileNotFoundError:
    print(f"오류: 데이터베이스 파일 '{db_filename}'을 찾을 수 없습니다.")
except Exception as e:
    print(f"데이터 조회 중 예상치 못한 오류 발생: {e}")
finally:
    if 'conn' in locals() and conn:
        conn.close()
# --- df_filtered 준비 완료 ---


# --- 새로운 순회 및 필터링, 재정렬 로직 ---
output_rows_info = [] # 최종 DataFrame을 구성할 행 정보 리스트 (데이터와 스타일 플래그 포함)

if not df_filtered.empty:
    # 필수 컬럼 존재 여부 확인
    required_cols = [column_d_name, column_b_name, column_c_name]
    if not all(col in df_filtered.columns for col in required_cols):
        missing_cols = [col for col in required_cols if col not in df_filtered.columns]
        print(f"오류: 필수 컬럼 {missing_cols}이(가) df_filtered에 없습니다. 컬럼명을 확인하세요.")
        df_filtered = pd.DataFrame() # 오류 발생 시 처리 중단

if not df_filtered.empty:
    print("조건에 따라 행 재정렬 및 삽입 작업을 시작합니다...")
    for i, current_row in df_filtered.iterrows():
        # 원본 행 추가 (노란색 아님)
        output_rows_info.append({'data_row': current_row.copy(), 'make_yellow': False})

        # 현재 행의 컬럼 값 (문자열로 변환 및 NaN 처리)
        current_d_val = str(current_row[column_d_name]) if pd.notna(current_row[column_d_name]) else ""
        current_d_val_stripped = current_d_val.strip()
        
        current_b_val = str(current_row[column_b_name]) if pd.notna(current_row[column_b_name]) else ""
        current_c_val = str(current_row[column_c_name]) if pd.notna(current_row[column_c_name]) else ""

        # df_filtered의 다른 행들과 비교하여 조건에 맞는 행 찾기
        for j, target_row in df_filtered.iterrows():
            if i == j: # 같은 행은 비교 대상에서 제외
                continue

            target_d_val = str(target_row[column_d_name]) if pd.notna(target_row[column_d_name]) else ""
            target_d_val_stripped = target_d_val.strip()

            # 1차 필터링: 컬럼D의 strip() 결과가 같아야 함
            cond1_match = (current_d_val_stripped == target_d_val_stripped)
            
            if not cond1_match:
                continue

            # 조건 1을 만족한 경우, 추가 조건 확인
            target_b_val = str(target_row[column_b_name]) if pd.notna(target_row[column_b_name]) else ""
            target_c_val = str(target_row[column_c_name]) if pd.notna(target_row[column_c_name]) else ""

            # 2차 필터링: 컬럼B 조건 확인
            cond2_b_match_for_target = False
            if val_ly in current_b_val: # 현재 행의 컬럼B에 'LY'가 있고
                transformed_b_for_ly = current_b_val.replace(val_ly, replace_ly_with)
                if target_b_val == transformed_b_for_ly: # 타겟 행의 컬럼B가 변환된 문자열과 같다면
                    cond2_b_match_for_target = True
            
            if not cond2_b_match_for_target and (val_lz in current_b_val): # 또는 현재 행의 컬럼B에 'LZ'가 있고
                transformed_b_for_lz = current_b_val.replace(val_lz, replace_lz_with)
                if target_b_val == transformed_b_for_lz: # 타겟 행의 컬럼B가 변환된 문자열과 같다면
                    cond2_b_match_for_target = True

            # 3차 필터링: 컬럼C 조건 확인
            cond3_c_match_for_target = False
            if val_ly in current_c_val: # 현재 행의 컬럼C에 'LY'가 있고
                transformed_c_for_ly = current_c_val.replace(val_ly, replace_ly_with)
                if target_c_val == transformed_c_for_ly: # 타겟 행의 컬럼C가 변환된 문자열과 같다면
                    cond3_c_match_for_target = True
            
            if not cond3_c_match_for_target and (val_lz in current_c_val): # 또는 현재 행의 컬럼C에 'LZ'가 있고
                transformed_c_for_lz = current_c_val.replace(val_lz, replace_lz_with)
                if target_c_val == transformed_c_for_lz: # 타겟 행의 컬럼C가 변환된 문자열과 같다면
                    cond3_c_match_for_target = True
            
            # 1차 필터링을 통과하고, (2차 필터링 통과 OR 3차 필터링 통과)한 경우
            if cond2_b_match_for_target or cond3_c_match_for_target:
                output_rows_info.append({'data_row': target_row.copy(), 'make_yellow': True})
    print("행 재정렬 및 삽입 작업 완료.")

# 최종 DataFrame 생성
if output_rows_info:
    final_df_data = [item['data_row'] for item in output_rows_info]
    df_excel_output = pd.DataFrame(final_df_data).reset_index(drop=True)
    # 삽입된 행(노란색으로 칠할 행)의 인덱스 리스트 생성 (df_excel_output 기준)
    yellow_row_indices_in_final_df = [idx for idx, item in enumerate(output_rows_info) if item['make_yellow']]
else:
    df_excel_output = pd.DataFrame() # 빈 DataFrame
    yellow_row_indices_in_final_df = []


# --- DataFrame을 Excel 파일로 저장하고 스타일 적용 ---
if not df_excel_output.empty:
    try:
        with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
            df_excel_output.to_excel(writer, sheet_name='ProcessedData', index=False)

            workbook = writer.book
            worksheet = writer.sheets['ProcessedData']

            # 노란색 배경 서식 정의
            yellow_format = workbook.add_format({'bg_color': '#FFFF00'}) # 노란색

            # 다른 특정 색상 서식은 적용하지 않음 (헤더는 기본 스타일 유지)
            # Pandas to_excel이 기본적으로 헤더를 쓰며, xlsxwriter 엔진은 헤더를 굵게 표시합니다.
            # "색상"을 없애는 것이므로, 기본 굵은 헤더는 유지됩니다.

            # 삽입된 행(make_yellow=True)에 노란색 배경 적용
            if yellow_row_indices_in_final_df:
                for zero_based_row_idx in yellow_row_indices_in_final_df:
                    # Excel에서 데이터 행은 헤더 다음부터 시작하므로 인덱스 +1
                    # set_row(row, height, cell_format, options)
                    worksheet.set_row(zero_based_row_idx + 1, None, yellow_format)
            
            # 컬럼 너비 자동 조절 (가독성을 위해)
            for i, col_name_str in enumerate(df_excel_output.columns.astype(str)):
                # 데이터 중 가장 긴 문자열 길이 + 헤더 문자열 길이 중 큰 값을 기준으로 설정
                # NaN 값 등으로 인해 data_max_len이 float가 될 수 있으므로 int로 변환
                data_max_len_series = df_excel_output[col_name_str].astype(str).map(len)
                data_max_len = data_max_len_series.max() if not data_max_len_series.empty else 0

                header_len = len(col_name_str)
                if pd.isna(data_max_len): data_max_len = 0
                
                column_width = max(int(data_max_len), header_len) + 2 # 약간의 여유 공간
                worksheet.set_column(i, i, column_width)

        print(f"\n결과가 '{excel_filename}' 파일로 저장되었습니다.")
        print("조건에 따라 삽입된 행들은 노란색으로 표시됩니다. 그 외 다른 특정 배경색은 적용되지 않았습니다.")

    except ImportError:
        print("Excel 파일 저장을 위해 'xlsxwriter' 라이브러리가 필요합니다. 'pip install xlsxwriter' 명령어로 설치해주세요.")
    except Exception as e_excel:
        print(f"Excel 파일 저장 중 오류 발생: {e_excel}")

elif df_filtered.empty: # 초기 df_filtered가 비어있던 경우
    print("처리할 원본 데이터(df_filtered)가 없어 Excel 파일을 생성하지 않았습니다.")
else: # output_rows_info가 비어있는 경우 (재정렬 후 최종 데이터가 없는 경우)
    print("조건에 맞는 데이터가 없어 최종적으로 Excel 파일에 저장할 내용이 없습니다.")


