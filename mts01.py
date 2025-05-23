import sqlite3
import pandas as pd
import sys
import os

# --- 설정 변수 ---
db_filename = 'info.sqlite'
table_name = 'list'

column_b_name = '컬럼B'
column_c_name = '컬럼C'
# !!! 중요: '컬럼D'의 실제 이름을 아래 변수에 지정해주세요. !!!
column_d_name = '컬럼D'  # 예: 'D_Data_Column' 등 실제 엑셀/DB의 컬럼명

val_ly = 'LY'
val_lz = 'LZ'
replace_ly_with = 'LH'
replace_lz_with = 'VO'
# -----------------

# 1. 출력 Excel 파일 이름 설정
try:
    script_basename = os.path.basename(sys.argv[0])
    script_name_without_ext = os.path.splitext(script_basename)[0]
    excel_filename = f"{script_name_without_ext}_reordered_v2.xlsx" # 버전 표시
except Exception:
    excel_filename = "output_reordered_v2.xlsx"
    print(f"스크립트 이름을 감지할 수 없어 기본 파일명 '{excel_filename}'을 사용합니다.")

df_complete_table = pd.DataFrame() # 원본 전체 테이블
df_filtered = pd.DataFrame()       # 초기 필터링된 테이블

# --- DB에서 전체 데이터 로드 및 df_filtered 생성 ---
try:
    conn = sqlite3.connect(db_filename)
    cursor = conn.cursor()

    # 1. DB에서 전체 데이터 로드
    cursor.execute(f'SELECT * FROM "{table_name}"')
    all_rows_from_db = cursor.fetchall()

    if not all_rows_from_db:
        print(f"원본 테이블 '{table_name}'에 데이터가 없습니다. 처리를 중단합니다.")
    else:
        column_names_from_db = [description[0] for description in cursor.description]
        df_complete_table = pd.DataFrame(all_rows_from_db, columns=column_names_from_db)
        print(f"원본 전체 테이블에 총 {len(df_complete_table)}개의 행이 로드되었습니다.")

        # 2. df_filtered 생성: 컬럼B 또는 컬럼C에 'LY' 또는 'LZ' 포함 조건
        # 문자열로 안전하게 변환 후 'contains' 사용 (NaN은 False로 처리)
        cond_b_contains = (
            df_complete_table[column_b_name].astype(str).str.contains(val_ly, na=False) |
            df_complete_table[column_b_name].astype(str).str.contains(val_lz, na=False)
        )
        cond_c_contains = (
            df_complete_table[column_c_name].astype(str).str.contains(val_ly, na=False) |
            df_complete_table[column_c_name].astype(str).str.contains(val_lz, na=False)
        )
        df_filtered = df_complete_table[cond_b_contains | cond_c_contains].copy() # 중요: .copy()로 사본 생성
        
        if df_filtered.empty:
            print("초기 필터링 조건('LY'/'LZ' 포함)에 맞는 데이터가 없습니다. 후속 처리를 진행할 수 없습니다.")
        else:
            print(f"초기 필터링 후 df_filtered에 {len(df_filtered)}개의 행이 남았습니다.")

except sqlite3.Error as e:
    print(f"SQLite 오류 발생: {e}")
except FileNotFoundError:
    print(f"오류: 데이터베이스 파일 '{db_filename}'을 찾을 수 없습니다.")
except Exception as e:
    print(f"데이터 조회 또는 초기 필터링 중 예상치 못한 오류 발생: {e}")
finally:
    if 'conn' in locals() and conn:
        conn.close()
# --- 데이터 준비 완료 ---

output_rows_info = []

if not df_filtered.empty and not df_complete_table.empty:
    # 필수 컬럼 존재 여부 확인 (df_filtered와 df_complete_table 모두에 필요)
    required_cols = [column_d_name, column_b_name, column_c_name]
    if not all(col in df_filtered.columns for col in required_cols) or \
       not all(col in df_complete_table.columns for col in required_cols):
        missing_cols_filtered = [col for col in required_cols if col not in df_filtered.columns]
        missing_cols_complete = [col for col in required_cols if col not in df_complete_table.columns]
        if missing_cols_filtered: print(f"오류: 필수 컬럼 {missing_cols_filtered}이(가) df_filtered에 없습니다.")
        if missing_cols_complete: print(f"오류: 필수 컬럼 {missing_cols_complete}이(가) df_complete_table에 없습니다.")
        print("컬럼명을 확인하세요. 처리를 중단합니다.")
        df_filtered = pd.DataFrame() # 처리를 중단하기 위해 비움

if not df_filtered.empty and not df_complete_table.empty:
    print("조건에 따라 행 재정렬 및 삽입 작업을 시작합니다 (비교 대상: 원본 전체 테이블)...")
    for idx_filtered, current_row in df_filtered.iterrows(): # current_row는 초기 필터링된 결과
        output_rows_info.append({'data_row': current_row.copy(), 'make_yellow': False})

        current_d_val = str(current_row[column_d_name]) if pd.notna(current_row[column_d_name]) else ""
        current_d_val_stripped = current_d_val.strip()
        
        current_b_val = str(current_row[column_b_name]) if pd.notna(current_row[column_b_name]) else ""
        current_c_val = str(current_row[column_c_name]) if pd.notna(current_row[column_c_name]) else ""

        # 비교 대상은 원본 전체 테이블 (df_complete_table)
        for idx_complete, target_row in df_complete_table.iterrows():
            # current_row와 target_row가 원본 테이블에서 같은 행인지 확인 후 건너뛰기
            # df_filtered는 df_complete_table의 부분집합이므로, 인덱스(current_row.name)를 사용해 비교 가능
            if current_row.name == target_row.name:
                continue

            target_d_val = str(target_row[column_d_name]) if pd.notna(target_row[column_d_name]) else ""
            target_d_val_stripped = target_d_val.strip()

            # 1차 필터링: 컬럼D의 strip() 결과 비교
            cond1_match = (current_d_val_stripped == target_d_val_stripped)
            
            if not cond1_match:
                continue

            target_b_val = str(target_row[column_b_name]) if pd.notna(target_row[column_b_name]) else ""
            target_c_val = str(target_row[column_c_name]) if pd.notna(target_row[column_c_name]) else ""

            # 2차 필터링: 컬럼B 조건
            cond2_b_match_for_target = False
            if val_ly in current_b_val:
                transformed_b_for_ly = current_b_val.replace(val_ly, replace_ly_with)
                if target_b_val == transformed_b_for_ly:
                    cond2_b_match_for_target = True
            if not cond2_b_match_for_target and (val_lz in current_b_val):
                transformed_b_for_lz = current_b_val.replace(val_lz, replace_lz_with)
                if target_b_val == transformed_b_for_lz:
                    cond2_b_match_for_target = True

            # 3차 필터링: 컬럼C 조건
            cond3_c_match_for_target = False
            if val_ly in current_c_val:
                transformed_c_for_ly = current_c_val.replace(val_ly, replace_ly_with)
                if target_c_val == transformed_c_for_ly:
                    cond3_c_match_for_target = True
            if not cond3_c_match_for_target and (val_lz in current_c_val):
                transformed_c_for_lz = current_c_val.replace(val_lz, replace_lz_with)
                if target_c_val == transformed_c_for_lz:
                    cond3_c_match_for_target = True
            
            if cond2_b_match_for_target or cond3_c_match_for_target:
                output_rows_info.append({'data_row': target_row.copy(), 'make_yellow': True})
    print("행 재정렬 및 삽입 작업 완료.")

# 최종 DataFrame 생성 (이전 코드와 동일)
if output_rows_info:
    final_df_data = [item['data_row'] for item in output_rows_info]
    # DataFrame 생성 시 컬럼 순서 유지를 위해 df_complete_table의 컬럼 사용 (데이터가 있을 경우)
    # 또는 df_filtered의 컬럼 사용 (output_rows_info의 첫번째 요소 기준도 가능)
    cols_for_final_df = column_names_from_db if 'column_names_from_db' in locals() and column_names_from_db else (df_filtered.columns if not df_filtered.empty else None)
    if cols_for_final_df is not None:
         df_excel_output = pd.DataFrame(final_df_data, columns=cols_for_final_df).reset_index(drop=True)
    else: # 비상시
         df_excel_output = pd.DataFrame(final_df_data).reset_index(drop=True)

    yellow_row_indices_in_final_df = [idx for idx, item in enumerate(output_rows_info) if item['make_yellow']]
else:
    df_excel_output = pd.DataFrame()
    yellow_row_indices_in_final_df = []

# --- DataFrame을 Excel 파일로 저장하고 스타일 적용 (이전 코드와 동일) ---
if not df_excel_output.empty:
    try:
        with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
            df_excel_output.to_excel(writer, sheet_name='ProcessedData', index=False)

            workbook = writer.book
            worksheet = writer.sheets['ProcessedData']
            yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

            if yellow_row_indices_in_final_df:
                for zero_based_row_idx in yellow_row_indices_in_final_df:
                    worksheet.set_row(zero_based_row_idx + 1, None, yellow_format)
            
            for i, col_name_str in enumerate(df_excel_output.columns.astype(str)):
                data_max_len_series = df_excel_output[col_name_str].astype(str).map(len)
                data_max_len = data_max_len_series.max() if not data_max_len_series.empty else 0
                header_len = len(col_name_str)
                if pd.isna(data_max_len): data_max_len = 0
                column_width = max(int(data_max_len), header_len) + 2
                worksheet.set_column(i, i, column_width)

        print(f"\n결과가 '{excel_filename}' 파일로 저장되었습니다.")
        print("삽입된 행들은 노란색으로 표시됩니다. 다른 특정 색상 서식은 적용되지 않았습니다.")

    except ImportError:
        print("Excel 파일 저장을 위해 'xlsxwriter' 라이브러리가 필요합니다. 'pip install xlsxwriter' 명령어로 설치해주세요.")
    except Exception as e_excel:
        print(f"Excel 파일 저장 중 오류 발생: {e_excel}")

elif not df_complete_table.empty and df_filtered.empty : # 초기 필터링 결과가 없었던 경우
    print("초기 필터링된 데이터(df_filtered)가 없어 Excel 파일을 생성하지 않았습니다.")
elif df_complete_table.empty : # 원본 데이터 자체가 없었던 경우
     print("원본 데이터(df_complete_table)가 없어 Excel 파일을 생성하지 않았습니다.")
else: # 그 외 output_rows_info가 비어있는 경우
    print("조건에 맞는 데이터가 없어 최종적으로 Excel 파일에 저장할 내용이 없습니다.")

