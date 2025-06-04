"""
인터페이스 데이터 처리 모듈

SQLite 데이터베이스에서 인터페이스 목록을 추출하고 가공하여 Excel 파일로 출력합니다.
주로 'LY'와 'LZ'로 표시된 시스템을 'LH'와 'VO'로 변환된 대응 인터페이스를 찾아 매핑합니다.
"""

import sqlite3
import pandas as pd
import os
from typing import Optional, Dict, Any


class InterfaceProcessor:
    """인터페이스 데이터를 처리하는 클래스"""
    
    def __init__(self, db_filename: str = 'iflist.sqlite', table_name: str = 'iflist'):
        """
        InterfaceProcessor 초기화
        
        Args:
            db_filename: SQLite 데이터베이스 파일명
            table_name: 테이블명
        """
        self.db_filename = db_filename
        self.table_name = table_name
        
        # 컬럼명 설정
        self.column_b_name = '송신시스템'
        self.column_c_name = '수신시스템'
        self.column_d_name = 'I/F명'
        
        # 추가 컬럼 이름
        self.column_send_corp_name = '송신\n법인'
        self.column_recv_corp_name = '수신\n법인'
        self.column_send_pkg_name = '송신패키지'
        self.column_recv_pkg_name = '수신패키지'
        self.column_send_task_name = '송신\n업무명'
        self.column_recv_task_name = '수신\n업무명'
        self.column_ems_name = 'EMS명'
        self.column_group_id = 'Group ID'
        self.column_event_id = 'Event_ID'
        
        # 변환 값
        self.val_ly = 'LY'
        self.val_lz = 'LZ'
        self.replace_ly_with = 'LH'
        self.replace_lz_with = 'VO'
        
        # 디버깅 모드
        self.debug_mode = 2
    
    def replace_ly_lz(self, text: str) -> str:
        """문자열에서 'LY'를 'LH'로, 'LZ'를 'VO'로 교체"""
        if not isinstance(text, str):
            return text
        return text.replace('LY', 'LH').replace('LZ', 'VO')
    
    def create_file_path(self, row: pd.Series, is_send: bool = True, color_flag: Optional[str] = None) -> str:
        """
        주어진 행 데이터로부터 파일 경로를 생성
        
        Args:
            row: 데이터프레임의 행
            is_send: 송신 파일 경로인지 여부 (False면 수신 파일 경로)
            color_flag: 행의 색상 정보 (None: 기본행, 'yellow'/'green': 매칭행)
            
        Returns:
            생성된 파일 경로 문자열
        """
        try:
            # 기본 경로 시작
            base_path = "C:\\BwProject"
            
            # 사용할 컬럼 선택 (송신/수신에 따라)
            corp_col = self.column_send_corp_name if is_send else self.column_recv_corp_name
            pkg_col = self.column_send_pkg_name if is_send else self.column_recv_pkg_name
            task_col = self.column_send_task_name if is_send else self.column_recv_task_name
            
            # 안전하게 컬럼값 가져오기
            def safe_get_value(df_row, column_name):
                try:
                    val = df_row[column_name] if column_name in df_row.index else ""
                    return str(val).strip() if pd.notna(val) else ""
                except:
                    return ""
            
            # 필요한 값들 가져오기
            corp_val = safe_get_value(row, corp_col)
            pkg_val = safe_get_value(row, pkg_col)
            task_val = safe_get_value(row, task_col)
            ems_val = safe_get_value(row, self.column_ems_name)
            group_id = safe_get_value(row, self.column_group_id)
            event_id = safe_get_value(row, self.column_event_id)
            recv_task = "" if is_send else safe_get_value(row, self.column_recv_task_name)
            
            # 1번 디렉토리 (법인 정보에 따라)
            dir1 = ""
            if corp_val == "KR":
                dir1 = "KR"
            elif corp_val == "NJ":
                dir1 = "CN"
            elif corp_val == "VH":
                dir1 = "VN"
            else:
                dir1 = "UNK"
            
            # 접미사 결정 (기본행은 _TEST_SOURCE, 매칭행은 _PROD_SOURCE)
            if color_flag is None:
                dir1 += "_TEST_SOURCE"
            else:
                dir1 += "_PROD_SOURCE"
            
            # 2번 디렉토리 (패키지의 첫 '_' 이전 부분)
            dir2 = pkg_val.split('_')[0] if '_' in pkg_val and pkg_val else pkg_val
            
            # 3번 디렉토리 (조건부)
            dir3 = ""
            if task_val and any(keyword in task_val for keyword in ["PNL", "EAS", "MOD", "MES"]):
                parts = task_val.split('_')
                if len(parts) > 1:
                    dir3 = parts[-1]
            
            # 4번 디렉토리 (EMS명에 따라)
            dir4 = "EMS_64000" if ems_val == "MES01" else "EMS_63000"
            
            # 5번 디렉토리 (패키지 전체 이름)
            dir5 = pkg_val
            
            # 파일명
            if is_send:
                filename = f"{group_id}.{event_id}.process" if group_id and event_id else "unknown.process"
            else:
                filename = f"{group_id}.{event_id}.{recv_task}.process" if group_id and event_id else "unknown.process"
            
            # 경로 구성
            path_parts = [base_path, dir1, dir2, "Processes"]
            if dir3:
                path_parts.append(dir3)
            path_parts.extend([dir4, dir5, filename])
            
            return "\\".join(path_parts)
        except Exception as e:
            print(f"파일 경로 생성 중 오류 발생: {e}")
            return ""
    
    def create_schema_file_path(self, row: pd.Series, is_send: bool = True, color_flag: Optional[str] = None) -> str:
        """
        주어진 행 데이터로부터 스키마 파일 경로를 생성
        
        Args:
            row: 데이터프레임의 행
            is_send: 송신 스키마 파일 경로인지 여부
            color_flag: 행의 색상 정보
            
        Returns:
            생성된 스키마 파일 경로 문자열
        """
        try:
            base_path = "C:\\BwProject"
            
            corp_col = self.column_send_corp_name if is_send else self.column_recv_corp_name
            pkg_col = self.column_send_pkg_name if is_send else self.column_recv_pkg_name
            db_name_col = '송신\nDB Name'
            schema_col = '송신 \nSchema'
            
            def safe_get_value(df_row, column_name):
                try:
                    val = df_row[column_name] if column_name in df_row.index else ""
                    return str(val).strip() if pd.notna(val) else ""
                except:
                    return ""
            
            corp_val = safe_get_value(row, corp_col)
            pkg_val = safe_get_value(row, pkg_col)
            db_name = safe_get_value(row, db_name_col)
            schema = safe_get_value(row, schema_col)
            source_table = safe_get_value(row, 'Source Table')
            
            # 법인별 디렉토리 설정
            dir1 = ""
            if corp_val == "KR":
                dir1 = "KR"
            elif corp_val == "NJ":
                dir1 = "CN"
            elif corp_val == "VH":
                dir1 = "VN"
            else:
                dir1 = "UNK"
            
            if color_flag is None:
                dir1 += "_TEST_SOURCE"
            else:
                dir1 += "_PROD_SOURCE"
            
            dir2 = pkg_val.split('_')[0] if '_' in pkg_val and pkg_val else pkg_val
            dir3 = "SharedResources\\Schema\\source"
            dir4 = db_name
            dir5 = schema
            
            # 파일명 생성
            file_name = ""
            if '.' in source_table:
                file_name = source_table.split('.')[1] + '.xsd'
            else:
                file_name = source_table + '.xsd'
            
            path_parts = [base_path, dir1, dir2, dir3, dir4, dir5, file_name]
            path_parts = [part for part in path_parts if part]
            
            return "\\".join(path_parts)
        
        except Exception as e:
            print(f"스키마 파일 경로 생성 오류: {e}")
            return "경로 생성 오류"
    
    def check_file_exists(self, file_path: str) -> int:
        """파일 존재 여부 확인"""
        try:
            return 1 if os.path.isfile(file_path) else 0
        except:
            return 0
    
    def count_files_in_directory(self, file_path: str) -> int:
        """디렉토리 내 파일 개수 확인"""
        try:
            directory = os.path.dirname(file_path)
            if not directory or not os.path.isdir(directory):
                return 0
            return len([f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))])
        except:
            return 0
    
    def process_interface_data(self, output_filename: Optional[str] = None) -> bool:
        """
        인터페이스 데이터를 처리하여 Excel 파일로 출력
        
        Args:
            output_filename: 출력 파일명 (None이면 자동 생성)
            
        Returns:
            처리 성공 여부
        """
        try:
            # 데이터베이스 연결
            if not os.path.exists(self.db_filename):
                print(f"오류: 데이터베이스 파일을 찾을 수 없습니다 - {self.db_filename}")
                return False
            
            conn = sqlite3.connect(self.db_filename)
            cursor = conn.cursor()
            
            # 전체 데이터 로드
            cursor.execute(f'SELECT * FROM "{self.table_name}"')
            all_rows_from_db = cursor.fetchall()
            
            if not all_rows_from_db:
                print(f"테이블 '{self.table_name}'에 데이터가 없습니다.")
                conn.close()
                return False
            
            column_names_from_db = [description[0] for description in cursor.description]
            df_complete_table = pd.DataFrame(all_rows_from_db, columns=column_names_from_db)
            print(f"전체 테이블에 총 {len(df_complete_table)}개의 행이 로드되었습니다.")
            
            # 초기 필터링 (LY 또는 LZ 포함)
            cond_b_contains = (
                df_complete_table[self.column_b_name].astype(str).str.contains(self.val_ly, na=False) |
                df_complete_table[self.column_b_name].astype(str).str.contains(self.val_lz, na=False)
            )
            cond_c_contains = (
                df_complete_table[self.column_c_name].astype(str).str.contains(self.val_ly, na=False) |
                df_complete_table[self.column_c_name].astype(str).str.contains(self.val_lz, na=False)
            )
            df_filtered = df_complete_table[cond_b_contains | cond_c_contains].copy()
            
            if df_filtered.empty:
                print("초기 필터링 조건에 맞는 데이터가 없습니다.")
                conn.close()
                return False
            
            print(f"초기 필터링 후 {len(df_filtered)}개의 행이 남았습니다.")
            
            # 매칭 작업 수행
            output_rows_info = []
            
            for idx_filtered, current_row in df_filtered.iterrows():
                output_rows_info.append({'data_row': current_row.copy(), 'color_flag': None})
                
                current_d_val = str(current_row[self.column_d_name]) if pd.notna(current_row[self.column_d_name]) else ""
                current_d_val_stripped = current_d_val.strip()
                
                current_b_val = str(current_row[self.column_b_name]) if pd.notna(current_row[self.column_b_name]) else ""
                current_c_val = str(current_row[self.column_c_name]) if pd.notna(current_row[self.column_c_name]) else ""
                
                matching_rows = []
                
                # 매칭 행 찾기
                for idx_complete, target_row in df_complete_table.iterrows():
                    if current_row.name == target_row.name:
                        continue
                    
                    target_d_val = str(target_row[self.column_d_name]) if pd.notna(target_row[self.column_d_name]) else ""
                    target_d_val_stripped = target_d_val.strip()
                    
                    if current_d_val_stripped != target_d_val_stripped:
                        continue
                    
                    target_b_val = str(target_row[self.column_b_name]) if pd.notna(target_row[self.column_b_name]) else ""
                    target_c_val = str(target_row[self.column_c_name]) if pd.notna(target_row[self.column_c_name]) else ""
                    
                    # B 컬럼 매칭 확인
                    cond2_b_match = False
                    if self.val_ly in current_b_val:
                        transformed_b = current_b_val.replace(self.val_ly, self.replace_ly_with)
                        if target_b_val == transformed_b:
                            cond2_b_match = True
                    if not cond2_b_match and (self.val_lz in current_b_val):
                        transformed_b = current_b_val.replace(self.val_lz, self.replace_lz_with)
                        if target_b_val == transformed_b:
                            cond2_b_match = True
                    
                    # C 컬럼 매칭 확인
                    cond3_c_match = False
                    if self.val_ly in current_c_val:
                        transformed_c = current_c_val.replace(self.val_ly, self.replace_ly_with)
                        if target_c_val == transformed_c:
                            cond3_c_match = True
                    if not cond3_c_match and (self.val_lz in current_c_val):
                        transformed_c = current_c_val.replace(self.val_lz, self.replace_lz_with)
                        if target_c_val == transformed_c:
                            cond3_c_match = True
                    
                    if cond2_b_match or cond3_c_match:
                        matching_rows.append({
                            'row': target_row.copy(),
                            'b_match': cond2_b_match,
                            'c_match': cond3_c_match,
                            'same_b_val': target_b_val == current_b_val,
                            'same_c_val': target_c_val == current_c_val
                        })
                
                # 매칭된 행 처리
                if matching_rows:
                    if len(matching_rows) == 1:
                        output_rows_info.append({'data_row': matching_rows[0]['row'], 'color_flag': 'green'})
                    else:
                        # 우선순위 필터링
                        if self.debug_mode == 1:
                            for row in matching_rows:
                                output_rows_info.append({'data_row': row['row'], 'color_flag': 'yellow'})
                        
                        filtered_row = None
                        
                        # 케이스 1: B와 C 모두 매칭
                        case1_rows = [row for row in matching_rows if row['b_match'] and row['c_match']]
                        if case1_rows:
                            filtered_row = case1_rows[0]['row']
                        else:
                            # 케이스 2: B가 같은 행
                            case2_rows = [row for row in matching_rows if row['same_b_val']]
                            if case2_rows:
                                filtered_row = case2_rows[0]['row']
                            else:
                                # 케이스 2-1: C가 같은 행
                                case2_1_rows = [row for row in matching_rows if row['same_c_val']]
                                if case2_1_rows:
                                    filtered_row = case2_1_rows[0]['row']
                        
                        if filtered_row is not None:
                            output_rows_info.append({'data_row': filtered_row.copy(), 'color_flag': 'green'})
            
            # 최종 DataFrame 생성
            if output_rows_info:
                for item in output_rows_info:
                    item['data_row']['color_flag'] = item['color_flag']
                
                final_df_data = [item['data_row'] for item in output_rows_info]
                cols_for_final_df = list(column_names_from_db) + ['color_flag']
                df_excel_output = pd.DataFrame(final_df_data, columns=cols_for_final_df).reset_index(drop=True)
                
                # 파일 경로 추가
                df_excel_output['송신파일경로'] = df_excel_output.apply(
                    lambda row: self.create_file_path(row, is_send=True, color_flag=row.get('color_flag')), axis=1)
                df_excel_output['수신파일경로'] = df_excel_output.apply(
                    lambda row: self.create_file_path(row, is_send=False, color_flag=row.get('color_flag')), axis=1)
                
                # 파일 존재 여부 확인
                df_excel_output['송신파일존재'] = df_excel_output['송신파일경로'].apply(self.check_file_exists)
                df_excel_output['수신파일존재'] = df_excel_output['수신파일경로'].apply(self.check_file_exists)
                
                # 파일 생성 여부 컬럼 추가
                df_excel_output['송신파일생성여부'] = df_excel_output.apply(
                    lambda row: '' if row.get('color_flag') is not None else ('1' if row.get('개발구분') == '신규' else ''), axis=1)
                df_excel_output['수신파일생성여부'] = df_excel_output.apply(
                    lambda row: '' if row.get('color_flag') is not None else '1', axis=1)
                
                # 디렉토리 파일 개수
                df_excel_output['송신DF'] = df_excel_output.apply(
                    lambda row: self.count_files_in_directory(row['송신파일경로']) if row['송신파일존재'] == 1 else 0, axis=1)
                df_excel_output['수신DF'] = df_excel_output.apply(
                    lambda row: self.count_files_in_directory(row['수신파일경로']) if row['수신파일존재'] == 1 else 0, axis=1)
                
                # 스키마 파일 경로 추가
                df_excel_output['송신스키마파일명'] = df_excel_output.apply(
                    lambda row: self.create_schema_file_path(row, is_send=True, color_flag=row.get('color_flag')), axis=1)
                df_excel_output['수신스키마파일명'] = df_excel_output.apply(
                    lambda row: self.create_schema_file_path(row, is_send=False, color_flag=row.get('color_flag')), axis=1)
                
                # 스키마 파일 존재 여부
                df_excel_output['송신스키마파일존재'] = df_excel_output['송신스키마파일명'].apply(self.check_file_exists)
                df_excel_output['수신스키마파일존재'] = df_excel_output['수신스키마파일명'].apply(self.check_file_exists)
                
                # 스키마 파일 생성 여부
                df_excel_output['송신스키마파일생성여부'] = df_excel_output.apply(
                    lambda row: '' if row.get('color_flag') is not None else ('1' if row.get('개발구분') == '신규' else ''), axis=1)
                df_excel_output['수신스키마파일생성여부'] = df_excel_output.apply(
                    lambda row: '' if row.get('color_flag') is not None else '1', axis=1)
                
                # CSV 파일로 저장 (테스트 편의성을 위해)
                if output_filename is None:
                    output_filename = "rft_interface_processed.csv"
                
                df_excel_output.to_csv(output_filename, index=False, encoding='utf-8-sig')
                print(f"결과가 '{output_filename}' 파일로 저장되었습니다.")
                print(f"총 {len(df_excel_output)}개의 행이 처리되었습니다.")
                
                conn.close()
                return True
            
            else:
                print("처리할 데이터가 없습니다.")
                conn.close()
                return False
                
        except Exception as e:
            print(f"인터페이스 데이터 처리 중 오류 발생: {str(e)}")
            return False


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("인터페이스 데이터 처리 도구")
    print("=" * 60)
    
    processor = InterfaceProcessor()
    
    while True:
        print("\n메뉴:")
        print("1. 인터페이스 데이터 처리 (CSV 출력)")
        print("0. 종료")
        
        choice = input("\n선택하세요: ").strip()
        
        if choice == "1":
            output_file = input("출력 파일명을 입력하세요 (Enter: 기본값): ").strip()
            if not output_file:
                output_file = None
            processor.process_interface_data(output_file)
            
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
            
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")


if __name__ == "__main__":
    main()