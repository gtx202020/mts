"""
BW Tools Excel Generator
SQLite 데이터베이스에서 인터페이스 정보를 추출하고 매칭하여 Excel/CSV로 출력합니다.
기존 iflist03a.py의 역할을 수행합니다.
"""

import sqlite3
import pandas as pd
import os
from typing import Optional, Dict, List, Tuple
from datetime import datetime
from bwtools_config import (
    DB_FILENAME, TABLE_NAME, COLUMN_NAMES, ADDITIONAL_COLUMNS,
    SYSTEM_MAPPING, FILE_PATH_TEMPLATES, EXCEL_COLORS, TEST_CONFIG
)

class ExcelGenerator:
    def __init__(self, db_path: Optional[str] = None):
        """
        ExcelGenerator 초기화
        
        Args:
            db_path: SQLite 데이터베이스 경로 (기본값: config의 DB_FILENAME)
        """
        self.db_path = db_path or DB_FILENAME
        self.table_name = TABLE_NAME
        self.df_complete_table = None
        
    def generate_excel(self, output_path: Optional[str] = None, 
                      output_format: str = 'xlsx') -> bool:
        """
        데이터베이스에서 데이터를 읽어 Excel/CSV 파일을 생성합니다.
        
        Args:
            output_path: 출력 파일 경로 (기본값: 자동 생성)
            output_format: 출력 형식 ('xlsx' 또는 'csv')
            
        Returns:
            성공 여부
        """
        try:
            # 데이터베이스 로드
            if not self._load_database():
                return False
            
            # 데이터 필터링 및 매칭
            result_df = self._process_data()
            
            # 출력 경로 설정
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"bwtools_excel_output_{timestamp}.{output_format}"
            
            # 파일 저장
            if output_format == 'xlsx':
                self._save_to_excel(result_df, output_path)
            elif output_format == 'csv':
                self._save_to_csv(result_df, output_path)
            else:
                raise ValueError(f"지원하지 않는 출력 형식: {output_format}")
            
            print(f"파일 생성 완료: {output_path}")
            print(f"총 {len(result_df)}개 행 출력")
            return True
            
        except Exception as e:
            print(f"Excel 생성 중 오류 발생: {str(e)}")
            return False
    
    def _load_database(self) -> bool:
        """데이터베이스에서 테이블을 로드합니다."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                query = f'SELECT * FROM "{self.table_name}"'
                self.df_complete_table = pd.read_sql_query(query, conn)
                print(f"데이터베이스 로드 완료: {len(self.df_complete_table)}개 행")
                return True
        except Exception as e:
            print(f"데이터베이스 로드 실패: {str(e)}")
            return False
    
    def _process_data(self) -> pd.DataFrame:
        """데이터를 처리하고 매칭을 수행합니다."""
        result_rows = []
        
        # LY/LZ 시스템 필터링
        filtered_df = self._filter_ly_lz_systems()
        print(f"LY/LZ 시스템 행 수: {len(filtered_df)}")
        
        # 각 행에 대해 매칭 수행
        for idx, base_row in filtered_df.iterrows():
            # 기본행 추가
            base_row_dict = base_row.to_dict()
            base_row_dict['color_flag'] = 'base'
            base_row_dict = self._add_file_paths(base_row_dict)
            base_row_dict = self._add_comparison_result(base_row_dict, None)
            result_rows.append(base_row_dict)
            
            # 매칭행 찾기
            matched_rows = self._find_matching_rows(base_row)
            
            if not matched_rows.empty:
                # 우선순위 적용
                selected_row = self._apply_priority(base_row, matched_rows)
                if selected_row is not None:
                    selected_row_dict = selected_row.to_dict()
                    selected_row_dict['color_flag'] = 'priority_filtered'
                    selected_row_dict = self._add_file_paths(selected_row_dict)
                    selected_row_dict = self._add_comparison_result(base_row_dict, selected_row_dict)
                    result_rows.append(selected_row_dict)
                else:
                    # 모든 매칭행 추가
                    for _, matched_row in matched_rows.iterrows():
                        matched_row_dict = matched_row.to_dict()
                        matched_row_dict['color_flag'] = 'match'
                        matched_row_dict = self._add_file_paths(matched_row_dict)
                        matched_row_dict = self._add_comparison_result(base_row_dict, matched_row_dict)
                        result_rows.append(matched_row_dict)
        
        return pd.DataFrame(result_rows)
    
    def _filter_ly_lz_systems(self) -> pd.DataFrame:
        """LY/LZ가 포함된 시스템을 필터링합니다."""
        send_col = COLUMN_NAMES['send_system']
        recv_col = COLUMN_NAMES['recv_system']
        
        mask = (
            self.df_complete_table[send_col].astype(str).str.contains('LY|LZ', na=False) |
            self.df_complete_table[recv_col].astype(str).str.contains('LY|LZ', na=False)
        )
        
        return self.df_complete_table[mask].copy()
    
    def _find_matching_rows(self, base_row: pd.Series) -> pd.DataFrame:
        """기본행에 대한 매칭행을 찾습니다."""
        if_name_col = COLUMN_NAMES['if_name']
        send_col = COLUMN_NAMES['send_system']
        recv_col = COLUMN_NAMES['recv_system']
        
        # 동일한 I/F명 찾기
        same_if = self.df_complete_table[
            self.df_complete_table[if_name_col] == base_row[if_name_col]
        ]
        
        # LY->LH, LZ->VO 변환 후 매칭
        base_send = str(base_row[send_col])
        base_recv = str(base_row[recv_col])
        
        # 변환
        for old, new in SYSTEM_MAPPING.items():
            base_send = base_send.replace(old, new)
            base_recv = base_recv.replace(old, new)
        
        # 매칭 조건
        mask = (
            (same_if[send_col] == base_send) |
            (same_if[recv_col] == base_recv)
        )
        
        matched = same_if[mask]
        
        # 자기 자신 제외
        return matched[matched.index != base_row.name]
    
    def _apply_priority(self, base_row: pd.Series, matched_rows: pd.DataFrame) -> Optional[pd.Series]:
        """우선순위를 적용하여 매칭행을 선택합니다."""
        if len(matched_rows) <= 1:
            return None
        
        send_col = COLUMN_NAMES['send_system']
        recv_col = COLUMN_NAMES['recv_system']
        
        # 변환된 값
        base_send = str(base_row[send_col])
        base_recv = str(base_row[recv_col])
        for old, new in SYSTEM_MAPPING.items():
            base_send = base_send.replace(old, new)
            base_recv = base_recv.replace(old, new)
        
        # 케이스 1: 송신시스템과 수신시스템 모두 매칭
        case1 = matched_rows[
            (matched_rows[send_col] == base_send) &
            (matched_rows[recv_col] == base_recv)
        ]
        if not case1.empty:
            return case1.iloc[0]
        
        # 케이스 2: 송신시스템만 매칭
        case2 = matched_rows[matched_rows[send_col] == base_send]
        if not case2.empty:
            return case2.iloc[0]
        
        # 케이스 2-1: 수신시스템만 매칭
        case2_1 = matched_rows[matched_rows[recv_col] == base_recv]
        if not case2_1.empty:
            return case2_1.iloc[0]
        
        return None
    
    def _add_file_paths(self, row_dict: dict) -> dict:
        """파일 경로 정보를 추가합니다."""
        # 송신 파일 경로
        send_path = self._create_file_path(row_dict, 'send')
        row_dict[ADDITIONAL_COLUMNS['send_file_path']] = send_path
        row_dict[ADDITIONAL_COLUMNS['send_file_exists']] = 'Y' if os.path.exists(send_path) else 'N'
        row_dict[ADDITIONAL_COLUMNS['send_file_created']] = 'Y' if os.path.exists(send_path) else 'N'
        
        # 수신 파일 경로
        recv_path = self._create_file_path(row_dict, 'recv')
        row_dict[ADDITIONAL_COLUMNS['recv_file_path']] = recv_path
        row_dict[ADDITIONAL_COLUMNS['recv_file_exists']] = 'Y' if os.path.exists(recv_path) else 'N'
        row_dict[ADDITIONAL_COLUMNS['recv_file_created']] = 'Y' if os.path.exists(recv_path) else 'N'
        
        # 스키마 파일 경로
        send_schema_path = self._create_schema_file_path(row_dict, 'send')
        row_dict[ADDITIONAL_COLUMNS['send_schema_file']] = send_schema_path
        row_dict[ADDITIONAL_COLUMNS['send_schema_exists']] = 'Y' if os.path.exists(send_schema_path) else 'N'
        row_dict[ADDITIONAL_COLUMNS['send_schema_created']] = 'Y' if os.path.exists(send_schema_path) else 'N'
        
        recv_schema_path = self._create_schema_file_path(row_dict, 'recv')
        row_dict[ADDITIONAL_COLUMNS['recv_schema_file']] = recv_schema_path
        row_dict[ADDITIONAL_COLUMNS['recv_schema_exists']] = 'Y' if os.path.exists(recv_schema_path) else 'N'
        row_dict[ADDITIONAL_COLUMNS['recv_schema_created']] = 'Y' if os.path.exists(recv_schema_path) else 'N'
        
        # 디렉토리 파일 수
        row_dict[ADDITIONAL_COLUMNS['send_df']] = self._count_files_in_directory(send_path)
        row_dict[ADDITIONAL_COLUMNS['recv_df']] = self._count_files_in_directory(recv_path)
        
        return row_dict
    
    def _create_file_path(self, row_dict: dict, file_type: str) -> str:
        """파일 경로를 생성합니다."""
        template = FILE_PATH_TEMPLATES.get(file_type, '')
        
        corp_key = f'{file_type}_corp'
        corp = str(row_dict.get(COLUMN_NAMES[corp_key], '')).lower()
        
        return template.format(
            corp=corp,
            pkg=str(row_dict.get(COLUMN_NAMES[f'{file_type}_pkg'], '')),
            task=str(row_dict.get(COLUMN_NAMES[f'{file_type}_task'], '')),
            ems=str(row_dict.get(COLUMN_NAMES['ems_name'], '')),
            group_id=str(row_dict.get(COLUMN_NAMES['group_id'], '')),
            event_id=str(row_dict.get(COLUMN_NAMES['event_id'], ''))
        )
    
    def _create_schema_file_path(self, row_dict: dict, file_type: str) -> str:
        """스키마 파일 경로를 생성합니다."""
        template = FILE_PATH_TEMPLATES.get(f'{file_type}_schema', '')
        
        corp_key = f'{file_type}_corp'
        corp = str(row_dict.get(COLUMN_NAMES[corp_key], '')).lower()
        
        if file_type == 'send':
            return template.format(
                corp=corp,
                pkg=str(row_dict.get(COLUMN_NAMES['send_pkg'], '')),
                task=str(row_dict.get(COLUMN_NAMES['send_task'], '')),
                db_name=str(row_dict.get(COLUMN_NAMES['send_db_name'], '')),
                schema=str(row_dict.get(COLUMN_NAMES['send_schema'], '')),
                table=str(row_dict.get(COLUMN_NAMES['source_table'], ''))
            )
        else:
            return template.format(
                corp=corp,
                pkg=str(row_dict.get(COLUMN_NAMES['recv_pkg'], '')),
                task=str(row_dict.get(COLUMN_NAMES['recv_task'], '')),
                table=str(row_dict.get(COLUMN_NAMES['dest_table'], ''))
            )
    
    def _count_files_in_directory(self, file_path: str) -> int:
        """디렉토리의 파일 수를 계산합니다."""
        dir_path = os.path.dirname(file_path)
        if os.path.exists(dir_path) and os.path.isdir(dir_path):
            return len([f for f in os.listdir(dir_path) if os.path.isfile(os.path.join(dir_path, f))])
        return 0
    
    def _add_comparison_result(self, base_row: dict, matched_row: Optional[dict]) -> dict:
        """비교 결과를 추가합니다."""
        if matched_row is None:
            base_row[ADDITIONAL_COLUMNS['compare_log']] = "매칭행 없음"
            return base_row
        
        errors = []
        
        # 15가지 비교 검증 규칙 (간소화된 버전)
        comparisons = [
            ('송신시스템', COLUMN_NAMES['send_system']),
            ('수신시스템', COLUMN_NAMES['recv_system']),
            ('송신법인', COLUMN_NAMES['send_corp']),
            ('수신법인', COLUMN_NAMES['recv_corp']),
            ('송신패키지', COLUMN_NAMES['send_pkg']),
            ('수신패키지', COLUMN_NAMES['recv_pkg']),
            ('송신업무명', COLUMN_NAMES['send_task']),
            ('수신업무명', COLUMN_NAMES['recv_task']),
            ('EMS명', COLUMN_NAMES['ems_name']),
            ('Group ID', COLUMN_NAMES['group_id']),
            ('Event_ID', COLUMN_NAMES['event_id']),
            ('Source Table', COLUMN_NAMES['source_table']),
            ('Destination Table', COLUMN_NAMES['dest_table']),
            ('주기구분', COLUMN_NAMES['cycle_type']),
            ('주기', COLUMN_NAMES['cycle'])
        ]
        
        for name, col in comparisons:
            base_val = str(base_row.get(col, ''))
            matched_val = str(matched_row.get(col, ''))
            
            # LY/LZ -> LH/VO 변환 고려
            base_val_converted = base_val
            for old, new in SYSTEM_MAPPING.items():
                base_val_converted = base_val_converted.replace(old, new)
            
            if base_val_converted != matched_val and base_val != matched_val:
                errors.append(f"{name} 불일치")
        
        matched_row[ADDITIONAL_COLUMNS['compare_log']] = ', '.join(errors) if errors else "정상"
        return matched_row
    
    def _save_to_excel(self, df: pd.DataFrame, output_path: str):
        """DataFrame을 Excel 파일로 저장합니다."""
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # color_flag 컬럼 제거
            df_output = df.drop('color_flag', axis=1, errors='ignore')
            df_output.to_excel(writer, sheet_name='Interface_Matching', index=False)
            
            # 서식 적용 (색상)
            workbook = writer.book
            worksheet = writer.sheets['Interface_Matching']
            
            # 색상 포맷 정의
            yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
            green_format = workbook.add_format({'bg_color': '#90EE90'})
            
            # 행별로 색상 적용
            for idx, color_flag in enumerate(df['color_flag'].values):
                if color_flag == 'match':
                    worksheet.set_row(idx + 1, None, yellow_format)
                elif color_flag == 'priority_filtered':
                    worksheet.set_row(idx + 1, None, green_format)
            
            # 열 너비 자동 조정
            for i, col in enumerate(df_output.columns):
                max_len = max(df_output[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, min(max_len, 50))
    
    def _save_to_csv(self, df: pd.DataFrame, output_path: str):
        """DataFrame을 CSV 파일로 저장합니다."""
        # color_flag 컬럼 제거
        df_output = df.drop('color_flag', axis=1, errors='ignore')
        df_output.to_csv(output_path, index=False, encoding='utf-8-sig')


def main():
    """메인 실행 함수"""
    generator = ExcelGenerator()
    
    # Excel 생성
    print("Excel 파일 생성 중...")
    if TEST_CONFIG['use_csv']:
        # 테스트 모드에서는 CSV 출력
        generator.generate_excel(output_format='csv')
    else:
        generator.generate_excel(output_format='xlsx')


if __name__ == "__main__":
    main()