"""
BW Tools Excel Generator 단위 테스트
"""

import unittest
import os
import pandas as pd
from bwtools_db_creator import DBCreator
from bwtools_excel_generator import ExcelGenerator
from bwtools_config import COLUMN_NAMES, ADDITIONAL_COLUMNS, SYSTEM_MAPPING

class TestExcelGenerator(unittest.TestCase):
    def setUp(self):
        """테스트 설정"""
        self.test_db_path = 'test_iflist.sqlite'
        self.test_output_csv = 'test_output.csv'
        self.test_output_xlsx = 'test_output.xlsx'
        
        # 테스트 DB 생성
        self.db_creator = DBCreator(self.test_db_path)
        self.db_creator.create_test_database()
        
        # Excel Generator 생성
        self.generator = ExcelGenerator(self.test_db_path)
        
    def tearDown(self):
        """테스트 정리"""
        # 테스트 파일 삭제
        for file in [self.test_db_path, self.test_output_csv, self.test_output_xlsx]:
            if os.path.exists(file):
                os.remove(file)
    
    def test_load_database(self):
        """데이터베이스 로드 테스트"""
        result = self.generator._load_database()
        self.assertTrue(result)
        self.assertIsNotNone(self.generator.df_complete_table)
        self.assertGreater(len(self.generator.df_complete_table), 0)
    
    def test_filter_ly_lz_systems(self):
        """LY/LZ 시스템 필터링 테스트"""
        self.generator._load_database()
        filtered_df = self.generator._filter_ly_lz_systems()
        
        # 필터링된 데이터가 있는지 확인
        self.assertGreater(len(filtered_df), 0)
        
        # 모든 행이 LY 또는 LZ를 포함하는지 확인
        for _, row in filtered_df.iterrows():
            send_system = str(row[COLUMN_NAMES['send_system']])
            recv_system = str(row[COLUMN_NAMES['recv_system']])
            has_ly_lz = 'LY' in send_system or 'LZ' in send_system or \
                       'LY' in recv_system or 'LZ' in recv_system
            self.assertTrue(has_ly_lz)
    
    def test_find_matching_rows(self):
        """매칭행 찾기 테스트"""
        self.generator._load_database()
        
        # LY 시스템 행 찾기
        ly_rows = self.generator.df_complete_table[
            self.generator.df_complete_table[COLUMN_NAMES['send_system']].str.contains('LY')
        ]
        
        if not ly_rows.empty:
            base_row = ly_rows.iloc[0]
            matched = self.generator._find_matching_rows(base_row)
            
            # 자기 자신은 제외되었는지 확인
            self.assertNotIn(base_row.name, matched.index)
    
    def test_apply_priority(self):
        """우선순위 적용 테스트"""
        self.generator._load_database()
        
        # 테스트용 DataFrame 생성
        test_data = {
            COLUMN_NAMES['send_system']: ['LHMES', 'LHMES', 'VOMES'],
            COLUMN_NAMES['recv_system']: ['VOWMS', 'LZWMS', 'VOWMS'],
            COLUMN_NAMES['if_name']: ['IF_001', 'IF_001', 'IF_001']
        }
        matched_rows = pd.DataFrame(test_data)
        
        # 기준행 (LY -> LH로 변환되어야 함)
        base_data = {
            COLUMN_NAMES['send_system']: 'LYMES',
            COLUMN_NAMES['recv_system']: 'LZWMS',
            COLUMN_NAMES['if_name']: 'IF_001'
        }
        base_row = pd.Series(base_data)
        
        # 우선순위 적용
        result = self.generator._apply_priority(base_row, matched_rows)
        
        # 케이스1이 선택되었는지 확인 (송신/수신 모두 매칭)
        if result is not None:
            self.assertEqual(result[COLUMN_NAMES['send_system']], 'LHMES')
            self.assertEqual(result[COLUMN_NAMES['recv_system']], 'VOWMS')
    
    def test_create_file_path(self):
        """파일 경로 생성 테스트"""
        row_dict = {
            COLUMN_NAMES['send_corp']: 'LYCORP',
            COLUMN_NAMES['send_pkg']: 'PKG_LY',
            COLUMN_NAMES['send_task']: 'TASK_01',
            COLUMN_NAMES['ems_name']: 'EMS_TEST',
            COLUMN_NAMES['group_id']: '001',
            COLUMN_NAMES['event_id']: 'EVT_0001'
        }
        
        path = self.generator._create_file_path(row_dict, 'send')
        
        # 경로에 필요한 요소들이 포함되었는지 확인
        self.assertIn('lycorp', path)  # 소문자 변환 확인
        self.assertIn('PKG_LY', path)
        self.assertIn('TASK_01', path)
        self.assertIn('EMS_TEST_GRP_001_EVT_0001_SND.process', path)
    
    def test_add_file_paths(self):
        """파일 경로 정보 추가 테스트"""
        row_dict = {
            COLUMN_NAMES['send_corp']: 'LYCORP',
            COLUMN_NAMES['recv_corp']: 'LZCORP',
            COLUMN_NAMES['send_pkg']: 'PKG_LY',
            COLUMN_NAMES['recv_pkg']: 'PKG_LZ',
            COLUMN_NAMES['send_task']: 'TASK_01',
            COLUMN_NAMES['recv_task']: 'TASK_02',
            COLUMN_NAMES['ems_name']: 'EMS_TEST',
            COLUMN_NAMES['group_id']: '001',
            COLUMN_NAMES['event_id']: 'EVT_0001',
            COLUMN_NAMES['send_db_name']: 'LYDB',
            COLUMN_NAMES['send_schema']: 'LYSCHEMA',
            COLUMN_NAMES['source_table']: 'TB_SOURCE',
            COLUMN_NAMES['dest_table']: 'TB_DEST'
        }
        
        result = self.generator._add_file_paths(row_dict)
        
        # 모든 추가 컬럼이 생성되었는지 확인
        self.assertIn(ADDITIONAL_COLUMNS['send_file_path'], result)
        self.assertIn(ADDITIONAL_COLUMNS['recv_file_path'], result)
        self.assertIn(ADDITIONAL_COLUMNS['send_schema_file'], result)
        self.assertIn(ADDITIONAL_COLUMNS['recv_schema_file'], result)
        self.assertIn(ADDITIONAL_COLUMNS['send_file_exists'], result)
        self.assertIn(ADDITIONAL_COLUMNS['recv_file_exists'], result)
    
    def test_add_comparison_result(self):
        """비교 결과 추가 테스트"""
        base_row = {
            COLUMN_NAMES['send_system']: 'LYMES',
            COLUMN_NAMES['recv_system']: 'LZWMS',
            COLUMN_NAMES['send_corp']: 'LYCORP',
            COLUMN_NAMES['recv_corp']: 'LZCORP'
        }
        
        # 매칭행 없는 경우
        result1 = self.generator._add_comparison_result(base_row.copy(), None)
        self.assertEqual(result1[ADDITIONAL_COLUMNS['compare_log']], "매칭행 없음")
        
        # 매칭행 있는 경우 (정상)
        matched_row = {
            COLUMN_NAMES['send_system']: 'LHMES',  # LY -> LH
            COLUMN_NAMES['recv_system']: 'VOWMS',  # LZ -> VO
            COLUMN_NAMES['send_corp']: 'LHCORP',   # LY -> LH
            COLUMN_NAMES['recv_corp']: 'VOCORP'    # LZ -> VO
        }
        
        result2 = self.generator._add_comparison_result(base_row.copy(), matched_row)
        self.assertEqual(result2[ADDITIONAL_COLUMNS['compare_log']], "정상")
    
    def test_generate_excel_csv(self):
        """CSV 파일 생성 테스트"""
        result = self.generator.generate_excel(self.test_output_csv, 'csv')
        self.assertTrue(result)
        self.assertTrue(os.path.exists(self.test_output_csv))
        
        # CSV 파일 읽기
        df = pd.read_csv(self.test_output_csv)
        self.assertGreater(len(df), 0)
        
        # 필수 컬럼 확인
        self.assertIn(ADDITIONAL_COLUMNS['send_file_path'], df.columns)
        self.assertIn(ADDITIONAL_COLUMNS['compare_log'], df.columns)
    
    def test_generate_excel_xlsx(self):
        """Excel 파일 생성 테스트"""
        # xlsxwriter가 설치되어 있는 경우에만 테스트
        try:
            import xlsxwriter
            result = self.generator.generate_excel(self.test_output_xlsx, 'xlsx')
            self.assertTrue(result)
            self.assertTrue(os.path.exists(self.test_output_xlsx))
        except ImportError:
            self.skipTest("xlsxwriter가 설치되지 않음")


if __name__ == '__main__':
    unittest.main()