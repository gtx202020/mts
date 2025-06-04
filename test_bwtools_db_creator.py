"""
BW Tools DB Creator 단위 테스트
"""

import unittest
import os
import sqlite3
import pandas as pd
from bwtools_db_creator import DBCreator
from bwtools_config import COLUMN_NAMES, TEST_CONFIG

class TestDBCreator(unittest.TestCase):
    def setUp(self):
        """테스트 설정"""
        self.test_db_path = 'test_iflist.sqlite'
        self.test_csv_path = 'test_input.csv'
        self.creator = DBCreator(self.test_db_path)
        
    def tearDown(self):
        """테스트 정리"""
        # 테스트 파일 삭제
        for file in [self.test_db_path, self.test_csv_path]:
            if os.path.exists(file):
                os.remove(file)
    
    def test_create_test_database(self):
        """테스트 데이터베이스 생성 테스트"""
        # 테스트 DB 생성
        result = self.creator.create_test_database()
        self.assertTrue(result)
        
        # DB 파일 존재 확인
        self.assertTrue(os.path.exists(self.test_db_path))
        
        # 데이터 검증
        verify_result = self.creator.verify_database()
        self.assertTrue(verify_result['success'])
        
        # 예상 행 수 확인
        expected_rows = sum(TEST_CONFIG['sample_rows'].values())
        self.assertEqual(verify_result['row_count'], expected_rows)
        
        # 필수 컬럼 존재 확인
        columns = verify_result['columns']
        for key, col_name in COLUMN_NAMES.items():
            self.assertIn(col_name, columns)
    
    def test_create_database_from_csv(self):
        """CSV 파일로부터 데이터베이스 생성 테스트"""
        # 테스트 CSV 생성
        test_data = {
            COLUMN_NAMES['send_system']: ['LYMES', 'LZWMS'],
            COLUMN_NAMES['recv_system']: ['LZWMS', 'LYMES'],
            COLUMN_NAMES['if_name']: ['IF_001', 'IF_002']
        }
        df = pd.DataFrame(test_data)
        df.to_csv(self.test_csv_path, index=False)
        
        # DB 생성
        result = self.creator.create_database(self.test_csv_path)
        self.assertTrue(result)
        
        # 데이터 확인
        with sqlite3.connect(self.test_db_path) as conn:
            loaded_df = pd.read_sql_query(f"SELECT * FROM {self.creator.table_name}", conn)
            self.assertEqual(len(loaded_df), 2)
            self.assertEqual(loaded_df[COLUMN_NAMES['if_name']].tolist(), ['IF_001', 'IF_002'])
    
    def test_create_database_from_dataframe(self):
        """DataFrame으로부터 데이터베이스 생성 테스트"""
        # 테스트 DataFrame
        test_data = {
            COLUMN_NAMES['send_system']: ['LHMES'],
            COLUMN_NAMES['recv_system']: ['VOWMS'],
            COLUMN_NAMES['if_name']: ['IF_TEST']
        }
        df = pd.DataFrame(test_data)
        
        # DB 생성
        result = self.creator.create_database(df)
        self.assertTrue(result)
        
        # 검증
        verify_result = self.creator.verify_database()
        self.assertTrue(verify_result['success'])
        self.assertEqual(verify_result['row_count'], 1)
    
    def test_verify_database(self):
        """데이터베이스 검증 테스트"""
        # DB 없을 때
        verify_result = self.creator.verify_database()
        self.assertFalse(verify_result['success'])
        
        # DB 생성 후
        self.creator.create_test_database()
        verify_result = self.creator.verify_database()
        self.assertTrue(verify_result['success'])
        self.assertIn('columns', verify_result)
        self.assertIn('row_count', verify_result)
    
    def test_generate_test_data(self):
        """테스트 데이터 생성 테스트"""
        # private 메서드 테스트
        df = self.creator._generate_test_data()
        
        # 데이터 구조 확인
        self.assertIsInstance(df, pd.DataFrame)
        
        # 행 수 확인
        expected_rows = sum(TEST_CONFIG['sample_rows'].values())
        self.assertEqual(len(df), expected_rows)
        
        # LY/LZ 시스템 데이터 확인
        ly_lz_mask = (
            df[COLUMN_NAMES['send_system']].str.contains('LY|LZ') |
            df[COLUMN_NAMES['recv_system']].str.contains('LY|LZ')
        )
        ly_lz_count = ly_lz_mask.sum()
        self.assertEqual(ly_lz_count, TEST_CONFIG['sample_rows']['ly_lz_systems'])
        
        # LH/VO 시스템 데이터 확인
        lh_vo_mask = (
            df[COLUMN_NAMES['send_system']].str.contains('LH|VO') |
            df[COLUMN_NAMES['recv_system']].str.contains('LH|VO')
        )
        lh_vo_count = lh_vo_mask.sum()
        self.assertEqual(lh_vo_count, TEST_CONFIG['sample_rows']['lh_vo_systems'])


if __name__ == '__main__':
    unittest.main()