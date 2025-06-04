"""
BW Tools YAML Processor 단위 테스트
"""

import unittest
import os
import yaml
import pandas as pd
from bwtools_yaml_processor import YAMLProcessor
from bwtools_config import COLUMN_NAMES, ADDITIONAL_COLUMNS, REPLACEMENT_RULES

class TestYAMLProcessor(unittest.TestCase):
    def setUp(self):
        """테스트 설정"""
        self.processor = YAMLProcessor()
        self.test_csv_path = 'test_input.csv'
        self.test_yaml_path = 'test_rules.yaml'
        self.test_log_path = 'test_log.txt'
        self.test_result_path = 'test_result.xlsx'
        self.test_batch_path = 'test_delete.bat'
        
        # 테스트 CSV 생성
        self._create_test_csv()
        
    def tearDown(self):
        """테스트 정리"""
        # 테스트 파일 삭제
        test_files = [
            self.test_csv_path, self.test_yaml_path, self.test_log_path,
            self.test_result_path, self.test_batch_path
        ]
        for file in test_files:
            if os.path.exists(file):
                os.remove(file)
    
    def _create_test_csv(self):
        """테스트용 CSV 파일 생성"""
        # 2줄 단위 데이터 (기준행, 매칭행)
        data = []
        
        # 첫 번째 쌍 - 기준행 (LY/LZ)
        data.append({
            COLUMN_NAMES['send_system']: 'LYMES',
            COLUMN_NAMES['recv_system']: 'LZWMS',
            ADDITIONAL_COLUMNS['send_file_path']: '/home/lycorp/test_ly_send.process',
            ADDITIONAL_COLUMNS['recv_file_path']: '/home/lzcorp/test_lz_recv.process',
            ADDITIONAL_COLUMNS['send_schema_file']: '/home/lycorp/test_ly.xsd',
            ADDITIONAL_COLUMNS['recv_schema_file']: '/home/lzcorp/test_lz.xsd'
        })
        
        # 첫 번째 쌍 - 매칭행 (LH/VO)
        data.append({
            COLUMN_NAMES['send_system']: 'LHMES',
            COLUMN_NAMES['recv_system']: 'VOWMS',
            ADDITIONAL_COLUMNS['send_file_path']: '/home/lhcorp/test_lh_send.process',
            ADDITIONAL_COLUMNS['recv_file_path']: '/home/vocorp/test_vo_recv.process',
            ADDITIONAL_COLUMNS['send_schema_file']: '/home/lhcorp/test_lh.xsd',
            ADDITIONAL_COLUMNS['recv_schema_file']: '/home/vocorp/test_vo.xsd'
        })
        
        df = pd.DataFrame(data)
        df.to_csv(self.test_csv_path, index=False)
    
    def test_read_input_file(self):
        """입력 파일 읽기 테스트"""
        df = self.processor._read_input_file(self.test_csv_path)
        self.assertIsInstance(df, pd.DataFrame)
        self.assertEqual(len(df), 2)
    
    def test_generate_yaml_from_excel(self):
        """Excel/CSV에서 YAML 생성 테스트"""
        result = self.processor.generate_yaml_from_excel(
            self.test_csv_path, self.test_yaml_path
        )
        self.assertTrue(result)
        self.assertTrue(os.path.exists(self.test_yaml_path))
        
        # YAML 파일 읽기
        with open(self.test_yaml_path, 'r', encoding='utf-8') as f:
            yaml_data = yaml.safe_load(f)
        
        self.assertIsNotNone(yaml_data)
        self.assertIn('row_1', yaml_data)
        
        # 구조 확인
        row_1 = yaml_data['row_1']
        self.assertIn('send_file', row_1)
        self.assertIn('원본파일', row_1['send_file'])
        self.assertIn('복사파일', row_1['send_file'])
        self.assertIn('치환목록', row_1['send_file'])
    
    def test_generate_replacement_rules(self):
        """치환 규칙 생성 테스트"""
        base_row = pd.Series({
            COLUMN_NAMES['send_system']: 'LYMES',
            COLUMN_NAMES['recv_system']: 'LZWMS',
            COLUMN_NAMES['ems_name']: 'EMS_LY_LZ',
            COLUMN_NAMES['group_id']: '001'
        })
        
        matched_row = pd.Series({
            COLUMN_NAMES['send_system']: 'LHMES',
            COLUMN_NAMES['recv_system']: 'VOWMS',
            COLUMN_NAMES['ems_name']: 'EMS_LH_VO',
            COLUMN_NAMES['group_id']: '001'
        })
        
        rules = self.processor._generate_replacement_rules(
            base_row, matched_row, 'send_file'
        )
        
        self.assertIsInstance(rules, list)
        self.assertGreater(len(rules), 0)
        
        # 기본 시스템 치환 규칙 확인
        rule_descriptions = [r['설명'] for r in rules]
        self.assertTrue(any('LHMES_MGR → LYMES_MGR' in desc for desc in rule_descriptions))
    
    def test_create_yaml_structure(self):
        """YAML 구조 생성 테스트"""
        df = pd.read_csv(self.test_csv_path)
        yaml_data = self.processor._create_yaml_structure(df)
        
        self.assertIsInstance(yaml_data, dict)
        self.assertIn('row_1', yaml_data)
        
        # 파일 타입별 데이터 확인
        row_data = yaml_data['row_1']
        for file_type in ['send_file', 'recv_file', 'send_schema', 'recv_schema']:
            if file_type in row_data:
                self.assertIn('원본파일', row_data[file_type])
                self.assertIn('복사파일', row_data[file_type])
                self.assertIn('치환목록', row_data[file_type])
    
    def test_execute_replacements_without_files(self):
        """파일이 없을 때 치환 실행 테스트"""
        # YAML 생성
        self.processor.generate_yaml_from_excel(
            self.test_csv_path, self.test_yaml_path
        )
        
        # 치환 실행 (실제 파일이 없으므로 실패 예상)
        result = self.processor.execute_replacements(
            self.test_yaml_path,
            self.test_log_path,
            self.test_result_path
        )
        
        # 실행은 성공하지만 개별 파일 처리는 실패
        self.assertTrue(result)
        
        # 로그 파일 확인
        self.assertTrue(os.path.exists(self.test_log_path))
        with open(self.test_log_path, 'r', encoding='utf-8') as f:
            log_content = f.read()
            self.assertIn('실패', log_content)
    
    def test_save_log_file(self):
        """로그 파일 저장 테스트"""
        # 테스트 로그 엔트리 추가
        self.processor.log_entries = [
            {
                'timestamp': '2024-01-01 12:00:00',
                'row': 'row_1',
                'file_type': 'send_file',
                'source': '/source/file.txt',
                'dest': '/dest/file.txt',
                'success': True,
                'replacements': 5
            }
        ]
        
        self.processor._save_log_file(self.test_log_path)
        self.assertTrue(os.path.exists(self.test_log_path))
        
        # 로그 내용 확인
        with open(self.test_log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn('BW Tools 치환 작업 로그', content)
            self.assertIn('성공', content)
            self.assertIn('/source/file.txt', content)
    
    def test_generate_delete_batch(self):
        """삭제 배치 파일 생성 테스트"""
        # 테스트 파일 목록
        self.processor.copied_files = [
            'C:\\test\\file1.txt',
            'C:\\test\\file2.txt'
        ]
        
        self.processor._generate_delete_batch(self.test_batch_path)
        self.assertTrue(os.path.exists(self.test_batch_path))
        
        # 배치 파일 내용 확인
        with open(self.test_batch_path, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn('@echo off', content)
            self.assertIn('chcp 65001', content)
            self.assertIn('del /f /q', content)
            self.assertIn('file1.txt', content)
            self.assertIn('file2.txt', content)
            self.assertIn('attrib -r -h -s', content)
    
    def test_apply_replacements(self):
        """치환 적용 테스트"""
        # 임시 파일 생성
        temp_file = 'test_temp.txt'
        original_content = 'LHMES_MGR test LH system'
        
        with open(temp_file, 'w', encoding='utf-8') as f:
            f.write(original_content)
        
        # 치환 규칙
        replacements = [
            {
                '찾기': {'정규식': 'LHMES_MGR'},
                '교체': {'값': 'LYMES_MGR'}
            },
            {
                '찾기': {'정규식': 'LH'},
                '교체': {'값': 'LY'}
            }
        ]
        
        # 치환 적용
        self.processor._apply_replacements(temp_file, replacements)
        
        # 결과 확인
        with open(temp_file, 'r', encoding='utf-8') as f:
            new_content = f.read()
        
        self.assertEqual(new_content, 'LYMES_MGR test LY system')
        
        # 정리
        os.remove(temp_file)


if __name__ == '__main__':
    unittest.main()