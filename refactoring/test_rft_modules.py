"""
리팩토링된 모듈들의 단위 테스트

각 모듈의 주요 기능을 테스트합니다.
"""

import os
import tempfile
import shutil
import sqlite3
import pandas as pd
import yaml
from typing import Dict, List, Any

# 리팩토링된 모듈들 import
from rft_ex_sqlite import ExcelToSQLiteConverter
from rft_interface_processor import InterfaceProcessor
from rft_yaml_processor import YAMLProcessor
from rft_interface_reader import InterfaceExcelReader, BWProcessFileParser


class TestRFTModules:
    """리팩토링된 모듈들의 테스트 클래스"""
    
    def __init__(self):
        """테스트 클래스 초기화"""
        self.test_dir = None
        self.test_results = []
        
    def setup_test_environment(self):
        """테스트 환경 설정"""
        print("테스트 환경 설정 중...")
        
        # 임시 디렉토리 생성
        self.test_dir = tempfile.mkdtemp(prefix="rft_test_")
        print(f"테스트 디렉토리: {self.test_dir}")
        
        # 테스트용 Excel 데이터 생성
        self.create_test_excel()
        
        return True
    
    def cleanup_test_environment(self):
        """테스트 환경 정리"""
        if self.test_dir and os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)
            print(f"테스트 디렉토리 정리 완료: {self.test_dir}")
    
    def create_test_excel(self):
        """테스트용 Excel 파일 생성"""
        test_data = {
            '송신시스템': ['LY_SYS1', 'LH_SYS1', 'LZ_SYS2', 'VO_SYS2'],
            '수신시스템': ['LH_REC1', 'LH_REC1', 'VO_REC2', 'VO_REC2'],
            'I/F명': ['TEST_IF_001', 'TEST_IF_001', 'TEST_IF_002', 'TEST_IF_002'],
            '송신\n법인': ['KR', 'KR', 'NJ', 'NJ'],
            '수신\n법인': ['KR', 'KR', 'NJ', 'NJ'],
            '송신패키지': ['PKG_LY_001', 'PKG_LH_001', 'PKG_LZ_002', 'PKG_VO_002'],
            '수신패키지': ['PKG_LH_001', 'PKG_LH_001', 'PKG_VO_002', 'PKG_VO_002'],
            '송신\n업무명': ['PNL_LY', 'MES_LH', 'MOD_LZ', 'MES_VO'],
            '수신\n업무명': ['MES_LH', 'MES_LH', 'MES_VO', 'MES_VO'],
            'EMS명': ['MES01', 'MES01', 'MES02', 'MES02'],
            'Group ID': ['GRP01', 'GRP01', 'GRP02', 'GRP02'],
            'Event_ID': ['EVT001', 'EVT001', 'EVT002', 'EVT002'],
            '개발구분': ['신규', '신규', '수정', '수정'],
            'Source Table': ['LY.TB_TEST01', 'LH.TB_TEST01', 'LZ.TB_TEST02', 'VO.TB_TEST02'],
            'Destination Table': ['LH.TB_DEST01', 'LH.TB_DEST01', 'VO.TB_DEST02', 'VO.TB_DEST02'],
            'Routing': ['RT_LY_01', 'RT_LH_01', 'RT_LZ_02', 'RT_VO_02'],
            '스케쥴': ['매일 09:00', '매일 09:00', '매일 18:00', '매일 18:00'],
            '주기구분': ['일배치', '일배치', '일배치', '일배치'],
            '주기': ['Daily', 'Daily', 'Daily', 'Daily'],
            '송신\nDB Name': ['LYDB', 'LHDB', 'LZDB', 'VODB'],
            '송신 \nSchema': ['LYSCH', 'LHSCH', 'LZSCH', 'VOSCH']
        }
        
        df = pd.DataFrame(test_data)
        excel_path = os.path.join(self.test_dir, "test_input.xlsx")
        df.to_excel(excel_path, index=False)
        
        # CSV 버전도 생성
        csv_path = os.path.join(self.test_dir, "test_input.csv")
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        
        print(f"테스트 Excel 파일 생성: {excel_path}")
        print(f"테스트 CSV 파일 생성: {csv_path}")
    
    def log_test_result(self, test_name: str, success: bool, message: str = ""):
        """테스트 결과 로깅"""
        result = {
            'test_name': test_name,
            'success': success,
            'message': message
        }
        self.test_results.append(result)
        
        status = "PASS" if success else "FAIL"
        print(f"[{status}] {test_name}: {message}")
    
    def test_excel_to_sqlite(self) -> bool:
        """ExcelToSQLiteConverter 테스트"""
        test_name = "Excel to SQLite 변환"
        try:
            print(f"\n=== {test_name} 테스트 시작 ===")
            
            # 테스트용 데이터베이스 파일 경로
            db_path = os.path.join(self.test_dir, "test.sqlite")
            
            # 변환기 초기화
            converter = ExcelToSQLiteConverter(db_path)
            
            # Excel 파일 경로
            excel_path = os.path.join(self.test_dir, "test_input.xlsx")
            
            # 변환 실행
            success = converter.convert_excel_to_sqlite(excel_path)
            
            if not success:
                self.log_test_result(test_name, False, "변환 실행 실패")
                return False
            
            # 데이터베이스 파일 생성 확인
            if not os.path.exists(db_path):
                self.log_test_result(test_name, False, "데이터베이스 파일이 생성되지 않음")
                return False
            
            # 데이터 검증
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(*) FROM iflist")
            row_count = cursor.fetchone()[0]
            
            cursor.execute("PRAGMA table_info(iflist)")
            columns = cursor.fetchall()
            
            conn.close()
            
            if row_count != 4:
                self.log_test_result(test_name, False, f"예상 행 수: 4, 실제: {row_count}")
                return False
            
            if len(columns) < 10:
                self.log_test_result(test_name, False, f"컬럼 수가 부족함: {len(columns)}")
                return False
            
            self.log_test_result(test_name, True, f"행: {row_count}, 컬럼: {len(columns)}")
            return True
            
        except Exception as e:
            self.log_test_result(test_name, False, f"예외 발생: {str(e)}")
            return False
    
    def test_test_database_creation(self) -> bool:
        """테스트 데이터베이스 생성 테스트"""
        test_name = "테스트 데이터베이스 생성"
        try:
            print(f"\n=== {test_name} 테스트 시작 ===")
            
            db_path = os.path.join(self.test_dir, "test_db.sqlite")
            converter = ExcelToSQLiteConverter(db_path)
            
            success = converter.create_test_database()
            
            if not success:
                self.log_test_result(test_name, False, "테스트 데이터베이스 생성 실패")
                return False
            
            # 데이터 검증
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM iflist")
            row_count = cursor.fetchone()[0]
            conn.close()
            
            if row_count != 3:
                self.log_test_result(test_name, False, f"예상 행 수: 3, 실제: {row_count}")
                return False
            
            self.log_test_result(test_name, True, f"테스트 데이터 {row_count}행 생성")
            return True
            
        except Exception as e:
            self.log_test_result(test_name, False, f"예외 발생: {str(e)}")
            return False
    
    def test_interface_processor(self) -> bool:
        """InterfaceProcessor 테스트"""
        test_name = "인터페이스 데이터 처리"
        try:
            print(f"\n=== {test_name} 테스트 시작 ===")
            
            # 먼저 테스트 데이터베이스 생성
            db_path = os.path.join(self.test_dir, "processor_test.sqlite")
            converter = ExcelToSQLiteConverter(db_path)
            converter.create_test_database()
            
            # 프로세서 초기화
            processor = InterfaceProcessor(db_path)
            
            # 출력 파일 경로
            output_path = os.path.join(self.test_dir, "processed_output.csv")
            
            # 처리 실행
            success = processor.process_interface_data(output_path)
            
            if not success:
                self.log_test_result(test_name, False, "처리 실행 실패")
                return False
            
            # 출력 파일 확인
            if not os.path.exists(output_path):
                self.log_test_result(test_name, False, "출력 파일이 생성되지 않음")
                return False
            
            # 출력 데이터 검증
            df = pd.read_csv(output_path, encoding='utf-8-sig')
            
            if len(df) == 0:
                self.log_test_result(test_name, False, "출력 데이터가 비어있음")
                return False
            
            # 추가된 컬럼들 확인
            expected_columns = ['송신파일경로', '수신파일경로', '송신파일존재', '수신파일존재']
            missing_columns = [col for col in expected_columns if col not in df.columns]
            
            if missing_columns:
                self.log_test_result(test_name, False, f"누락된 컬럼: {missing_columns}")
                return False
            
            self.log_test_result(test_name, True, f"출력 행 수: {len(df)}, 컬럼 수: {len(df.columns)}")
            return True
            
        except Exception as e:
            self.log_test_result(test_name, False, f"예외 발생: {str(e)}")
            return False
    
    def test_yaml_processor(self) -> bool:
        """YAMLProcessor 테스트"""
        test_name = "YAML 처리"
        try:
            print(f"\n=== {test_name} 테스트 시작 ===")
            
            # 먼저 처리된 인터페이스 데이터 생성
            db_path = os.path.join(self.test_dir, "yaml_test.sqlite")
            converter = ExcelToSQLiteConverter(db_path)
            converter.create_test_database()
            
            processor = InterfaceProcessor(db_path)
            csv_path = os.path.join(self.test_dir, "yaml_input.csv")
            processor.process_interface_data(csv_path)
            
            # YAML 프로세서 테스트
            yaml_processor = YAMLProcessor(debug_mode=False)
            yaml_path = os.path.join(self.test_dir, "test_rules.yaml")
            
            # YAML 생성
            success = yaml_processor.generate_yaml_from_excel(csv_path, yaml_path)
            
            if not success:
                self.log_test_result(test_name, False, "YAML 생성 실패")
                return False
            
            # YAML 파일 확인
            if not os.path.exists(yaml_path):
                self.log_test_result(test_name, False, "YAML 파일이 생성되지 않음")
                return False
            
            # YAML 내용 검증
            with open(yaml_path, 'r', encoding='utf-8') as f:
                yaml_data = yaml.safe_load(f)
            
            if not yaml_data:
                self.log_test_result(test_name, False, "YAML 데이터가 비어있음")
                return False
            
            self.log_test_result(test_name, True, f"YAML 작업 수: {len(yaml_data)}")
            return True
            
        except Exception as e:
            self.log_test_result(test_name, False, f"예외 발생: {str(e)}")
            return False
    
    def test_interface_reader(self) -> bool:
        """InterfaceExcelReader 테스트"""
        test_name = "인터페이스 Excel 읽기"
        try:
            print(f"\n=== {test_name} 테스트 시작 ===")
            
            # 테스트용 인터페이스 Excel 생성
            interface_data = {
                'A': ['', '', '', '', 'col1', 'col2'],
                'B': ['Interface1', 'ID001', "{'host': 'localhost'}", "{'table': 'test'}", 'source1', 'source2'],
                'C': ['', '', '', '', 'target1', 'target2'],
                'D': ['', '', '', '', 'varchar', 'int'],
                'E': ['Interface2', 'ID002', "{'host': 'server'}", "{'table': 'test2'}", 'src1', 'src2'],
                'F': ['', '', '', '', 'tgt1', 'tgt2'],
                'G': ['', '', '', '', 'text', 'number']
            }
            
            df = pd.DataFrame(interface_data)
            excel_path = os.path.join(self.test_dir, "interface_test.xlsx")
            df.to_excel(excel_path, index=False, header=False)
            
            # 리더 테스트
            reader = InterfaceExcelReader()
            interfaces = reader.read_excel(excel_path)
            
            if not interfaces:
                self.log_test_result(test_name, False, "인터페이스 정보를 읽지 못함")
                return False
            
            expected_interfaces = ['Interface1', 'Interface2']
            for expected in expected_interfaces:
                if expected not in interfaces:
                    self.log_test_result(test_name, False, f"인터페이스 누락: {expected}")
                    return False
            
            # 컬럼 매핑 확인
            interface1 = interfaces['Interface1']
            if len(interface1['column_mappings']) != 2:
                self.log_test_result(test_name, False, f"Interface1 컬럼 매핑 수 오류: {len(interface1['column_mappings'])}")
                return False
            
            self.log_test_result(test_name, True, f"인터페이스 수: {len(interfaces)}")
            return True
            
        except Exception as e:
            self.log_test_result(test_name, False, f"예외 발생: {str(e)}")
            return False
    
    def run_all_tests(self) -> bool:
        """모든 테스트 실행"""
        print("=" * 80)
        print("리팩토링된 모듈 테스트 시작")
        print("=" * 80)
        
        # 테스트 환경 설정
        if not self.setup_test_environment():
            print("테스트 환경 설정 실패")
            return False
        
        try:
            # 각 테스트 실행
            tests = [
                self.test_excel_to_sqlite,
                self.test_test_database_creation,
                self.test_interface_processor,
                self.test_yaml_processor,
                self.test_interface_reader
            ]
            
            passed_tests = 0
            total_tests = len(tests)
            
            for test_func in tests:
                if test_func():
                    passed_tests += 1
            
            # 결과 요약
            print(f"\n" + "=" * 80)
            print("테스트 결과 요약")
            print("=" * 80)
            
            for result in self.test_results:
                status = "PASS" if result['success'] else "FAIL"
                print(f"[{status}] {result['test_name']}: {result['message']}")
            
            print(f"\n총 테스트: {total_tests}")
            print(f"성공: {passed_tests}")
            print(f"실패: {total_tests - passed_tests}")
            print(f"성공률: {(passed_tests / total_tests * 100):.1f}%")
            
            return passed_tests == total_tests
            
        finally:
            # 테스트 환경 정리
            self.cleanup_test_environment()


def main():
    """메인 실행 함수"""
    print("리팩토링된 모듈 테스트 도구")
    
    while True:
        print("\n메뉴:")
        print("1. 모든 테스트 실행")
        print("2. Excel to SQLite 테스트")
        print("3. 인터페이스 처리 테스트")
        print("4. YAML 처리 테스트")
        print("5. 인터페이스 읽기 테스트")
        print("0. 종료")
        
        choice = input("\n선택하세요: ").strip()
        
        tester = TestRFTModules()
        
        if choice == "1":
            tester.run_all_tests()
            
        elif choice == "2":
            tester.setup_test_environment()
            try:
                tester.test_excel_to_sqlite()
            finally:
                tester.cleanup_test_environment()
                
        elif choice == "3":
            tester.setup_test_environment()
            try:
                tester.test_interface_processor()
            finally:
                tester.cleanup_test_environment()
                
        elif choice == "4":
            tester.setup_test_environment()
            try:
                tester.test_yaml_processor()
            finally:
                tester.cleanup_test_environment()
                
        elif choice == "5":
            tester.setup_test_environment()
            try:
                tester.test_interface_reader()
            finally:
                tester.cleanup_test_environment()
                
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
            
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")


if __name__ == "__main__":
    main()