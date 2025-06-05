"""
리팩토링된 인터페이스 도구 통합 실행 파일

모든 기능을 하나의 메뉴에서 실행할 수 있는 통합 도구입니다.
string_replacer.py와 유사한 메뉴 구조를 제공합니다.
"""

import os
import sys
import datetime
from typing import Optional

# 리팩토링된 모듈들 import
try:
    from rft_interface_processor import InterfaceProcessor
    from rft_yaml_processor import YAMLProcessor
    from rft_interface_reader import InterfaceExcelReader, BWProcessFileParser
    from test_rft_modules import TestRFTModules
except ImportError as e:
    print(f"모듈 import 오류: {e}")
    print("모든 리팩토링된 모듈이 같은 디렉토리에 있는지 확인하세요.")
    sys.exit(1)


class RFTMainController:
    """리팩토링된 도구의 메인 컨트롤러"""
    
    def __init__(self):
        """메인 컨트롤러 초기화"""
        self.interface_processor = InterfaceProcessor()
        self.yaml_processor = YAMLProcessor()
        self.interface_reader = InterfaceExcelReader()
        self.bw_parser = BWProcessFileParser()
        self.tester = TestRFTModules()
        
        # 기본 파일 경로들
        self.default_db = "iflist.sqlite"
        self.default_output_csv = "rft_interface_processed.csv"
        self.default_yaml = "rft_rules.yaml"
        
    def show_main_menu(self):
        """메인 메뉴 표시"""
        print("\n" + "=" * 80)
        print("리팩토링된 인터페이스 처리 도구 (RFT - Refactored Tools)")
        print("=" * 80)
        print("1. 인터페이스 데이터 처리 (복사행 기준으로 매칭행 찾기)")
        print("2. YAML 생성 (Excel/CSV → YAML)")
        print("3. YAML 실행 (파일 복사 및 치환)")
        print("4. 인터페이스 정보 읽기")
        print("5. BW 프로세스 파일 파싱")
        print("6. 전체 파이프라인 실행")
        print("7. 테스트 실행")
        print("8. 도구 정보 및 도움말")
        print("0. 종료")
        print("=" * 80)
    
    
    def run_interface_processing(self):
        """1. 인터페이스 데이터 처리"""
        print("\n=== 인터페이스 데이터 처리 ===")
        print("SQLite 데이터베이스에서 LY/LZ 시스템을 찾아 LH/VO 시스템과 매칭합니다.")
        
        # 데이터베이스 파일 확인
        if not os.path.exists(self.default_db):
            print(f"오류: 데이터베이스 파일을 찾을 수 없습니다 - {self.default_db}")
            print("먼저 '1. Excel을 SQLite로 변환' 기능을 실행하세요.")
            return
        
        output_file = input(f"출력 파일명 (Enter: {self.default_output_csv}): ").strip()
        if not output_file:
            output_file = self.default_output_csv
        
        success = self.interface_processor.process_interface_data(output_file)
        if success:
            print(f"✓ 인터페이스 처리 완료: {output_file}")
        else:
            print("✗ 인터페이스 처리 실패")
    
    def run_yaml_generation(self):
        """2. YAML 생성"""
        print("\n=== YAML 생성 ===")
        print("Excel/CSV 파일을 읽어 파일 복사 및 치환 규칙이 담긴 YAML을 생성합니다.")
        
        excel_path = input("입력 Excel/CSV 파일 경로를 입력하세요: ").strip()
        if not excel_path:
            print("파일 경로를 입력해야 합니다.")
            return
        
        if not os.path.exists(excel_path):
            print(f"오류: 파일을 찾을 수 없습니다 - {excel_path}")
            return
        
        yaml_path = input(f"출력 YAML 파일명 (Enter: {self.default_yaml}): ").strip()
        if not yaml_path:
            yaml_path = self.default_yaml
        
        success = self.yaml_processor.generate_yaml_from_excel(excel_path, yaml_path)
        if success:
            print(f"✓ YAML 생성 완료: {yaml_path}")
        else:
            print("✗ YAML 생성 실패")
    
    def run_yaml_execution(self):
        """3. YAML 실행"""
        print("\n=== YAML 실행 ===")
        print("YAML 파일에 정의된 파일 복사 및 치환 작업을 실행합니다.")
        print("⚠️  주의: 실제 파일이 복사되고 수정됩니다!")
        
        yaml_path = input("YAML 파일 경로를 입력하세요: ").strip()
        if not yaml_path:
            print("YAML 파일 경로를 입력해야 합니다.")
            return
        
        if not os.path.exists(yaml_path):
            print(f"오류: YAML 파일을 찾을 수 없습니다 - {yaml_path}")
            return
        
        # 확인 메시지
        confirm = input("정말로 파일 복사 및 치환을 실행하시겠습니까? (y/N): ").strip().lower()
        if confirm != 'y':
            print("실행이 취소되었습니다.")
            return
        
        log_path = input("로그 파일 경로 (Enter: 자동생성): ").strip()
        result_path = input("결과 파일 경로 (Enter: 자동생성): ").strip()
        
        if not log_path:
            log_path = None
        if not result_path:
            result_path = None
        
        success = self.yaml_processor.execute_replacements(yaml_path, log_path, result_path)
        if success:
            print("✓ YAML 실행 완료")
        else:
            print("✗ YAML 실행 실패")
    
    def run_interface_reading(self):
        """4. 인터페이스 정보 읽기"""
        print("\n=== 인터페이스 정보 읽기 ===")
        print("특별한 형식의 Excel 파일에서 인터페이스 정보를 읽습니다.")
        
        excel_path = input("인터페이스 Excel 파일 경로를 입력하세요: ").strip()
        if not excel_path:
            print("파일 경로를 입력해야 합니다.")
            return
        
        if not os.path.exists(excel_path):
            print(f"오류: 파일을 찾을 수 없습니다 - {excel_path}")
            return
        
        interfaces = self.interface_reader.read_excel(excel_path)
        if interfaces:
            print(f"✓ {len(interfaces)}개의 인터페이스가 로드되었습니다.")
            
            # 요약 정보 표시
            summary = self.interface_reader.get_interface_summary()
            print(f"  - 총 컬럼 수: {summary['total_columns']}")
            print(f"  - 처리 성공: {summary['processed_count']}")
            print(f"  - 오류 발생: {summary['error_count']}")
            
            # CSV 내보내기 옵션
            export = input("CSV로 내보내시겠습니까? (y/N): ").strip().lower()
            if export == 'y':
                output_path = input("출력 CSV 파일 경로 (Enter: rft_interface_export.csv): ").strip()
                if not output_path:
                    output_path = "rft_interface_export.csv"
                
                if self.interface_reader.export_to_csv(output_path):
                    print(f"✓ CSV 내보내기 완료: {output_path}")
        else:
            print("✗ 인터페이스 정보 읽기 실패")
    
    def run_bw_parsing(self):
        """5. BW 프로세스 파일 파싱"""
        print("\n=== BW 프로세스 파일 파싱 ===")
        print("TIBCO BW .process 파일에서 INSERT 쿼리와 파라미터를 추출합니다.")
        
        process_path = input("BW 프로세스 파일 경로를 입력하세요: ").strip()
        if not process_path:
            print("파일 경로를 입력해야 합니다.")
            return
        
        if not os.path.exists(process_path):
            print(f"오류: 파일을 찾을 수 없습니다 - {process_path}")
            return
        
        result = self.bw_parser.parse_process_file(process_path)
        if result:
            print(f"✓ 파싱 완료:")
            print(f"  - INSERT 쿼리: {len(result['insert_queries'])}개")
            print(f"  - 파라미터: {len(result['parameters'])}개")
            print(f"  - 활동: {len(result['activities'])}개")
            
            # 결과 내보내기 옵션
            export = input("CSV로 내보내시겠습니까? (y/N): ").strip().lower()
            if export == 'y':
                output_path = input("출력 CSV 파일 경로 (Enter: rft_bw_parsing_export.csv): ").strip()
                if not output_path:
                    output_path = "rft_bw_parsing_export.csv"
                
                if self.bw_parser.export_parsing_results(output_path):
                    print(f"✓ CSV 내보내기 완료: {output_path}")
        else:
            print("✗ BW 프로세스 파일 파싱 실패")
    
    def run_full_pipeline(self):
        """6. 전체 파이프라인 실행"""
        print("\n=== 전체 파이프라인 실행 ===")
        print("Excel → SQLite → 인터페이스 처리 → YAML 생성 순서로 실행됩니다.")
        
        excel_path = input("입력 Excel 파일 경로를 입력하세요: ").strip()
        if not excel_path:
            print("파일 경로를 입력해야 합니다.")
            return
        
        if not os.path.exists(excel_path):
            print(f"오류: 파일을 찾을 수 없습니다 - {excel_path}")
            return
        
        print("\n주의: iflist.sqlite 데이터베이스가 이미 존재한다고 가정합니다.")
        
        print("\n1단계: 인터페이스 데이터 처리")
        if not self.interface_processor.process_interface_data(self.default_output_csv):
            print("✗ 1단계 실패")
            return
        print("✓ 1단계 완료")
        
        print("\n2단계: YAML 생성")
        if not self.yaml_processor.generate_yaml_from_excel(self.default_output_csv, self.default_yaml):
            print("✗ 2단계 실패")
            return
        print("✓ 2단계 완료")
        
        print("\n✓ 전체 파이프라인 실행 완료!")
        print(f"  - 데이터베이스: {self.default_db}")
        print(f"  - 처리된 데이터: {self.default_output_csv}")
        print(f"  - YAML 규칙: {self.default_yaml}")
        print("\n다음 단계: '4. YAML 실행'을 통해 실제 파일 복사 및 치환을 수행하세요.")
    
    def run_tests(self):
        """7. 테스트 실행"""
        print("\n=== 테스트 실행 ===")
        
        while True:
            print("\n8-1. 모든 테스트 실행")
            print("8-2. Excel to SQLite 테스트")
            print("8-3. 인터페이스 처리 테스트")
            print("8-4. YAML 처리 테스트")
            print("8-5. 인터페이스 읽기 테스트")
            print("0. 이전 메뉴로")
            
            choice = input("\n선택하세요: ").strip()
            
            if choice == "8-1":
                self.tester.run_all_tests()
                
            elif choice == "8-2":
                self.tester.setup_test_environment()
                try:
                    self.tester.test_excel_to_sqlite()
                finally:
                    self.tester.cleanup_test_environment()
                    
            elif choice == "8-3":
                self.tester.setup_test_environment()
                try:
                    self.tester.test_interface_processor()
                finally:
                    self.tester.cleanup_test_environment()
                    
            elif choice == "8-4":
                self.tester.setup_test_environment()
                try:
                    self.tester.test_yaml_processor()
                finally:
                    self.tester.cleanup_test_environment()
                    
            elif choice == "8-5":
                self.tester.setup_test_environment()
                try:
                    self.tester.test_interface_reader()
                finally:
                    self.tester.cleanup_test_environment()
                    
            elif choice == "0":
                break
            else:
                print("잘못된 선택입니다.")
    
    def show_help(self):
        """8. 도구 정보 및 도움말"""
        print("\n=== 도구 정보 및 도움말 ===")
        print(f"리팩토링된 인터페이스 처리 도구 v1.0")
        print(f"작성일: {datetime.datetime.now().strftime('%Y-%m-%d')}")
        print("\n개요:")
        print("  TIBCO BW 인터페이스 마이그레이션을 위한 통합 도구입니다.")
        print("  기존의 개별 스크립트들을 리팩토링하여 하나의 통합된 도구로 제공합니다.")
        
        print("\n주요 기능:")
        print("  1. LY/LZ 시스템을 LH/VO 시스템과 자동 매칭")
        print("  2. 파일 복사 및 내용 치환을 위한 YAML 규칙 생성")
        print("  3. YAML 기반 자동 파일 처리")
        print("  4. BW 프로세스 파일 분석")
        print("  5. 포괄적인 테스트 지원")
        
        print("\n사용 순서:")
        print("  1) Excel 파일 준비 (인터페이스 목록)")
        print("  2) iflist.sqlite 데이터베이스가 준비되어 있어야 함")
        print("  3) '1. 인터페이스 데이터 처리' 실행")
        print("  4) '2. YAML 생성' 실행")
        print("  5) '3. YAML 실행' 실행 (실제 파일 처리)")
        print("  또는 '6. 전체 파이프라인 실행'으로 자동 실행")
        
        print("\n파일 구조:")
        print("  - rft_interface_processor.py: 인터페이스 데이터 처리")
        print("  - rft_yaml_processor.py: YAML 생성 및 실행")
        print("  - rft_interface_reader.py: 인터페이스 정보 읽기")
        print("  - test_rft_modules.py: 단위 테스트")
        print("  - rft_main.py: 통합 실행 파일 (현재 파일)")
        
        print("\n주의사항:")
        print("  - YAML 실행 시 실제 파일이 복사되고 수정됩니다.")
        print("  - 중요한 파일은 미리 백업하세요.")
        print("  - 테스트 기능을 활용하여 동작을 확인하세요.")
    
    def run(self):
        """메인 실행 루프"""
        print("리팩토링된 인터페이스 처리 도구를 시작합니다...")
        
        while True:
            try:
                self.show_main_menu()
                choice = input("\n원하는 작업을 선택하세요: ").strip()
                
                if choice == "1":
                    self.run_interface_processing()
                elif choice == "2":
                    self.run_yaml_generation()
                elif choice == "3":
                    self.run_yaml_execution()
                elif choice == "4":
                    self.run_interface_reading()
                elif choice == "5":
                    self.run_bw_parsing()
                elif choice == "6":
                    self.run_full_pipeline()
                elif choice == "7":
                    self.run_tests()
                elif choice == "8":
                    self.show_help()
                elif choice == "0":
                    print("\n프로그램을 종료합니다.")
                    break
                else:
                    print("잘못된 선택입니다. 다시 시도하세요.")
                    
            except KeyboardInterrupt:
                print("\n\n프로그램이 중단되었습니다.")
                break
            except Exception as e:
                print(f"\n예상치 못한 오류가 발생했습니다: {str(e)}")
                print("계속 진행하려면 Enter를 누르세요...")
                input()


def main():
    """메인 함수"""
    # 현재 디렉토리 확인
    current_dir = os.getcwd()
    print(f"현재 작업 디렉토리: {current_dir}")
    
    # 필요한 모듈 파일들이 있는지 확인
    required_files = [
        "rft_interface_processor.py", 
        "rft_yaml_processor.py",
        "rft_interface_reader.py",
        "test_rft_modules.py"
    ]
    
    missing_files = [f for f in required_files if not os.path.exists(f)]
    if missing_files:
        print(f"오류: 다음 파일들을 찾을 수 없습니다:")
        for f in missing_files:
            print(f"  - {f}")
        print("모든 리팩토링된 모듈이 같은 디렉토리에 있는지 확인하세요.")
        return
    
    # 메인 컨트롤러 실행
    controller = RFTMainController()
    controller.run()


if __name__ == "__main__":
    main()