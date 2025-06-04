"""
BW Tools Main
모든 모듈을 통합하여 실행하는 메인 프로그램입니다.
"""

import os
import sys
import argparse
from typing import Optional
from bwtools_db_creator import DBCreator
from bwtools_excel_generator import ExcelGenerator
from bwtools_yaml_processor import YAMLProcessor
from bwtools_config import TEST_CONFIG

class BWToolsPipeline:
    def __init__(self):
        """BWToolsPipeline 초기화"""
        self.db_creator = DBCreator()
        self.excel_generator = ExcelGenerator()
        self.yaml_processor = YAMLProcessor()
        
    def run_full_pipeline(self, 
                         input_excel: Optional[str] = None,
                         use_test_data: bool = False,
                         output_format: str = 'xlsx') -> bool:
        """
        전체 파이프라인을 실행합니다.
        
        Args:
            input_excel: 입력 Excel 파일 경로 (None이면 테스트 데이터 사용)
            use_test_data: 테스트 데이터 사용 여부
            output_format: 출력 형식 ('xlsx' 또는 'csv')
            
        Returns:
            성공 여부
        """
        print("=" * 80)
        print("BW Tools 통합 파이프라인 시작")
        print("=" * 80)
        
        try:
            # 1단계: 데이터베이스 생성
            print("\n[1단계] 데이터베이스 생성")
            if use_test_data or not input_excel:
                print("테스트 데이터로 데이터베이스 생성 중...")
                if not self.db_creator.create_test_database():
                    print("데이터베이스 생성 실패")
                    return False
            else:
                print(f"Excel 파일에서 데이터베이스 생성 중: {input_excel}")
                if not self.db_creator.create_database(input_excel):
                    print("데이터베이스 생성 실패")
                    return False
            
            # 2단계: Excel 생성
            print("\n[2단계] Excel/CSV 파일 생성")
            excel_output = f"bwtools_output.{output_format}"
            if not self.excel_generator.generate_excel(excel_output, output_format):
                print("Excel 생성 실패")
                return False
            
            # 3단계: YAML 생성
            print("\n[3단계] YAML 파일 생성")
            yaml_output = "bwtools_rules.yaml"
            if not self.yaml_processor.generate_yaml_from_excel(excel_output, yaml_output):
                print("YAML 생성 실패")
                return False
            
            # 4단계: 치환 실행 (선택적)
            print("\n[4단계] 치환 작업")
            print("주의: 실제 파일 복사 및 치환은 수동으로 실행하세요.")
            print(f"생성된 YAML 파일: {yaml_output}")
            print("치환을 실행하려면 다음 명령을 사용하세요:")
            print(f"python bwtools_main.py --mode execute --yaml {yaml_output}")
            
            print("\n" + "=" * 80)
            print("파이프라인 실행 완료!")
            print("=" * 80)
            
            return True
            
        except Exception as e:
            print(f"\n파이프라인 실행 중 오류 발생: {str(e)}")
            return False
    
    def run_individual_step(self, mode: str, **kwargs) -> bool:
        """
        개별 단계를 실행합니다.
        
        Args:
            mode: 실행 모드 ('db', 'excel', 'yaml', 'execute')
            **kwargs: 모드별 필요한 인자
            
        Returns:
            성공 여부
        """
        try:
            if mode == 'db':
                # 데이터베이스 생성
                input_file = kwargs.get('input')
                if input_file:
                    return self.db_creator.create_database(input_file)
                else:
                    return self.db_creator.create_test_database()
                    
            elif mode == 'excel':
                # Excel 생성
                output_path = kwargs.get('output')
                output_format = kwargs.get('format', 'xlsx')
                return self.excel_generator.generate_excel(output_path, output_format)
                
            elif mode == 'yaml':
                # YAML 생성
                input_excel = kwargs.get('input')
                output_yaml = kwargs.get('output')
                if not input_excel:
                    print("입력 Excel 파일을 지정하세요 (--input)")
                    return False
                return self.yaml_processor.generate_yaml_from_excel(input_excel, output_yaml)
                
            elif mode == 'execute':
                # 치환 실행
                yaml_file = kwargs.get('yaml')
                if not yaml_file:
                    print("YAML 파일을 지정하세요 (--yaml)")
                    return False
                log_path = kwargs.get('log')
                result_excel = kwargs.get('result')
                return self.yaml_processor.execute_replacements(yaml_file, log_path, result_excel)
                
            else:
                print(f"알 수 없는 모드: {mode}")
                return False
                
        except Exception as e:
            print(f"작업 실행 중 오류 발생: {str(e)}")
            return False


def main():
    """메인 실행 함수"""
    parser = argparse.ArgumentParser(
        description='BW Tools - TIBCO BW 인터페이스 마이그레이션 도구',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예제:
  # 전체 파이프라인 실행 (테스트 데이터)
  python bwtools_main.py --test
  
  # 전체 파이프라인 실행 (실제 Excel 파일)
  python bwtools_main.py --input input.xlsx
  
  # 개별 단계 실행
  python bwtools_main.py --mode db --input data.xlsx
  python bwtools_main.py --mode excel --format csv
  python bwtools_main.py --mode yaml --input output.csv
  python bwtools_main.py --mode execute --yaml rules.yaml
        """
    )
    
    # 전체 파이프라인 옵션
    parser.add_argument('--input', help='입력 Excel 파일 경로')
    parser.add_argument('--test', action='store_true', help='테스트 데이터 사용')
    parser.add_argument('--format', choices=['xlsx', 'csv'], default='xlsx', 
                       help='출력 형식 (기본값: xlsx)')
    
    # 개별 단계 실행 옵션
    parser.add_argument('--mode', choices=['db', 'excel', 'yaml', 'execute'],
                       help='개별 단계 실행 모드')
    parser.add_argument('--output', help='출력 파일 경로')
    parser.add_argument('--yaml', help='YAML 파일 경로 (execute 모드)')
    parser.add_argument('--log', help='로그 파일 경로 (execute 모드)')
    parser.add_argument('--result', help='결과 Excel 파일 경로 (execute 모드)')
    
    args = parser.parse_args()
    
    # 파이프라인 생성
    pipeline = BWToolsPipeline()
    
    # 실행
    if args.mode:
        # 개별 단계 실행
        success = pipeline.run_individual_step(
            args.mode,
            input=args.input,
            output=args.output,
            format=args.format,
            yaml=args.yaml,
            log=args.log,
            result=args.result
        )
    else:
        # 전체 파이프라인 실행
        success = pipeline.run_full_pipeline(
            input_excel=args.input,
            use_test_data=args.test,
            output_format=args.format
        )
    
    # 종료 코드 반환
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()