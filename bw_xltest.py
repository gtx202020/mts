#!/usr/bin/env python3
"""
BW-XLTEST: 데이터베이스 컬럼 검증 도구
송신/수신 테이블 간 컬럼 호환성을 검증하고 결과를 Excel로 출력합니다.
"""

import sys
import logging
from typing import Dict, List, Any, Optional
from datetime import datetime

# 모듈 임포트
from bw_xltest_core import DatabaseHandler, ColumnValidator
from bw_xltest_io import ExcelReader, ExcelWriter

# 상수 정의
DEFAULT_INPUT_FILE = 'input.xlsx'
DEFAULT_OUTPUT_FILE = 'output_bw.xlsx'
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'

# 로깅 설정
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
logger = logging.getLogger(__name__)


def process_interface(interface_info: Dict[str, Any], db_handler: DatabaseHandler, 
                     validator: ColumnValidator) -> Dict[str, Any]:
    """단일 인터페이스 처리
    
    Args:
        interface_info: 인터페이스 정보
        db_handler: 데이터베이스 핸들러
        validator: 컬럼 검증기
    
    Returns:
        처리 결과
    """
    result = {
        'success': False,
        'validation_results': [],
        'errors': []
    }
    
    try:
        # 송신 DB 연결
        send_db = interface_info['send']['db_info']
        send_conn = db_handler.connect_db(
            send_db['sid'], 
            send_db['username'], 
            send_db['password'], 
            'send'
        )
        
        if not send_conn:
            result['errors'].append("송신 DB 연결 실패")
            return result
        
        # 수신 DB 연결
        recv_db = interface_info['recv']['db_info']
        recv_conn = db_handler.connect_db(
            recv_db['sid'], 
            recv_db['username'], 
            recv_db['password'], 
            'recv'
        )
        
        if not recv_conn:
            result['errors'].append("수신 DB 연결 실패")
            return result
        
        # 송신 테이블 컬럼 정보 조회
        send_table = interface_info['send']['table_info']
        send_columns = db_handler.get_column_info(
            send_table['owner'], 
            send_table['table_name'], 
            'send'
        )
        
        if not send_columns:
            result['errors'].append("송신 테이블 컬럼 정보 조회 실패")
            return result
        
        # 수신 테이블 컬럼 정보 조회
        recv_table = interface_info['recv']['table_info']
        recv_columns = db_handler.get_column_info(
            recv_table['owner'], 
            recv_table['table_name'], 
            'recv'
        )
        
        if not recv_columns:
            result['errors'].append("수신 테이블 컬럼 정보 조회 실패")
            return result
        
        # 컬럼 검증 수행
        validation_results = validator.validate_columns(
            interface_info['send']['columns'],
            interface_info['recv']['columns'],
            send_columns,
            recv_columns
        )
        
        result['validation_results'] = validation_results
        result['success'] = True
        
        # 결과 요약 로그
        error_count = sum(1 for r in validation_results if r['status'] == '오류')
        warning_count = sum(1 for r in validation_results if r['status'] == '경고')
        logger.info(f"검증 완료: 오류 {error_count}건, 경고 {warning_count}건")
        
    except Exception as e:
        logger.error(f"인터페이스 처리 중 오류: {str(e)}")
        result['errors'].append(str(e))
    
    finally:
        # DB 연결 종료
        db_handler.close_connections()
    
    return result


def main(input_file: str = DEFAULT_INPUT_FILE, output_file: str = DEFAULT_OUTPUT_FILE):
    """메인 실행 함수
    
    Args:
        input_file: 입력 Excel 파일 경로
        output_file: 출력 Excel 파일 경로
    """
    start_time = datetime.now()
    logger.info("=" * 50)
    logger.info("BW-XLTEST 시작")
    logger.info(f"입력 파일: {input_file}")
    logger.info(f"출력 파일: {output_file}")
    logger.info("=" * 50)
    
    # Excel Reader 초기화
    reader = ExcelReader(input_file)
    if not reader.open_workbook():
        logger.error("입력 파일을 열 수 없습니다.")
        return 1
    
    # Excel Writer 초기화
    writer = ExcelWriter(output_file)
    writer.create_workbook()
    
    # 처리 통계
    total_interfaces = 0
    success_count = 0
    error_interfaces = []
    
    # 인터페이스별 처리
    current_col = 2  # B열부터 시작
    
    while True:
        try:
            # 인터페이스 정보 읽기
            interface_info = reader.read_interface_info(current_col)
            if not interface_info:
                break
            
            total_interfaces += 1
            interface_name = interface_info.get('interface_name', f'Interface_{total_interfaces}')
            
            logger.info(f"\n처리 중: {interface_name} (인터페이스 {total_interfaces})")
            
            # 데이터베이스 핸들러와 검증기 생성
            db_handler = DatabaseHandler()
            validator = ColumnValidator()
            
            # 인터페이스 처리
            result = process_interface(interface_info, db_handler, validator)
            
            if result['success']:
                # 결과를 Excel에 기록
                writer.write_interface_result(
                    interface_info,
                    result['validation_results'],
                    total_interfaces
                )
                success_count += 1
                logger.info(f"{interface_name} 처리 완료")
            else:
                error_interfaces.append({
                    'name': interface_name,
                    'errors': result['errors']
                })
                logger.error(f"{interface_name} 처리 실패: {', '.join(result['errors'])}")
            
        except Exception as e:
            logger.error(f"인터페이스 읽기 오류 (열 {current_col}): {str(e)}")
            error_interfaces.append({
                'name': f'Column_{current_col}',
                'errors': [str(e)]
            })
        
        # 다음 인터페이스로 이동 (3열씩)
        current_col += 3
        
        # 최대 컬럼 수 체크 (무한 루프 방지)
        if current_col > 100:  # 안전장치
            logger.warning("최대 컬럼 수 도달, 처리 중단")
            break
    
    # 파일 저장 및 정리
    reader.close_workbook()
    writer.save_workbook()
    writer.close_workbook()
    
    # 처리 시간 계산
    end_time = datetime.now()
    elapsed_time = end_time - start_time
    
    # 최종 결과 출력
    logger.info("\n" + "=" * 50)
    logger.info("BW-XLTEST 완료")
    logger.info("=" * 50)
    logger.info(f"총 인터페이스 수: {total_interfaces}")
    logger.info(f"성공: {success_count}")
    logger.info(f"실패: {len(error_interfaces)}")
    logger.info(f"처리 시간: {elapsed_time}")
    
    if error_interfaces:
        logger.info("\n실패한 인터페이스:")
        for error in error_interfaces:
            logger.info(f"  - {error['name']}: {', '.join(error['errors'])}")
    
    logger.info(f"\n결과 파일: {output_file}")
    
    return 0 if len(error_interfaces) == 0 else 1


def print_usage():
    """사용법 출력"""
    print("사용법: python bw_xltest.py [입력파일] [출력파일]")
    print("  입력파일: 인터페이스 정보가 포함된 Excel 파일 (기본값: input.xlsx)")
    print("  출력파일: 검증 결과를 저장할 Excel 파일 (기본값: output_bw.xlsx)")
    print("\n예제:")
    print("  python bw_xltest.py")
    print("  python bw_xltest.py myinput.xlsx myoutput.xlsx")


if __name__ == "__main__":
    # 명령행 인자 처리
    if len(sys.argv) > 1:
        if sys.argv[1] in ['-h', '--help', '/?']:
            print_usage()
            sys.exit(0)
        
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUTPUT_FILE
    else:
        input_file = DEFAULT_INPUT_FILE
        output_file = DEFAULT_OUTPUT_FILE
    
    # 메인 함수 실행
    try:
        exit_code = main(input_file, output_file)
        sys.exit(exit_code)
    except KeyboardInterrupt:
        logger.info("\n사용자에 의해 중단됨")
        sys.exit(1)
    except Exception as e:
        logger.error(f"예상치 못한 오류: {str(e)}")
        sys.exit(1)