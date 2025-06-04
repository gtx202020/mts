"""
BW Tools 공통 설정 파일
모든 모듈에서 사용하는 공통 설정값들을 관리합니다.
"""

import os

# 데이터베이스 설정
DB_FILENAME = 'iflist.sqlite'
TABLE_NAME = 'iflist'

# 파일 경로 설정
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
TEST_DATA_PATH = os.path.join(BASE_PATH, 'test_data')

# 컬럼명 정의 (iflist03a.py와 정확히 일치)
COLUMN_NAMES = {
    # 핵심 필터링 컬럼
    'send_system': '송신시스템',
    'recv_system': '수신시스템',
    'if_name': 'I/F명',
    
    # 법인 정보 (개행문자 포함)
    'send_corp': '송신\n법인',
    'recv_corp': '수신\n법인',
    
    # 패키지 정보 (개행문자 없음)
    'send_pkg': '송신패키지',
    'recv_pkg': '수신패키지',
    
    # 업무명 (개행문자 포함)
    'send_task': '송신\n업무명',
    'recv_task': '수신\n업무명',
    
    # EMS 및 ID 정보
    'ems_name': 'EMS명',
    'group_id': 'Group ID',
    'event_id': 'Event_ID',
    
    # 데이터베이스 및 스키마 정보 (정확한 개행문자 위치)
    'send_db_name': '송신\nDB Name',
    'send_schema': '송신 \nSchema',
    'source_table': 'Source Table',
    'dest_table': 'Destination Table',
    
    # 기타 정보
    'dev_type': '개발구분',
    'routing': 'Routing',
    'cycle_type': '주기구분',
    'cycle': '주기',
    'schedule': '스케쥴'
}

# 추가 컬럼명
ADDITIONAL_COLUMNS = {
    'send_file_path': '송신파일경로',
    'recv_file_path': '수신파일경로',
    'send_file_exists': '송신파일존재',
    'recv_file_exists': '수신파일존재',
    'send_file_created': '송신파일생성여부',
    'recv_file_created': '수신파일생성여부',
    'send_df': '송신DF',
    'recv_df': '수신DF',
    'send_schema_file': '송신스키마파일명',
    'recv_schema_file': '수신스키마파일명',
    'send_schema_exists': '송신스키마파일존재',
    'recv_schema_exists': '수신스키마파일존재',
    'send_schema_created': '송신스키마파일생성여부',
    'recv_schema_created': '수신스키마파일생성여부',
    'compare_log': '비교로그'
}

# 시스템 변환 규칙
SYSTEM_MAPPING = {
    'LY': 'LH',
    'LZ': 'VO'
}

# 파일 경로 템플릿
FILE_PATH_TEMPLATES = {
    'send': '/home/{corp}/process/bw/Application/{pkg}/{task}/Process/{ems}_GRP_{group_id}_{event_id}_SND.process',
    'recv': '/home/{corp}/process/bw/Application/{pkg}/{task}/Process/{ems}_GRP_{group_id}_{event_id}_RCV.process',
    'send_schema': '/home/{corp}/process/bw/Application/{pkg}/{task}/SharedResources/{db_name}_{schema}_{table}.xsd',
    'recv_schema': '/home/{corp}/process/bw/Application/{pkg}/{task}/SharedResources/{table}.xsd'
}

# 치환 규칙 (string_replacer에서 사용)
REPLACEMENT_RULES = {
    'system': {
        'LHMES_MGR': 'LYMES_MGR',
        'VOMES_MGR': 'LZMES_MGR',
        'LH': 'LY',
        'VO': 'LZ'
    }
}

# Excel 출력 색상 설정
EXCEL_COLORS = {
    'match': 'yellow',
    'priority_filtered': 'lightgreen'
}

# 테스트 데이터 설정
TEST_CONFIG = {
    'use_csv': True,  # 테스트 시 CSV 사용 여부
    'sample_rows': {
        'ly_lz_systems': 8,  # LY/LZ 시스템 샘플 행 수
        'lh_vo_systems': 8,  # LH/VO 시스템 샘플 행 수
        'unmatched': 3      # 매칭되지 않는 행 수
    }
}