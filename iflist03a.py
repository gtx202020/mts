#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
iflist03a.py - 통합 데이터 처리 및 검증 프로그램

이 프로그램은 iflist03.py와 iflist04.py의 기능을 통합하여 다음과 같은 작업을 수행합니다:
1. SQLite 데이터베이스에서 'LY', 'LZ' 포함 데이터 추출
2. 매칭 행 찾기 및 우선순위 필터링
3. 기본행-매칭행 쌍에 대한 15가지 규칙 검증
4. 파일 경로 계산 및 존재 여부 확인
5. 결과를 단일 엑셀 파일로 저장

개발 버전: v1.0
"""

import os
import sys
import re
import sqlite3
import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 사용자 설정 옵션 (소스 내부에 정의)
# --------------------------------------
# 디버깅 모드 설정
# debug_mode = 0 또는 2: 최종 필터링된 행(연두색)만 표시
# debug_mode = 1: 모든 매칭 행(노란색)과 필터링된 행(연두색) 모두 표시
debug_mode = 0  # 기본값: 일반 모드 (최종 필터링된 행만 표시)

# 파일 경로 기본 설정
base_path = r"C:\BwProject"

# 검색 및 대체 문자열 정의
LY_PATTERN = 'LY'
LZ_PATTERN = 'LZ'
LH_REPLACEMENT = 'LH'
VO_REPLACEMENT = 'VO'

# 출력 엑셀 파일명 설정 (자동 생성)
output_file_name = f"iflist03a_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# 색상 정의
YELLOW_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # 매칭된 모든 행
GREEN_FILL = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')   # 우선순위 필터링된 행
ORANGE_FILL = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')  # 오류 있는 셀
BLUE_FILLS = {
    'very_light': PatternFill(start_color='DEEBF7', end_color='DEEBF7', fill_type='solid'),  # 1-3개
    'light': PatternFill(start_color='9ECAE1', end_color='9ECAE1', fill_type='solid'),       # 4-10개
    'medium': PatternFill(start_color='3182BD', end_color='3182BD', fill_type='solid'),      # 11-20개
    'dark': PatternFill(start_color='08519C', end_color='08519C', fill_type='solid')         # 21개 이상
}
GRAY_FILL = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')  # 0개

# ==================== 유틸리티 함수 ====================

def safe_str(value):
    """값을 안전하게 문자열로 변환"""
    if pd.isna(value):
        return ""
    return str(value)

def replace_ly_lz(text):
    """문자열에서 'LY'를 'LH'로, 'LZ'를 'VO'로 교체"""
    if not isinstance(text, str):
        return text
    result = text.replace(LY_PATTERN, LH_REPLACEMENT).replace(LZ_PATTERN, VO_REPLACEMENT)
    return result

def has_ly_lz(text):
    """문자열에 'LY' 또는 'LZ'가 포함되어 있는지 확인"""
    if not isinstance(text, str):
        return False
    return LY_PATTERN in text or LZ_PATTERN in text

def create_file_path(row, is_send=True):
    """행 데이터로부터 송신/수신 파일 경로 생성"""
    try:
        # 기본 경로 설정
        file_path = base_path

        # 필수 값 확인
        if pd.isna(row.get('법인')) or pd.isna(row.get('패키지')):
            return ""

        # 법인 정보에 따른 1번 디렉토리 결정
        corp_dir = safe_str(row['법인']).strip()
        if not corp_dir:
            corp_dir = "Unknown"

        # 환경 정보 (TEST/PROD)
        env_type = "TEST"  # 기본값
        if '운영' in safe_str(row.get('환경', '')):
            env_type = "PROD"

        # 1번 디렉토리 완성
        dir1 = f"{corp_dir}_{env_type}"
        file_path = os.path.join(file_path, dir1)

        # 패키지 정보로 2번 디렉토리 결정
        package_col = '송신패키지' if is_send else '수신패키지'
        dir2 = safe_str(row.get(package_col, row.get('패키지', 'Unknown'))).strip()
        if not dir2:
            dir2 = "Unknown"
        file_path = os.path.join(file_path, dir2)

        # 고정 디렉토리 'Processes'
        file_path = os.path.join(file_path, "Processes")

        # 업무명에 따른 3번 디렉토리
        business_col = '송신업무명' if is_send else '수신업무명'
        dir3 = safe_str(row.get(business_col, "Unknown")).strip()
        if not dir3:
            dir3 = "Common"
        file_path = os.path.join(file_path, dir3)

        # EMS명에 따른 4번 디렉토리
        dir4 = safe_str(row.get('EMS명', "Unknown")).strip()
        if not dir4:
            dir4 = "Common"
        file_path = os.path.join(file_path, dir4)

        # 5번 디렉토리 (패키지와 동일)
        file_path = os.path.join(file_path, dir2)

        # 파일명 생성
        group_id = safe_str(row.get('Group_ID', "")).strip()
        event_id = safe_str(row.get('Event_ID', "")).strip()
        if_name = safe_str(row.get('I/F명', "")).strip()

        # 파일명 조합
        file_name = f"{group_id}_{event_id}_{if_name}.ini"
        if not group_id:
            file_name = f"{event_id}_{if_name}.ini"
        if not event_id:
            file_name = f"{if_name}.ini"
        
        # 특수문자 제거
        file_name = re.sub(r'[<>:"/\\|?*]', '_', file_name)
        
        # 최종 파일 경로
        return os.path.join(file_path, file_name)
    
    except Exception as e:
        print(f"파일 경로 생성 중 오류: {str(e)}")
        return ""

def check_file_exists(file_path):
    """파일 경로가 실제로 존재하는지 확인"""
    if not file_path:
        return 0
    return 1 if os.path.exists(file_path) else 0

def count_files_in_directory(file_path):
    """파일이 위치한 디렉토리 내 파일 개수 반환"""
    if not file_path:
        return 0
    
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        return 0
    
    # 디렉토리 내 파일만 카운트 (하위 디렉토리 제외)
    return len([f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))])

def get_blue_fill_by_count(count):
    """파일 개수에 따라 적절한 파란색 배경 반환"""
    if count == 0:
        return GRAY_FILL
    elif count <= 3:
        return BLUE_FILLS['very_light']
    elif count <= 10:
        return BLUE_FILLS['light']
    elif count <= 20:
        return BLUE_FILLS['medium']
    else:
        return BLUE_FILLS['dark']

# ==================== 검증 함수 ====================

def check_systems(base_value, match_value, column_name):
    """송신시스템/수신시스템 비교 로직"""
    if not isinstance(base_value, str) or not isinstance(match_value, str):
        return f"{column_name} 비교오류 (비어있는 값)"
    
    expected_value = replace_ly_lz(base_value)
    if expected_value != match_value:
        return f"{column_name} 비교오류"
    return ""

def check_business_name(base_value, match_value, column_name):
    """업무명 비교 로직"""
    if not isinstance(base_value, str) or not isinstance(match_value, str):
        return f"{column_name} 비교오류 (비어있는 값)"
    
    if base_value == 'PNL_LY' and match_value != 'MES_LH':
        return f"{column_name} 비교오류"
    elif base_value == 'MOD_LZ' and match_value != 'MES_VO':
        return f"{column_name} 비교오류"
    return ""

def check_package(base_value, match_value, column_name):
    """패키지 비교 로직"""
    if not isinstance(base_value, str) or not isinstance(match_value, str):
        return f"{column_name} 비교오류 (비어있는 값)"
    
    if LY_PATTERN in base_value and LH_REPLACEMENT not in match_value:
        return f"{column_name} 비교오류"
    elif LZ_PATTERN in base_value and VO_REPLACEMENT not in match_value:
        return f"{column_name} 비교오류"
    return ""

def check_same_content(base_value, match_value, column_name):
    """내용이 같아야 하는 컬럼 비교"""
    if not isinstance(base_value, str) and not isinstance(match_value, str):
        # 둘 다 문자열이 아닌 경우 (예: NaN)
        if pd.isna(base_value) and pd.isna(match_value):
            return ""
        return f"{column_name} 비교오류 (유형 불일치)"
    
    if not isinstance(base_value, str) or not isinstance(match_value, str):
        return f"{column_name} 비교오류 (유형 불일치)"
    
    if base_value.strip() != match_value.strip():
        return f"{column_name} 비교오류"
    return ""

def check_table_or_routing(base_value, match_value, column_name):
    """테이블명 또는 라우팅 비교 로직"""
    if not isinstance(base_value, str) or not isinstance(match_value, str):
        if pd.isna(base_value) and pd.isna(match_value):
            return ""
        return f"{column_name} 비교오류 (비어있는 값)"
    
    if LY_PATTERN in base_value or LZ_PATTERN in base_value:
        expected_value = replace_ly_lz(base_value)
        if expected_value != match_value:
            return f"{column_name} 비교오류"
    elif base_value.strip() != match_value.strip():
        return f"{column_name} 비교오류"
    return ""

def check_table_with_split(base_value, match_value, column_name):
    """Source Table, Destination Table, Event_ID용 비교 로직 (단어 분할 후 LY/LZ 확인)"""
    if not isinstance(base_value, str) or not isinstance(match_value, str):
        if pd.isna(base_value) and pd.isna(match_value):
            return ""
        return f"{column_name} 비교오류 (비어있는 값)"
    
    # 기본값으로 단순 비교 
    should_check_ly_lz = False
    
    # '.'과 '_'로 분할하여 단어 확인
    words = re.split('[._]', base_value)
    for word in words:
        if word.startswith(LY_PATTERN) or word.startswith(LZ_PATTERN):
            should_check_ly_lz = True
            break
    
    if should_check_ly_lz:
        expected_value = replace_ly_lz(base_value)
        if expected_value != match_value:
            return f"{column_name} 비교오류"
    elif base_value.strip() != match_value.strip():
        return f"{column_name} 비교오류"
    
    return ""

def validate_row_pair(base_row, match_row, column_names):
    """기본행과 매칭행 쌍을 15가지 규칙으로 검증"""
    comparison_log = []
    
    # 1. 송신시스템 비교
    if '송신시스템' in column_names:
        result = check_systems(base_row['송신시스템'], match_row['송신시스템'], '1.송신시스템')
        if result:
            comparison_log.append(result)
    
    # 2. 수신시스템 비교
    if '수신시스템' in column_names:
        result = check_systems(base_row['수신시스템'], match_row['수신시스템'], '2.수신시스템')
        if result:
            comparison_log.append(result)
    
    # 3. I/F명 비교
    if 'I/F명' in column_names:
        result = check_same_content(base_row['I/F명'], match_row['I/F명'], '3.I/F명')
        if result:
            comparison_log.append(result)
    
    # 4. Event_ID 비교
    if 'Event_ID' in column_names:
        result = check_table_with_split(base_row['Event_ID'], match_row['Event_ID'], '4.Event_ID')
        if result:
            comparison_log.append(result)
    
    # 5. 수신업무명 비교
    if '수신업무명' in column_names:
        result = check_business_name(base_row['수신업무명'], match_row['수신업무명'], '5.수신업무명')
        if result:
            comparison_log.append(result)
    
    # 6. 송신업무명 비교
    if '송신업무명' in column_names:
        result = check_business_name(base_row['송신업무명'], match_row['송신업무명'], '6.송신업무명')
        if result:
            comparison_log.append(result)
    
    # 7. 송신패키지 비교
    if '송신패키지' in column_names:
        result = check_package(base_row['송신패키지'], match_row['송신패키지'], '7.송신패키지')
        if result:
            comparison_log.append(result)
    
    # 8. 수신패키지 비교
    if '수신패키지' in column_names:
        result = check_package(base_row['수신패키지'], match_row['수신패키지'], '8.수신패키지')
        if result:
            comparison_log.append(result)
    
    # 9. EMS명 비교
    if 'EMS명' in column_names:
        result = check_same_content(base_row['EMS명'], match_row['EMS명'], '9.EMS명')
        if result:
            comparison_log.append(result)
    
    # 10. Source Table 비교
    if 'Source Table' in column_names:
        result = check_table_with_split(base_row['Source Table'], match_row['Source Table'], '10.Source Table')
        if result:
            comparison_log.append(result)
    
    # 11. Destination Table 비교
    if 'Destination Table' in column_names:
        result = check_table_with_split(base_row['Destination Table'], match_row['Destination Table'], '11.Destination Table')
        if result:
            comparison_log.append(result)
    
    # 12. Routing 비교
    if 'Routing' in column_names:
        result = check_table_or_routing(base_row['Routing'], match_row['Routing'], '12.Routing')
        if result:
            comparison_log.append(result)
    
    # 13. 스케쥴 비교
    schedule_col = next((col for col in column_names if '스케쥴' in col), None)
    if schedule_col:
        result = check_same_content(base_row[schedule_col], match_row[schedule_col], '13.스케쥴')
        if result:
            comparison_log.append(result)
    
    # 14. 주기구분 비교
    if '주기구분' in column_names:
        result = check_same_content(base_row['주기구분'], match_row['주기구분'], '14.주기구분')
        if result:
            comparison_log.append(result)
    
    # 15. 주기 비교
    if '주기' in column_names:
        result = check_same_content(base_row['주기'], match_row['주기'], '15.주기')
        if result:
            comparison_log.append(result)
    
    # 비교로그 생성
    return ', '.join(comparison_log) if comparison_log else 'OK'

# ==================== 데이터베이스 처리 함수 ====================

def load_data_from_db(db_file, table_name):
    """SQLite 데이터베이스에서 데이터 로드"""
    try:
        # 데이터베이스 연결
        conn = sqlite3.connect(db_file)
        print(f"데이터베이스 '{db_file}'에 연결되었습니다.")
        
        # 테이블 존재 여부 확인
        cursor = conn.cursor()
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
        if not cursor.fetchone():
            print(f"오류: '{table_name}' 테이블이 데이터베이스에 존재하지 않습니다.")
            conn.close()
            return None
        
        # 전체 테이블 데이터 로드
        query = f"SELECT * FROM {table_name}"
        df_complete_table = pd.read_sql_query(query, conn)
        
        if len(df_complete_table) == 0:
            print(f"경고: '{table_name}' 테이블에 데이터가 없습니다.")
            conn.close()
            return None
        
        print(f"'{table_name}' 테이블에서 {len(df_complete_table)}개의 행을 로드했습니다.")
        
        # 연결 종료
        conn.close()
        
        return df_complete_table
    
    except sqlite3.Error as e:
        print(f"데이터베이스 오류: {str(e)}")
        return None
    
    except Exception as e:
        print(f"데이터 로드 중 오류 발생: {str(e)}")
        return None

def filter_data_with_ly_lz(df, target_columns):
    """'LY' 또는 'LZ'가 포함된 행 필터링"""
    if df is None or len(df) == 0:
        return None
    
    # 데이터프레임 복사본 생성
    df_copy = df.copy()
    
    # 필터링 조건 생성
    filter_condition = False
    for col in target_columns:
        if col in df_copy.columns:
            # 문자열 변환 및 NaN 처리
            df_copy[col] = df_copy[col].astype(str)
            filter_condition |= df_copy[col].str.contains(LY_PATTERN, na=False)
            filter_condition |= df_copy[col].str.contains(LZ_PATTERN, na=False)
    
    # 필터링 적용
    df_filtered = df_copy[filter_condition]
    
    if len(df_filtered) == 0:
        print(f"경고: 'LY' 또는 'LZ'가 포함된 행을 찾을 수 없습니다.")
        return None
    
    print(f"'LY' 또는 'LZ'가 포함된 {len(df_filtered)}개의 행을 찾았습니다.")
    return df_filtered

def find_matching_rows(df_filtered, df_complete, key_column):
    """매칭 행 찾기 및 우선순위 필터링"""
    if df_filtered is None or df_complete is None:
        return None
    
    # 필수 컬럼 확인
    required_columns = [key_column, '송신시스템', '수신시스템']
    for col in required_columns:
        if col not in df_filtered.columns or col not in df_complete.columns:
            print(f"오류: '{col}' 컬럼이 데이터프레임에 없습니다.")
            return None
    
    output_rows_info = []
    
    # 각 필터링된 행에 대해 처리
    for _, row in df_filtered.iterrows():
        # 원본 행 추가 (노란색 표시 없음)
        row_dict = row.to_dict()
        row_dict['highlight'] = False
        output_rows_info.append(row_dict)
        
        # key_column 값으로 매칭되는 행 찾기
        key_value = safe_str(row[key_column]).strip()
        if not key_value:
            continue
        
        # 매칭 조건 생성
        matching_rows = df_complete[df_complete[key_column].astype(str).str.strip() == key_value].copy()
        
        # 추가 필터링 조건 적용
        filtered_matches = []
        
        for _, match_row in matching_rows.iterrows():
            # 송신시스템 또는 수신시스템에서 'LY'/'LZ' 변환 후 일치 확인
            send_sys_match = False
            recv_sys_match = False
            
            # 송신시스템 확인
            if has_ly_lz(row['송신시스템']):
                expected_send_sys = replace_ly_lz(row['송신시스템'])
                if safe_str(match_row['송신시스템']) == expected_send_sys:
                    send_sys_match = True
            
            # 수신시스템 확인
            if has_ly_lz(row['수신시스템']):
                expected_recv_sys = replace_ly_lz(row['수신시스템'])
                if safe_str(match_row['수신시스템']) == expected_recv_sys:
                    recv_sys_match = True
            
            # 매칭 조건을 만족하면 추가
            if send_sys_match or recv_sys_match:
                match_dict = match_row.to_dict()
                match_dict['highlight'] = True
                match_dict['send_sys_match'] = send_sys_match
                match_dict['recv_sys_match'] = recv_sys_match
                filtered_matches.append(match_dict)
        
        # 매칭된 행이 없으면 다음 행으로
        if not filtered_matches:
            continue
        
        # 우선순위에 따라 필터링
        priority_matches = []
        
        # 케이스 1: 송신시스템과 수신시스템 모두 매칭되는 행
        case1_matches = [m for m in filtered_matches if m['send_sys_match'] and m['recv_sys_match']]
        if case1_matches:
            priority_matches = case1_matches
        else:
            # 케이스 2: 송신시스템만 매칭되는 행
            case2_matches = [m for m in filtered_matches if m['send_sys_match']]
            if case2_matches:
                priority_matches = case2_matches
            else:
                # 케이스 2-1: 수신시스템만 매칭되는 행
                case2_1_matches = [m for m in filtered_matches if m['recv_sys_match']]
                if case2_1_matches:
                    priority_matches = case2_1_matches
        
        # 디버그 모드에 따라 출력할 행 결정
        if debug_mode == 1:
            # 모든 매칭 행 출력 (노란색)
            for match in filtered_matches:
                output_rows_info.append(match)
        
        # 우선순위 매칭 행 출력 (연두색)
        for match in priority_matches:
            match['highlight'] = 'priority'  # 연두색 표시를 위한 플래그
            if debug_mode != 1:  # 디버그 모드가 아닌 경우에만 추가 (이미 추가되지 않았다면)
                output_rows_info.append(match)
    
    return output_rows_info 

# ==================== 엑셀 출력 함수 ====================

def create_excel_file(output_rows_info, df_columns, output_file):
    """결과를 엑셀 파일로 저장"""
    if not output_rows_info:
        print("오류: 출력할 데이터가 없습니다.")
        return None
    
    try:
        # 결과 DataFrame 생성
        df_excel_output = pd.DataFrame(output_rows_info)
        
        # 컬럼 순서 유지
        all_columns = [col for col in df_columns if col in df_excel_output.columns]
        
        # 출력에만 있는 컬럼 추가 (highlight 플래그 등)
        for col in df_excel_output.columns:
            if col not in all_columns and col not in ['highlight', 'send_sys_match', 'recv_sys_match']:
                all_columns.append(col)
        
        # 송신/수신 파일 경로 컬럼 추가
        df_excel_output['송신파일경로'] = df_excel_output.apply(lambda row: create_file_path(row, is_send=True), axis=1)
        df_excel_output['수신파일경로'] = df_excel_output.apply(lambda row: create_file_path(row, is_send=False), axis=1)
        all_columns.extend(['송신파일경로', '수신파일경로'])
        
        # 파일 존재 여부 컬럼 추가
        df_excel_output['송신파일존재'] = df_excel_output['송신파일경로'].apply(check_file_exists)
        df_excel_output['수신파일존재'] = df_excel_output['수신파일경로'].apply(check_file_exists)
        all_columns.extend(['송신파일존재', '수신파일존재'])
        
        # 디렉토리 파일 개수 컬럼 추가
        # 파일이 존재하는 경우에만 계산
        df_excel_output['송신DF'] = df_excel_output.apply(
            lambda row: count_files_in_directory(row['송신파일경로']) if row['송신파일존재'] == 1 else 0, 
            axis=1
        )
        df_excel_output['수신DF'] = df_excel_output.apply(
            lambda row: count_files_in_directory(row['수신파일경로']) if row['수신파일존재'] == 1 else 0, 
            axis=1
        )
        all_columns.extend(['송신DF', '수신DF'])
        
        # '비교로그' 컬럼 추가 (초기값은 빈 문자열)
        df_excel_output['비교로그'] = ''
        all_columns.append('비교로그')
        
        # 기본행과 매칭행 쌍에 대해 검증 수행
        for i in range(0, len(df_excel_output), 2):
            if i + 1 < len(df_excel_output):
                base_row = df_excel_output.iloc[i]
                match_row = df_excel_output.iloc[i + 1]
                
                # 연두색 매칭행인 경우에만 검증 수행
                if match_row.get('highlight') == 'priority':
                    # 검증 결과
                    validation_result = validate_row_pair(base_row, match_row, all_columns)
                    
                    # 결과 기록 (기본행과 매칭행 모두에게)
                    df_excel_output.at[i, '비교로그'] = validation_result
                    df_excel_output.at[i + 1, '비교로그'] = validation_result
        
        # 컬럼 순서 적용 (highlight, send_sys_match, recv_sys_match 제외)
        output_columns = [col for col in all_columns if col not in ['highlight', 'send_sys_match', 'recv_sys_match']]
        df_excel_output = df_excel_output[output_columns]
        
        # 행 인덱스 목록 (색상 적용용)
        yellow_rows = []
        green_rows = []
        
        for i, row in df_excel_output.iterrows():
            highlight = output_rows_info[i].get('highlight', False)
            if highlight == 'priority':
                green_rows.append(i)
            elif highlight:
                yellow_rows.append(i)
        
        # Excel 파일 저장
        print(f"결과를 '{output_file}'로 저장 중...")
        
        # xlsxwriter 엔진으로 Excel 파일 생성
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        df_excel_output.to_excel(writer, sheet_name='ProcessedData', index=False)
        
        # 워크북과 워크시트 객체 가져오기
        workbook = writer.book
        worksheet = writer.sheets['ProcessedData']
        
        # 셀 서식 설정
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})  # 노란색
        green_format = workbook.add_format({'bg_color': '#CCFFCC'})   # 연두색
        
        # 행 색상 적용
        for row_idx in yellow_rows:
            worksheet.set_row(row_idx + 1, None, yellow_format)  # +1 for header row
        
        for row_idx in green_rows:
            worksheet.set_row(row_idx + 1, None, green_format)  # +1 for header row
        
        # 열 너비 자동 조절
        for i, col in enumerate(df_excel_output.columns):
            max_len = max(
                df_excel_output[col].astype(str).map(len).max(),
                len(str(col))
            ) + 2  # 여유 공간
            worksheet.set_column(i, i, max_len)
        
        # 파일 저장
        writer.close()
        
        print(f"Excel 파일이 '{output_file}'에 저장되었습니다.")
        
        # 추가 서식 적용을 위해 openpyxl로 파일 다시 열기
        apply_additional_formatting(output_file)
        
        return output_file
    
    except Exception as e:
        print(f"Excel 파일 생성 중 오류 발생: {str(e)}")
        return None

def apply_additional_formatting(file_path):
    """추가 서식 적용 (파일 존재 여부 및 디렉토리 파일 개수에 따른 색상)"""
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        
        # 헤더 행 찾기
        header_row = 1
        
        # 컬럼 인덱스 찾기
        col_indices = {}
        for col_idx, cell in enumerate(ws[header_row], 1):
            col_indices[cell.value] = col_idx
        
        # 파일 존재 여부와 디렉토리 파일 개수 컬럼이 있는지 확인
        required_cols = ['송신파일존재', '수신파일존재', '송신DF', '수신DF', '비교로그']
        for col in required_cols:
            if col not in col_indices:
                print(f"경고: '{col}' 컬럼을 찾을 수 없습니다.")
        
        # 각 행에 대해 처리
        for row_idx in range(header_row + 1, ws.max_row + 1):
            # 파일 존재 여부에 따라 색상 적용
            for col_name in ['송신파일존재', '수신파일존재']:
                if col_name in col_indices:
                    col_idx = col_indices[col_name]
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value == 1:
                        cell.fill = GREEN_FILL
                    elif cell.value == 0:
                        cell.fill = ORANGE_FILL
            
            # 디렉토리 파일 개수에 따라 색상 적용
            for col_name in ['송신DF', '수신DF']:
                if col_name in col_indices:
                    col_idx = col_indices[col_name]
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value is not None:
                        count = int(cell.value)
                        cell.fill = get_blue_fill_by_count(count)
            
            # 비교로그에 따라 색상 적용
            if '비교로그' in col_indices:
                col_idx = col_indices['비교로그']
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value and cell.value != 'OK':
                    cell.fill = ORANGE_FILL
        
        wb.save(file_path)
        print(f"추가 서식이 '{file_path}'에 적용되었습니다.")
    
    except Exception as e:
        print(f"추가 서식 적용 중 오류 발생: {str(e)}") 

# ==================== 메인 함수 ====================

def process_data(db_file, table_name, target_columns, key_column, output_file=None):
    """데이터 처리 메인 함수"""
    try:
        print("=" * 50)
        print("iflist03a.py - 통합 데이터 처리 및 검증 프로그램 실행")
        print("=" * 50)
        print(f"처리 모드: {'디버그 모드' if debug_mode == 1 else '일반 모드'}")
        
        # 1. 데이터베이스에서 데이터 로드
        df_complete_table = load_data_from_db(db_file, table_name)
        if df_complete_table is None:
            return False
        
        # 2. 'LY' 또는 'LZ'가 포함된 행 필터링
        df_filtered = filter_data_with_ly_lz(df_complete_table, target_columns)
        if df_filtered is None:
            return False
        
        # 3. 매칭 행 찾기 및 우선순위 필터링
        output_rows_info = find_matching_rows(df_filtered, df_complete_table, key_column)
        if not output_rows_info:
            print("오류: 매칭 행을 찾을 수 없습니다.")
            return False
        
        # 4. 결과를 엑셀 파일로 저장
        if output_file is None:
            output_file = output_file_name
        
        final_output_file = create_excel_file(output_rows_info, df_complete_table.columns, output_file)
        if not final_output_file:
            return False
        
        # 5. 결과 요약 출력
        print("\n" + "=" * 50)
        print("처리 결과 요약")
        print("=" * 50)
        print(f"- 원본 데이터 수: {len(df_complete_table)}개 행")
        print(f"- 'LY'/'LZ' 포함 행: {len(df_filtered)}개 행")
        print(f"- 최종 출력 행: {len(output_rows_info)}개 행")
        print(f"- 출력 파일: {final_output_file}")
        
        # 디렉토리 파일 개수 통계
        send_file_counts = [row.get('송신DF', 0) for row in output_rows_info if isinstance(row.get('송신DF', 0), (int, float))]
        recv_file_counts = [row.get('수신DF', 0) for row in output_rows_info if isinstance(row.get('수신DF', 0), (int, float))]
        
        if send_file_counts:
            total_send_files = sum(send_file_counts)
            avg_send_files = total_send_files / len(send_file_counts) if send_file_counts else 0
            print(f"- 송신 디렉토리 총 파일 수: {total_send_files}개 (디렉토리당 평균: {avg_send_files:.2f}개)")
        
        if recv_file_counts:
            total_recv_files = sum(recv_file_counts)
            avg_recv_files = total_recv_files / len(recv_file_counts) if recv_file_counts else 0
            print(f"- 수신 디렉토리 총 파일 수: {total_recv_files}개 (디렉토리당 평균: {avg_recv_files:.2f}개)")
        
        print("=" * 50)
        print("처리가 완료되었습니다.")
        
        return True
    
    except Exception as e:
        print(f"처리 중 오류 발생: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return False

def main():
    """메인 함수"""
    # 기본 설정값
    db_file = 'info.sqlite'
    table_name = 'list'
    target_columns = ['송신시스템', '수신시스템']
    key_column = 'Event_ID'
    
    # 명령행 인수 처리
    if len(sys.argv) > 1:
        db_file = sys.argv[1]
    
    if len(sys.argv) > 2:
        table_name = sys.argv[2]
    
    # 데이터 처리
    process_data(db_file, table_name, target_columns, key_column)

if __name__ == "__main__":
    main() 