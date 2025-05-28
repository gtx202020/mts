"""
인터페이스 정보 엑셀 파일 리더 및 BW 수신파일 파서 모듈

이 모듈은 다음과 같은 기능을 제공합니다:
1. 특정 형식의 엑셀 파일에서 인터페이스 정보를 읽어 파이썬 자료구조로 변환
2. TIBCO BW .process 파일에서 수신용 INSERT 쿼리를 추출하고 파라미터 매핑 처리

주요 클래스:
- InterfaceExcelReader: 엑셀 파일에서 인터페이스 정보 추출
- BWProcessFileParser: BW .process 파일에서 INSERT 쿼리 추출
- ProcessFileMapper: 일련번호와 string_replacer용 엑셀을 매핑하는 클래스

주요 함수:
- parse_bw_receive_file: BW 수신파일 파싱을 위한 편의 함수
"""

import os
import ast
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd


class InterfaceExcelReader:
    """
    인터페이스 정보가 담긴 엑셀 파일을 읽어 파이썬 자료구조로 변환하는 클래스
    
    엑셀 파일 구조:
    - B열부터 3컬럼 단위로 하나의 인터페이스 블록
    - 1행: 인터페이스명
    - 2행: 인터페이스ID  
    - 3행: DB 연결 정보 (문자열로 저장된 딕셔너리)
    - 4행: 테이블 정보 (문자열로 저장된 딕셔너리)
    - 5행부터: 컬럼 매핑 정보
    """
    
    def __init__(self, replacer_excel_path: str = None):
        """
        InterfaceExcelReader 클래스 초기화
        
        Args:
            replacer_excel_path (str, optional): string_replacer용 엑셀 파일 경로
        """
        self.processed_count = 0
        self.error_count = 0
        self.last_error_messages = []
        
        # ProcessFileMapper 초기화
        self.process_mapper = None
        if replacer_excel_path:
            self.process_mapper = ProcessFileMapper(replacer_excel_path)
    
    def load_interfaces(self, excel_path: str) -> List[Dict[str, Any]]:
        """
        엑셀 파일에서 모든 인터페이스 정보를 읽어 리스트로 반환
        
        Args:
            excel_path (str): 읽을 엑셀 파일의 경로
            
        Returns:
            List[Dict[str, Any]]: 인터페이스 정보 딕셔너리들의 리스트
            
        Raises:
            FileNotFoundError: 엑셀 파일이 존재하지 않는 경우
            PermissionError: 파일 접근 권한이 없는 경우
            ValueError: 엑셀 파일 형식이 올바르지 않은 경우
        """
        # 초기화
        self.processed_count = 0
        self.error_count = 0
        self.last_error_messages = []
        
        # 파일 존재 여부 확인
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")
        
        interfaces = []
        workbook = None
        
        try:
            # 엑셀 파일 열기
            workbook = openpyxl.load_workbook(excel_path, read_only=True)
            worksheet = workbook.active
            
            if worksheet is None:
                raise ValueError("활성 워크시트를 찾을 수 없습니다")
            
            # [디버깅용] 첫 번째 인터페이스 블록만 처리 (B열부터 시작)
            current_col = 2  # B열 = 2
            
            try:
                print(f"=== 디버깅 모드: 첫 번째 인터페이스만 처리 (컬럼 {current_col}) ===")
                
                # 첫 번째 인터페이스 블록 읽기
                interface_data = self._read_interface_block(worksheet, current_col)
                
                if interface_data is not None:
                    interfaces.append(interface_data)
                    self.processed_count += 1
                    print(f"첫 번째 인터페이스 처리 완료: {interface_data.get('interface_name', 'Unknown')}")
                else:
                    print("첫 번째 인터페이스 블록이 비어있습니다.")
                
            except Exception as e:
                self.error_count += 1
                error_msg = f"첫 번째 인터페이스 블록(컬럼 {current_col})에서 오류 발생: {str(e)}"
                self.last_error_messages.append(error_msg)
                print(f"Warning: {error_msg}")
            
            # 원래 루프 코드는 주석 처리 (디버깅 후 복원용)
            """
            # B열부터 시작하여 3컬럼 단위로 처리
            current_col = 2  # B열 = 2
            
            while current_col <= worksheet.max_column:
                try:
                    # 인터페이스 블록 읽기
                    interface_data = self._read_interface_block(worksheet, current_col)
                    
                    if interface_data is None:
                        # 빈 인터페이스 발견시 종료
                        break
                    
                    interfaces.append(interface_data)
                    self.processed_count += 1
                    
                except Exception as e:
                    self.error_count += 1
                    error_msg = f"컬럼 {current_col}에서 오류 발생: {str(e)}"
                    self.last_error_messages.append(error_msg)
                    print(f"Warning: {error_msg}")
                
                # 다음 인터페이스 블록으로 이동 (3컬럼씩)
                current_col += 3
            """
                
        except Exception as e:
            raise ValueError(f"엑셀 파일 처리 중 오류 발생: {str(e)}")
        
        finally:
            # 리소스 정리
            if workbook:
                workbook.close()
        
        return interfaces
    
    def _read_interface_block(self, worksheet: Worksheet, start_col: int) -> Optional[Dict[str, Any]]:
        """
        단일 인터페이스 블록(3컬럼)에서 정보를 읽어 딕셔너리로 반환
        
        Args:
            worksheet: 엑셀 워크시트 객체
            start_col (int): 인터페이스 블록의 시작 컬럼 번호
            
        Returns:
            Optional[Dict[str, Any]]: 인터페이스 정보 딕셔너리, 빈 블록이면 None
        """
        # 기본 구조 생성
        interface_info = {
            'interface_name': '',
            'interface_id': '',
            'serial_number': '',
            'send_original': '',        # 송신 원본파일 경로
            'send_copy': '',            # 송신 복사파일 경로  
            'recv_original': '',        # 수신 원본파일 경로
            'recv_copy': '',            # 수신 복사파일 경로
            'send_schema': '',          # 송신 스키마파일
            'recv_schema': '',          # 수신 스키마파일
            'send': {
                'owner': None,
                'table_name': None,
                'columns': [],
                'db_info': {}
            },
            'recv': {
                'owner': None,
                'table_name': None,
                'columns': [],
                'db_info': {}
            }
        }
        
        # 1단계: 필수 정보만 먼저 체크 (interface_name, interface_id)
        try:
            # 1행: 인터페이스명 읽기
            interface_name_cell = worksheet.cell(row=1, column=start_col)
            interface_info['interface_name'] = interface_name_cell.value or ''
            
            # 1행: 일련번호 읽기 (인터페이스명 오른쪽으로 두 칸)
            serial_number_cell = worksheet.cell(row=1, column=start_col + 2)
            interface_info['serial_number'] = serial_number_cell.value or ''
            
            # 2행: 인터페이스ID 읽기 (필수값)
            interface_id_cell = worksheet.cell(row=2, column=start_col)
            interface_id = interface_id_cell.value
            
            if not interface_id:
                # 인터페이스 ID가 없으면 빈 블록으로 간주
                return None
            
            interface_info['interface_id'] = str(interface_id).strip()
            
        except Exception as e:
            print(f"Warning: 필수 정보 읽기 실패 (컬럼 {start_col}): {str(e)}")
            return None
        
        # 2단계: 선택적 정보들을 개별적으로 안전하게 처리
        # DB 연결 정보 읽기 (실패해도 계속)
        try:
            send_db_cell = worksheet.cell(row=3, column=start_col)
            recv_db_cell = worksheet.cell(row=3, column=start_col + 1)
            
            interface_info['send']['db_info'] = self._parse_cell_dict(send_db_cell.value)
            interface_info['recv']['db_info'] = self._parse_cell_dict(recv_db_cell.value)
            
        except Exception as e:
            print(f"Warning: DB 정보 읽기 실패 (컬럼 {start_col}): {str(e)}")
            # DB 정보 읽기 실패해도 빈 딕셔너리로 계속 진행
        
        # 테이블 정보 읽기 (실패해도 계속)
        try:
            # 송신 테이블 정보 읽기 (row=4, column=start_col)
            send_table_cell = worksheet.cell(row=4, column=start_col)
            send_table_dict = self._parse_cell_dict(send_table_cell.value)
            if send_table_dict:
                interface_info['send']['owner'] = send_table_dict.get('owner')
                interface_info['send']['table_name'] = send_table_dict.get('table_name')
            
            # 수신 테이블 정보 읽기 (row=4, column=start_col+1)
            recv_table_cell = worksheet.cell(row=4, column=start_col + 1)
            recv_table_dict = self._parse_cell_dict(recv_table_cell.value)
            if recv_table_dict:
                interface_info['recv']['owner'] = recv_table_dict.get('owner')
                interface_info['recv']['table_name'] = recv_table_dict.get('table_name')
            
        except Exception as e:
            print(f"Warning: 테이블 정보 읽기 실패 (컬럼 {start_col}): {str(e)}")
        
        # 컬럼 매핑 정보 읽기 (실패해도 계속)
        try:
            send_columns, recv_columns = self._read_column_mappings(worksheet, start_col, 8)
            interface_info['send']['columns'] = send_columns
            interface_info['recv']['columns'] = recv_columns
            
        except Exception as e:
            print(f"Warning: 컬럼 매핑 읽기 실패 (컬럼 {start_col}): {str(e)}")
            # 컬럼 매핑 읽기 실패해도 빈 리스트로 계속 진행
        
        # 3단계: ProcessFileMapper로 .process 파일 정보 추가
        if self.process_mapper and interface_info['serial_number']:
            try:
                process_files = self.process_mapper.get_process_files_by_serial(interface_info['serial_number'])
                if process_files:
                    interface_info.update(process_files)
                    print(f"Info: 일련번호 {interface_info['serial_number']}의 process 파일 정보 추가됨")
            except Exception as e:
                print(f"Warning: Process 파일 정보 가져오기 실패: {str(e)}")
        
        return interface_info
    
    def _parse_cell_dict(self, cell_value: Any) -> Dict[str, Any]:
        """
        셀 값을 딕셔너리로 안전하게 파싱
        
        Args:
            cell_value: 엑셀 셀의 값
            
        Returns:
            Dict[str, Any]: 파싱된 딕셔너리, 실패시 빈 딕셔너리
        """
        if cell_value is None:
            return {}
        
        try:
            # 문자열을 딕셔너리로 안전하게 변환
            if isinstance(cell_value, str) and cell_value.strip():
                return ast.literal_eval(cell_value.strip())
            else:
                return {}
        except (SyntaxError, ValueError, TypeError):
            # 파싱 실패시 빈 딕셔너리 반환
            return {}
    
    def _read_column_mappings(self, worksheet: Worksheet, start_col: int, start_row: int) -> tuple[List[str], List[str]]:
        """
        컬럼 매핑 정보를 읽어 송신/수신 컬럼 리스트로 반환
        
        Args:
            worksheet: 엑셀 워크시트 객체
            start_col (int): 시작 컬럼 번호
            start_row (int): 시작 행 번호
            
        Returns:
            tuple[List[str], List[str]]: (송신 컬럼 리스트, 수신 컬럼 리스트)
        """
        send_columns = []
        recv_columns = []
        
        current_row = start_row
        
        # 빈 행이 나올 때까지 계속 읽기
        while current_row <= worksheet.max_row:
            send_cell = worksheet.cell(row=current_row, column=start_col)
            recv_cell = worksheet.cell(row=current_row, column=start_col + 1)
            
            send_value = send_cell.value
            recv_value = recv_cell.value
            
            # 둘 다 비어있으면 종료
            if not send_value and not recv_value:
                break
            
            # 값이 있으면 문자열로 변환하여 추가
            send_columns.append(str(send_value) if send_value else '')
            recv_columns.append(str(recv_value) if recv_value else '')
            
            current_row += 1
        
        return send_columns, recv_columns
    
    def get_statistics(self) -> Dict[str, int]:
        """
        마지막 처리 결과의 통계 정보 반환
        
        Returns:
            Dict[str, int]: 처리 통계 정보
        """
        return {
            'processed_count': self.processed_count,
            'error_count': self.error_count,
            'total_attempts': self.processed_count + self.error_count
        }
    
    def get_last_errors(self) -> List[str]:
        """
        마지막 처리에서 발생한 오류 메시지들 반환
        
        Returns:
            List[str]: 오류 메시지 리스트
        """
        return self.last_error_messages.copy()


class BWProcessFileParser:
    """
    TIBCO BW .process 파일에서 수신용 INSERT 쿼리를 추출하는 클래스
    """
    
    def __init__(self):
        """BWProcessFileParser 초기화"""
        self.ns = {
            'pd': 'http://xmlns.tibco.com/bw/process/2003',
            'xsl': 'http://www.w3.org/1999/XSL/Transform',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
        }
        self.processed_count = 0
        self.error_count = 0
        self.last_error_messages = []
    
    def parse_bw_process_file(self, process_file_path: str) -> List[str]:
        """
        BW .process 파일에서 수신용 INSERT 쿼리를 추출
        
        Args:
            process_file_path (str): .process 파일의 경로
            
        Returns:
            List[str]: 추출된 INSERT 쿼리 목록
            
        Raises:
            FileNotFoundError: 파일이 존재하지 않는 경우
            ValueError: 파일 형식이 올바르지 않은 경우
        """
        # 초기화
        self.processed_count = 0
        self.error_count = 0
        self.last_error_messages = []
        
        # 파일 존재 여부 확인
        if not os.path.exists(process_file_path):
            raise FileNotFoundError(f"BW process 파일을 찾을 수 없습니다: {process_file_path}")
        
        insert_queries = []
        
        try:
            # XML 파일 파싱
            tree = ET.parse(process_file_path)
            root = tree.getroot()
            
            print(f"\n=== BW Process 파일 처리 시작: {process_file_path} ===")
            
            # JDBC 액티비티 찾기
            activities = root.findall('.//pd:activity', self.ns)
            
            for activity in activities:
                try:
                    # JDBC 액티비티 타입 확인
                    activity_type = activity.find('./pd:type', self.ns)
                    if activity_type is None or 'jdbc' not in activity_type.text.lower():
                        continue
                    
                    activity_name = activity.get('name', 'Unknown')
                    print(f"\nJDBC 액티비티 발견: {activity_name}")
                    
                    # statement 추출
                    statement = activity.find('.//config/statement')
                    if statement is not None and statement.text:
                        query = statement.text.strip()
                        print(f"\n발견된 쿼리:\n{query}")
                        
                        # INSERT 쿼리인 경우만 처리
                        if query.lower().startswith('insert'):
                            # Oracle 힌트 제거
                            cleaned_query = self._remove_oracle_hints(query)
                            
                            # 파라미터 매핑 처리
                            mapped_query = self._process_query_parameters(activity, cleaned_query)
                            
                            insert_queries.append(mapped_query)
                            self.processed_count += 1
                            print(f"=> 최종 처리된 INSERT 쿼리:\n{mapped_query}")
                        else:
                            print(f"=> INSERT 쿼리가 아니므로 제외: {query[:50]}...")
                
                except Exception as e:
                    self.error_count += 1
                    error_msg = f"액티비티 '{activity.get('name', 'Unknown')}' 처리 중 오류: {str(e)}"
                    self.last_error_messages.append(error_msg)
                    print(f"Warning: {error_msg}")
            
            print(f"\n=== 처리된 INSERT 쿼리 수: {len(insert_queries)} ===")
            
        except ET.ParseError as e:
            raise ValueError(f"XML 파싱 오류: {str(e)}")
        except Exception as e:
            raise ValueError(f"BW process 파일 처리 중 오류 발생: {str(e)}")
        
        return insert_queries
    
    def _remove_oracle_hints(self, query: str) -> str:
        """
        SQL 쿼리에서 Oracle 힌트(/*+ ... */) 제거
        
        Args:
            query (str): 원본 SQL 쿼리
            
        Returns:
            str: 힌트가 제거된 SQL 쿼리
        """
        # /*+ ... */ 패턴의 힌트 제거
        cleaned_query = re.sub(r'/\*\+[^*]*\*/', '', query)
        # 불필요한 공백 정리 (여러 개의 공백을 하나로)
        cleaned_query = re.sub(r'\s+', ' ', cleaned_query).strip()
        
        if cleaned_query != query:
            print("\n=== Oracle 힌트 제거 ===")
            print(f"원본 쿼리: {query}")
            print(f"정리된 쿼리: {cleaned_query}")
        
        return cleaned_query
    
    def _process_query_parameters(self, activity, query: str) -> str:
        """
        쿼리의 파라미터를 실제 값으로 매핑
        
        Args:
            activity: JDBC 액티비티 XML 요소
            query (str): SQL 쿼리
            
        Returns:
            str: 파라미터가 매핑된 SQL 쿼리
        """
        try:
            # 1단계: prepared_Param_DataType의 파라미터 이름으로 매핑
            param_names = self._get_parameter_names(activity)
            stage1_query = self._replace_with_param_names(query, param_names)
            
            # 2단계: Record의 실제 값으로 매핑
            mappings = self._get_record_mappings(activity, param_names)
            stage2_query = self._replace_with_actual_values(stage1_query, mappings)
            
            return stage2_query
            
        except Exception as e:
            print(f"파라미터 매핑 중 오류 발생: {str(e)}")
            return query  # 오류 발생시 원본 쿼리 반환
    
    def _get_parameter_names(self, activity) -> List[str]:
        """
        Prepared_Param_DataType에서 파라미터 이름 목록 추출
        
        Args:
            activity: JDBC 액티비티 XML 요소
            
        Returns:
            List[str]: 파라미터 이름 목록
        """
        param_names = []
        
        prepared_params = activity.find('.//Prepared_Param_DataType', self.ns)
        if prepared_params is not None:
            for param in prepared_params.findall('./parameter', self.ns):
                param_name = param.find('./parameterName', self.ns)
                if param_name is not None and param_name.text:
                    name = param_name.text.strip()
                    param_names.append(name)
                    print(f"파라미터 이름 추출: {name}")
        
        return param_names
    
    def _replace_with_param_names(self, query: str, param_names: List[str]) -> str:
        """
        SQL 쿼리의 ? 플레이스홀더를 파라미터 이름으로 대체
        
        Args:
            query (str): 원본 SQL 쿼리
            param_names (List[str]): 파라미터 이름 목록
            
        Returns:
            str: 파라미터 이름이 대체된 SQL 쿼리
        """
        parts = query.split('?')
        if len(parts) == 1:  # 플레이스홀더가 없는 경우
            return query
        
        result = parts[0]
        for i, param_name in enumerate(param_names):
            if i < len(parts) - 1:
                result += f":{param_name}" + parts[i+1]
        
        print(f"\n1단계 매핑 결과: {result}")
        return result
    
    def _get_record_mappings(self, activity, param_names: List[str]) -> Dict[str, str]:
        """
        Record 태그에서 실제 값 매핑 정보 추출
        
        Args:
            activity: JDBC 액티비티 XML 요소
            param_names: 파라미터 이름 목록
            
        Returns:
            Dict[str, str]: 파라미터 이름과 매핑된 실제 값의 딕셔너리
        """
        mappings = {}
        
        input_bindings = activity.find('.//pd:inputBindings', self.ns)
        if input_bindings is None:
            return mappings
        
        # jdbcUpdateActivityInput/Record 찾기
        jdbc_input = input_bindings.find('.//jdbcUpdateActivityInput')
        if jdbc_input is None:
            return mappings
        
        # for-each/Record 찾기
        for_each = jdbc_input.find('.//xsl:for-each', self.ns)
        record = for_each.find('./Record') if for_each is not None else jdbc_input
        
        if record is not None:
            # 각 파라미터 이름에 대해 매핑 찾기
            for param_name in param_names:
                param_element = record.find(f'.//{param_name}')
                if param_element is not None:
                    # value-of 체크
                    value_of = param_element.find('.//xsl:value-of', self.ns)
                    if value_of is not None:
                        select_attr = value_of.get('select', '')
                        if select_attr:
                            value = select_attr.split('/')[-1]
                            mappings[param_name] = value
                            print(f"매핑 발견: {param_name} -> {value}")
                    
                    # choose/when 체크
                    else:
                        choose = param_element.find('.//xsl:choose', self.ns)
                        if choose is not None:
                            when = choose.find('.//xsl:when', self.ns)
                            if when is not None:
                                test_attr = when.get('test', '')
                                if 'exists(' in test_attr:
                                    value = test_attr[test_attr.find('(')+1:test_attr.find(')')]
                                    mappings[param_name] = value
                                    print(f"매핑 발견: {param_name} -> {value}")
        
        return mappings
    
    def _replace_with_actual_values(self, query: str, mappings: Dict[str, str]) -> str:
        """
        파라미터 이름을 실제 값으로 대체
        
        Args:
            query (str): 파라미터 이름이 대체된 쿼리
            mappings (Dict[str, str]): 파라미터 이름과 실제 값의 매핑
            
        Returns:
            str: 실제 값이 대체된 SQL 쿼리
        """
        result = query
        
        for param_name, actual_value in mappings.items():
            # 정확한 파라미터 이름만 대체
            result = re.sub(f":{param_name}\\b", f":{actual_value}", result)
        
        print(f"\n2단계 매핑 결과: {result}")
        return result
    
    def get_statistics(self) -> Dict[str, int]:
        """
        마지막 처리 결과의 통계 정보 반환
        
        Returns:
            Dict[str, int]: 처리 통계 정보
        """
        return {
            'processed_count': self.processed_count,
            'error_count': self.error_count,
            'total_attempts': self.processed_count + self.error_count
        }
    
    def get_last_errors(self) -> List[str]:
        """
        마지막 처리에서 발생한 오류 메시지들 반환
        
        Returns:
            List[str]: 오류 메시지 리스트
        """
        return self.last_error_messages.copy()


class ProcessFileMapper:
    """
    test_iflist.py의 일련번호와 string_replacer용 엑셀을 매핑하는 클래스
    """
    
    def __init__(self, replacer_excel_path: str):
        """ProcessFileMapper 초기화
        
        Args:
            replacer_excel_path (str): string_replacer.py에서 사용하는 엑셀 파일 경로
        """
        self.replacer_excel_path = replacer_excel_path
        self.df = None
        if os.path.exists(replacer_excel_path):
            try:
                self.df = pd.read_excel(replacer_excel_path, engine='openpyxl')
            except Exception as e:
                print(f"Warning: ProcessFileMapper - 엑셀 파일 로드 실패: {str(e)}")
    
    def get_process_files_by_serial(self, serial_number: str) -> Dict[str, str]:
        """
        일련번호(serial_number)로 .process 파일 경로들을 가져옴
        
        Args:
            serial_number (str): 인터페이스 일련번호
            
        Returns:
            Dict[str, str]: 프로세스 파일 정보
            {
                'send_original': '송신 원본파일 경로',
                'send_copy': '송신 복사파일 경로', 
                'recv_original': '수신 원본파일 경로',
                'recv_copy': '수신 복사파일 경로',
                'send_schema': '송신 스키마파일',
                'recv_schema': '수신 스키마파일'
            }
        """
        if self.df is None or not serial_number:
            return {}
        
        try:
            # N번째 행 = serial_number 매핑 (1-based to 0-based)
            row_index = int(serial_number) - 1
            
            if row_index * 2 + 1 >= len(self.df):
                return {}
            
            normal_row = self.df.iloc[row_index * 2]     # 기본행
            match_row = self.df.iloc[row_index * 2 + 1]  # 매칭행
            
            result = {}
            
            # 송신 파일 생성 여부 확인
            if (pd.notna(normal_row.get('송신파일생성여부')) and 
                float(normal_row['송신파일생성여부']) == 1.0):
                result['send_original'] = str(match_row.get('송신파일경로', ''))
                result['send_copy'] = str(normal_row.get('송신파일경로', ''))
            
            # 수신 파일 생성 여부 확인  
            if (pd.notna(normal_row.get('수신파일생성여부')) and 
                float(normal_row['수신파일생성여부']) == 1.0):
                result['recv_original'] = str(match_row.get('수신파일경로', ''))
                result['recv_copy'] = str(normal_row.get('수신파일경로', ''))
            
            # 송신 스키마 파일 생성 여부 확인
            if (pd.notna(normal_row.get('송신스키마파일생성여부')) and 
                float(normal_row['송신스키마파일생성여부']) == 1.0):
                result['send_schema'] = str(normal_row.get('송신스키마파일명', ''))
            
            # 수신 스키마 파일 생성 여부 확인
            if (pd.notna(normal_row.get('수신스키마파일생성여부')) and 
                float(normal_row['수신스키마파일생성여부']) == 1.0):
                result['recv_schema'] = str(normal_row.get('수신스키마파일명', ''))
            
            return result
            
        except Exception as e:
            print(f"Warning: ProcessFileMapper - 일련번호 {serial_number} 처리 실패: {str(e)}")
            return {}


def parse_bw_receive_file(process_file_path: str) -> List[str]:
    """
    BW의 수신파일(.process)을 파싱하여 INSERT 쿼리를 추출하는 함수
    
    Args:
        process_file_path (str): BW .process 파일의 경로
        
    Returns:
        List[str]: 추출된 INSERT 쿼리 목록
        
    Raises:
        FileNotFoundError: 파일이 존재하지 않는 경우
        ValueError: 파일 형식이 올바르지 않은 경우
    """
    parser = BWProcessFileParser()
    return parser.parse_bw_process_file(process_file_path)


# 사용 예시 및 테스트
if __name__ == "__main__":
    # 테스트용 샘플 코드
    def test_interface_reader():
        """InterfaceExcelReader 테스트 함수"""
        # replacer_excel_path는 string_replacer.py에서 사용하는 엑셀 파일 경로
        replacer_excel_path = "replacer_input.xlsx"  # 실제 환경에 맞게 수정 필요
        reader = InterfaceExcelReader(replacer_excel_path)
        
        # 테스트할 엑셀 파일 경로 (실제 환경에 맞게 수정 필요)
        test_excel_path = "input.xlsx"
        
        try:
            print("=== 인터페이스 엑셀 리더 테스트 시작 ===")
            print(f"파일 경로: {test_excel_path}")
            
            # 파일 존재 여부 확인
            if not os.path.exists(test_excel_path):
                print(f"테스트 파일이 없습니다: {test_excel_path}")
                print("테스트를 위해 실제 엑셀 파일 경로를 지정해주세요.")
                return
            
            # 인터페이스 정보 로드
            interfaces = reader.load_interfaces(test_excel_path)
            
            # 결과 출력
            print(f"\n=== 처리 결과 ===")
            print(f"총 {len(interfaces)}개의 인터페이스를 읽었습니다.")
            
            # 통계 정보 출력
            stats = reader.get_statistics()
            print(f"처리 성공: {stats['processed_count']}개")
            print(f"처리 실패: {stats['error_count']}개")
            
            # 오류가 있었다면 출력
            errors = reader.get_last_errors()
            if errors:
                print(f"\n=== 발생한 오류들 ===")
                for error in errors:
                    print(f"- {error}")
            
            # 첫 번째 인터페이스 정보 샘플 출력
            if interfaces:
                print(f"\n=== 첫 번째 인터페이스 샘플 ===")
                first_interface = interfaces[0]
                print(f"인터페이스명: {first_interface['interface_name']}")
                print(f"인터페이스ID: {first_interface['interface_id']}")
                print(f"일련번호: {first_interface['serial_number']}")
                print(f"송신 테이블: {first_interface['send']['table_name']}")
                print(f"수신 테이블: {first_interface['recv']['table_name']}")
                print(f"송신 컬럼 수: {len(first_interface['send']['columns'])}")
                print(f"수신 컬럼 수: {len(first_interface['recv']['columns'])}")
                print(f"송신 원본파일: {first_interface.get('send_original', 'N/A')}")
                print(f"송신 복사파일: {first_interface.get('send_copy', 'N/A')}")
                print(f"수신 원본파일: {first_interface.get('recv_original', 'N/A')}")
                print(f"수신 복사파일: {first_interface.get('recv_copy', 'N/A')}")
                print(f"송신 스키마파일: {first_interface.get('send_schema', 'N/A')}")
                print(f"수신 스키마파일: {first_interface.get('recv_schema', 'N/A')}")
            
            print("\n=== 테스트 완료 ===")
            
        except FileNotFoundError as e:
            print(f"파일 오류: {e}")
        except ValueError as e:
            print(f"데이터 오류: {e}")
        except Exception as e:
            print(f"예상치 못한 오류: {e}")
    
    # 간단한 사용법 예시
    def usage_example():
        """모듈 사용법 예시"""
        print("\n=== 사용법 예시 ===")
        print("""
# 1. InterfaceExcelReader 인스턴스 생성
reader = InterfaceExcelReader()

# 2. 엑셀 파일에서 인터페이스 정보 로드
interfaces = reader.load_interfaces('your_excel_file.xlsx')

# 3. 결과 활용
for interface in interfaces:
    print(f"인터페이스: {interface['interface_name']}")
    print(f"ID: {interface['interface_id']}")
    print(f"송신 테이블: {interface['send']['table_name']}")
    print(f"수신 테이블: {interface['recv']['table_name']}")

# 4. 처리 통계 확인
stats = reader.get_statistics()
print(f"처리된 인터페이스 수: {stats['processed_count']}")

# 5. BW 수신파일(.process) 파싱
insert_queries = parse_bw_receive_file('your_bw_file.process')
for query in insert_queries:
    print(f"추출된 INSERT 쿼리: {query}")

# 6. BWProcessFileParser 클래스 직접 사용
bw_parser = BWProcessFileParser()
queries = bw_parser.parse_bw_process_file('your_bw_file.process')
bw_stats = bw_parser.get_statistics()
print(f"BW 파싱 통계: {bw_stats}")
        """)
    
    # BW Process 파일 파싱 테스트 함수 추가
    def test_bw_process_parser():
        """BWProcessFileParser 테스트 함수"""
        print("\n=== BW Process 파일 파서 테스트 시작 ===")
        
        # 테스트할 .process 파일 경로 (실제 환경에 맞게 수정 필요)
        test_process_path = "sample.process"
        
        try:
            if not os.path.exists(test_process_path):
                print(f"테스트 파일이 없습니다: {test_process_path}")
                print("테스트를 위해 실제 .process 파일 경로를 지정해주세요.")
                return
            
            # BW 수신파일 파싱
            insert_queries = parse_bw_receive_file(test_process_path)
            
            # 결과 출력
            print(f"\n=== 처리 결과 ===")
            print(f"총 {len(insert_queries)}개의 INSERT 쿼리를 추출했습니다.")
            
            # 추출된 쿼리들 출력
            for i, query in enumerate(insert_queries, 1):
                print(f"\n=== INSERT 쿼리 {i} ===")
                print(query)
            
            print("\n=== BW Process 파일 파싱 테스트 완료 ===")
            
        except FileNotFoundError as e:
            print(f"파일 오류: {e}")
        except ValueError as e:
            print(f"데이터 오류: {e}")
        except Exception as e:
            print(f"예상치 못한 오류: {e}")
    
    # 테스트 실행
    test_interface_reader()
    usage_example()
    
    # 새로운 BW Process 파서 테스트 실행
    test_bw_process_parser()