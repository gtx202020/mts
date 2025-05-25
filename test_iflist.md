[역할 및 목표]
당신은 코드 리팩토링 및 모듈화에 특화된 전문가 파이썬 개발자입니다. 지금부터 여러 개의 파이썬 소스 파일을 제공할 것입니다. 당신의 임무는 이 파일들에서 특정 기능을 정확히 식별하고, 이를 재사용 가능하며 독립적인 단일 파이썬 모듈로 만들어내는 것입니다.

당신의 답변은 여러 LLM의 성능을 벤치마킹하는 데 사용될 것이며, 아래 평가 기준에 따라 분석될 것입니다. 따라서 코드의 정확성뿐만 아니라, 설계의 우수성과 설명의 명확성까지 모두 고려하여 답변해 주십시오.

[제공되는 소스 코드]
#############    
##test24.py
#############    

import sys
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import xml.etree.ElementTree as ET
import ast
from typing import Dict, List, Tuple, Optional

# comp_xml.py와 comp_q.py에서 필요한 클래스와 함수 import
from comp_xml import read_interface_block, XMLComparator
from comp_q import QueryParser

class InterfaceXMLToExcel:
    def __init__(self, excel_path: str, xml_dir: str, output_path: str = 'test24.xlsx'):
        """
        XML 파일에서 추출한 쿼리의 컬럼과 VALUES를 매핑하여 Excel 파일을 생성하는 클래스
        
        Args:
            excel_path (str): 인터페이스 정보가 있는 Excel 파일 경로
            xml_dir (str): XML 파일이 있는 디렉토리 경로
            output_path (str): 출력할 Excel 파일 경로
        """
        self.excel_path = excel_path
        self.xml_dir = xml_dir
        self.output_path = output_path
        self.query_parser = QueryParser()
        
        # Excel 파일 로드
        self.input_workbook = openpyxl.load_workbook(excel_path)
        self.input_worksheet = self.input_workbook.active
        
        # 출력 Excel 파일 생성
        self.output_workbook = openpyxl.Workbook()
        self.output_worksheet = self.output_workbook.active
        
        # 스타일 정의
        self.header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        self.header_font = Font(color='FFFFFF', bold=True, size=9)
        self.normal_font = Font(name='맑은 고딕', size=9)
        self.bold_font = Font(bold=True, size=9)
        self.center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        self.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                           top=Side(style='thin'), bottom=Side(style='thin'))
    
    def find_rcv_file(self, if_id: str) -> str:
        """
        주어진 인터페이스 ID에 해당하는 수신(RCV) XML 파일을 찾습니다.
        
        Args:
            if_id (str): 인터페이스 ID
            
        Returns:
            str: 찾은 파일의 경로, 없으면 None
        """
        if not if_id:
            print(f"Warning: Empty IF_ID provided")
            return None
            
        try:
            # 디렉토리 내의 모든 XML 파일 검색
            for file in os.listdir(self.xml_dir):
                if not file.startswith(if_id):
                    continue
                    
                # 수신 파일 (.RCV.xml)
                if file.endswith('.RCV.xml'):
                    file_path = os.path.join(self.xml_dir, file)
                    return file_path
            
            print(f"Warning: No receive file found for IF_ID: {if_id}")
            return None
            
        except Exception as e:
            print(f"Error finding interface files: {e}")
            return None
    
    def extract_query_from_xml(self, xml_path: str) -> str:
        """
        XML 파일에서 SQL 쿼리를 추출합니다.
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            str: 추출된 SQL 쿼리, 없으면 None
        """
        try:
            # XML 파일이 제대로 로드되었는지 확인
            if not os.path.exists(xml_path):
                print(f"Warning: XML file not found: {xml_path}")
                return None
                
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # XML 내용이 유효한지 확인
            if root is None:
                print(f"Warning: Invalid XML content in file: {xml_path}")
                return None
            
            # SQL 노드 찾기
            sql_node = root.find(".//SQL")
            if sql_node is None or not sql_node.text:
                print(f"Warning: No SQL content found in file: {xml_path}")
                return None
                
            query = sql_node.text.strip()
            
            # 추출된 쿼리가 유효한지 확인
            if not query:
                print(f"Warning: Empty SQL query in file: {xml_path}")
                return None
                
            return query
            
        except ET.ParseError as e:
            print(f"Error parsing XML file {xml_path}: {e}")
            return None
        except Exception as e:
            print(f"Unexpected error processing file {xml_path}: {e}")
            return None
    
    def clean_value(self, value: str) -> str:
        """
        VALUES 항목에서 콜론(:)을 제거하고 TO_DATE 함수 내의 실제 값만 추출합니다.
        
        Args:
            value (str): 원본 값
            
        Returns:
            str: 정제된 값
        """
        if not value:
            return ""
        
        # 콜론(:) 제거
        cleaned_value = value.replace(':', '')
        
        # TO_DATE 함수 처리
        to_date_pattern = r'TO_DATE\(\s*:?([A-Za-z0-9_]+)\s*,\s*[\'"](.*?)[\'"](\s*\))'
        to_date_match = re.search(to_date_pattern, value, re.IGNORECASE)
        
        if to_date_match:
            # TO_DATE 함수에서 첫 번째 인자(실제 값)만 추출
            param_name = to_date_match.group(1)
            return param_name
        
        return cleaned_value
    
    def get_column_value_mapping(self, query: str) -> Dict[str, str]:
        """
        INSERT 쿼리에서 컬럼과 VALUES를 매핑합니다.
        
        Args:
            query (str): INSERT SQL 쿼리
            
        Returns:
            Dict[str, str]: 컬럼과 값의 매핑 딕셔너리
        """
        if not query:
            return {}
            
        # QueryParser를 사용하여 INSERT 쿼리 파싱
        insert_parts = self.query_parser.parse_insert_parts(query)
        if not insert_parts:
            print(f"Failed to parse INSERT query: {query}")
            return {}
            
        # 테이블 이름과 컬럼-값 매핑 추출
        table_name, columns = insert_parts
        
        # 값 정제 처리
        cleaned_columns = {}
        for col, val in columns.items():
            cleaned_columns[col] = self.clean_value(val)
            
        return cleaned_columns
    
    def process_interfaces(self):
        """
        Excel 파일에서 인터페이스 정보를 읽고, XML 파일에서 쿼리를 추출하여 매핑 후 출력 Excel 파일에 작성합니다.
        """
        # 헤더 행 복사 및 스타일 적용
        for col in range(1, self.input_worksheet.max_column + 1):
            self.output_worksheet.cell(row=1, column=col).value = self.input_worksheet.cell(row=1, column=col).value
            self.output_worksheet.cell(row=1, column=col).font = self.normal_font
            
            self.output_worksheet.cell(row=2, column=col).value = self.input_worksheet.cell(row=2, column=col).value
            self.output_worksheet.cell(row=2, column=col).font = self.normal_font
            
            self.output_worksheet.cell(row=3, column=col).value = self.input_worksheet.cell(row=3, column=col).value
            self.output_worksheet.cell(row=3, column=col).font = self.normal_font
            
            self.output_worksheet.cell(row=4, column=col).value = self.input_worksheet.cell(row=4, column=col).value
            self.output_worksheet.cell(row=4, column=col).font = self.normal_font
        
        # 인터페이스 블록 처리
        current_col = 2  # B열부터 시작
        interface_count = 0
        
        while current_col <= self.input_worksheet.max_column:
            try:
                # 인터페이스 정보 읽기
                interface_info = read_interface_block(self.input_worksheet, current_col)
                if not interface_info:
                    break
                    
                interface_count += 1
                interface_id = interface_info.get('interface_id', '')
                interface_name = interface_info.get('interface_name', f'Interface_{interface_count}')
                
                print(f"\n처리 중인 인터페이스: {interface_name} (ID: {interface_id})")
                
                # 인터페이스 기본 정보 복사 및 스타일 적용
                self.output_worksheet.cell(row=1, column=current_col).value = interface_name
                self.output_worksheet.cell(row=1, column=current_col).font = self.normal_font
                
                self.output_worksheet.cell(row=2, column=current_col).value = interface_id
                self.output_worksheet.cell(row=2, column=current_col).font = self.normal_font
                
                self.output_worksheet.cell(row=3, column=current_col).value = self.input_worksheet.cell(row=3, column=current_col).value
                self.output_worksheet.cell(row=3, column=current_col).font = self.normal_font
                self.output_worksheet.cell(row=3, column=current_col + 1).value = self.input_worksheet.cell(row=3, column=current_col + 1).value
                self.output_worksheet.cell(row=3, column=current_col + 1).font = self.normal_font
                
                self.output_worksheet.cell(row=4, column=current_col).value = self.input_worksheet.cell(row=4, column=current_col).value
                self.output_worksheet.cell(row=4, column=current_col).font = self.normal_font
                self.output_worksheet.cell(row=4, column=current_col + 1).value = self.input_worksheet.cell(row=4, column=current_col + 1).value
                self.output_worksheet.cell(row=4, column=current_col + 1).font = self.normal_font
                
                # 수신 XML 파일 찾기
                rcv_file_path = self.find_rcv_file(interface_id)
                if not rcv_file_path:
                    print(f"Warning: No receive file found for interface {interface_name} (ID: {interface_id})")
                    current_col += 3  # 다음 인터페이스로 이동
                    continue
                
                # XML 파일에서 쿼리 추출
                query = self.extract_query_from_xml(rcv_file_path)
                if not query:
                    print(f"Warning: Failed to extract query from file {rcv_file_path}")
                    current_col += 3  # 다음 인터페이스로 이동
                    continue
                
                # 쿼리에서 컬럼-값 매핑 추출
                column_value_mapping = self.get_column_value_mapping(query)
                if not column_value_mapping:
                    print(f"Warning: Failed to extract column-value mapping from query")
                    current_col += 3  # 다음 인터페이스로 이동
                    continue
                
                # 특수 컬럼 제외
                special_columns = set(self.query_parser.special_columns['recv']['required'])
                filtered_mapping = {k: v for k, v in column_value_mapping.items() if k.upper() not in special_columns}
                
                # 매핑 정보를 Excel에 작성
                row = 5  # 5행부터 컬럼 매핑 시작
                for column, value in filtered_mapping.items():
                    # 수신 컬럼을 첫 번째 열(B열)에 배치
                    self.output_worksheet.cell(row=row, column=current_col).value = column  # 수신 컬럼을 첫 번째 열에 배치
                    self.output_worksheet.cell(row=row, column=current_col).font = self.normal_font
                    
                    # VALUES 항목을 오른쪽 열(C열)에 배치 - 이미 정제된 값 사용
                    self.output_worksheet.cell(row=row, column=current_col + 1).value = value  # VALUES 항목 (콜론 제거와 TO_DATE 함수 처리가 적용됨)
                    self.output_worksheet.cell(row=row, column=current_col + 1).font = self.normal_font
                    
                    row += 1
                
                print(f"인터페이스 {interface_name} (ID: {interface_id}) 처리 완료")
                
            except Exception as e:
                print(f"Error processing interface at column {current_col}: {str(e)}")
            
            current_col += 3  # 다음 인터페이스로 이동
        
        # 출력 파일 저장
        self.output_workbook.save(self.output_path)
        self.input_workbook.close()
        self.output_workbook.close()
        
        print(f"\n=== 처리 완료 ===")
        print(f"총 처리된 인터페이스 수: {interface_count}")
        print(f"출력 파일 저장 완료: {self.output_path}")

def main():
    try:
        # 하드코딩된 경로 설정 (python test24.py만 실행해도 작동하도록)
        # 현재 스크립트가 있는 디렉토리 기준으로 상대 경로 설정
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 기본 경로 설정 (실제 환경에 맞게 수정 필요)
        excel_path = os.path.join(current_dir, 'C:\\work\\LT\\input_W7.xlsx')  # 인터페이스 정보 파일
        xml_dir = os.path.join(current_dir, 'C:\\work\\LT\\W7xml')  # XML 파일 디렉토리
        output_path = os.path.join(current_dir, 'C:\\work\\LT\\test24.xlsx')  # 출력 파일
        
        # 명령행 인수가 있으면 덮어쓰기
        if len(sys.argv) > 1:
            excel_path = sys.argv[1]
        if len(sys.argv) > 2:
            xml_dir = sys.argv[2]
        if len(sys.argv) > 3:
            output_path = sys.argv[3]
            
        print(f"사용할 파일 경로:")
        print(f"- 인터페이스 정보 파일: {excel_path}")
        print(f"- XML 파일 디렉토리: {xml_dir}")
        print(f"- 출력 파일: {output_path}")
        
        # 인터페이스 처리 실행
        processor = InterfaceXMLToExcel(excel_path, xml_dir, output_path)
        processor.process_interfaces()
        
    except Exception as e:
        print(f"\n[심각한 오류] 프로그램 실행 중 오류 발생: {str(e)}")
        raise

if __name__ == "__main__":
    main()
#############    
##comp_xml.py
#############    
import sys
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from typing import Dict, List, Tuple, Optional
import xml.etree.ElementTree as ET
from comp_excel import ExcelManager, read_interface_block
from xltest import process_interface, read_interface_block
from comp_q import QueryParser, QueryDifference, FileSearcher, BWQueryExtractor
from maptest import ColumnMapper
import datetime
import ast

def read_interface_block(ws, start_col):
    """Excel에서 3컬럼 단위로 하나의 인터페이스 정보를 읽습니다.
    이 함수는 xltest.py의 동일한 함수를 대체하지 않고, 가져오지 못한 경우의 백업 역할만 합니다.
    """
    try:
        interface_info = {
            'interface_name': ws.cell(row=1, column=start_col).value or '',  # IF NAME (1행)
            'interface_id': ws.cell(row=2, column=start_col).value or '',    # IF ID (2행)
            'send': {'owner': None, 'table_name': None, 'columns': [], 'db_info': None},
            'recv': {'owner': None, 'table_name': None, 'columns': [], 'db_info': None}
        }
        
        # 인터페이스 ID가 없으면 빈 인터페이스로 간주
        if not interface_info['interface_id']:
            return None
            
        # DB 연결 정보 (3행에서 읽기)
        try:
            send_db_value = ws.cell(row=3, column=start_col).value
            send_db_info = ast.literal_eval(send_db_value) if send_db_value else {}
            
            recv_db_value = ws.cell(row=3, column=start_col + 1).value
            recv_db_info = ast.literal_eval(recv_db_value) if recv_db_value else {}
        except (SyntaxError, ValueError):
            # 데이터 형식 오류 시 빈 딕셔너리로 설정
            send_db_info = {}
            recv_db_info = {}
            
        interface_info['send']['db_info'] = send_db_info
        interface_info['recv']['db_info'] = recv_db_info
        
        # 테이블 정보 (4행에서 읽기)
        try:
            send_table_value = ws.cell(row=4, column=start_col).value
            send_table_info = ast.literal_eval(send_table_value) if send_table_value else {}
            
            recv_table_value = ws.cell(row=4, column=start_col + 1).value
            recv_table_info = ast.literal_eval(recv_table_value) if recv_table_value else {}
        except (SyntaxError, ValueError):
            # 데이터 형식 오류 시 빈 딕셔너리로 설정
            send_table_info = {}
            recv_table_info = {}
        
        interface_info['send']['owner'] = send_table_info.get('owner')
        interface_info['send']['table_name'] = send_table_info.get('table_name')
        interface_info['recv']['owner'] = recv_table_info.get('owner')
        interface_info['recv']['table_name'] = recv_table_info.get('table_name')
        
        # 컬럼 매핑 정보 (5행부터)
        row = 5
        while True:
            send_col = ws.cell(row=row, column=start_col).value
            recv_col = ws.cell(row=row, column=start_col + 1).value
            
            if not send_col and not recv_col:
                break
                
            interface_info['send']['columns'].append(send_col if send_col else '')
            interface_info['recv']['columns'].append(recv_col if recv_col else '')
            row += 1
            
    except Exception as e:
        print(f'인터페이스 정보 읽기 중 오류 발생: {str(e)}')
        return None
    
    return interface_info

class XMLComparator:
    # 클래스 변수로 BW_SEARCH_DIR 정의
    BW_SEARCH_DIR = "C:\\work\\LT\\BW소스"

    def __init__(self, excel_path: str, search_dir: str):
        """
        XML 비교를 위한 클래스 초기화
        
        Args:
            excel_path (str): 인터페이스 정보가 있는 Excel 파일 경로
            search_dir (str): XML 파일을 검색할 디렉토리 경로
        """
        self.excel_path = excel_path
        self.search_dir = search_dir
        self.workbook = openpyxl.load_workbook(excel_path)
        self.worksheet = self.workbook.active
        self.mapper = ColumnMapper()
        self.query_parser = QueryParser()  # QueryParser 인스턴스 생성
        self.excel_manager = ExcelManager()  # ExcelManager 인스턴스 생성
        self.interface_results = []  # 모든 인터페이스 처리 결과 저장
        self.output_path = 'C:\\work\\LT\\comp_mq_bw.xlsx'  # 기본 출력 경로

    def extract_from_xml(self, xml_path: str) -> Tuple[str, str]:
        """
        XML 파일에서 쿼리와 XML 내용을 추출합니다.
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            Tuple[str, str]: (쿼리, XML 내용)
        """
        try:
            # XML 파일이 제대로 로드되었는지 확인
            if not os.path.exists(xml_path):
                print(f"Warning: XML file not found: {xml_path}")
                return None, None
                
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # XML 내용이 유효한지 확인
            if root is None:
                print(f"Warning: Invalid XML content in file: {xml_path}")
                return None, None
            
            # SQL 노드 찾기
            sql_node = root.find(".//SQL")
            if sql_node is None or not sql_node.text:
                print(f"Warning: No SQL content found in file: {xml_path}")
                return None, None
                
            query = sql_node.text.strip()
            xml_content = ET.tostring(root, encoding='unicode')
            
            # 추출된 쿼리가 유효한지 확인
            if not query:
                print(f"Warning: Empty SQL query in file: {xml_path}")
                return None, None
                
            return query, xml_content
            
        except ET.ParseError as e:
            print(f"Error parsing XML file {xml_path}: {e}")
            return None, None
        except Exception as e:
            print(f"Unexpected error processing file {xml_path}: {e}")
            return None, None
            
    def compare_queries(self, query1: str, query2: str) -> QueryDifference:
        """
        두 쿼리를 비교합니다.
        
        Args:
            query1 (str): 첫 번째 쿼리
            query2 (str): 두 번째 쿼리
            
        Returns:
            QueryDifference: 쿼리 비교 결과
        """
        if not query1 or not query2:
            return None
        return self.query_parser.compare_queries(query1, query2)
        
    def find_interface_files(self, if_id: str) -> Dict[str, Dict]:
        """
        주어진 IF ID에 해당하는 송수신 XML 파일을 찾고 쿼리를 추출합니다.
        파일명 패턴: {if_id}로 시작하고 .SND.xml 또는 .RCV.xml로 끝나는 파일
        
        Args:
            if_id (str): 인터페이스 ID
            
        Returns:
            Dict[str, Dict]: {
                'send': {'path': 송신파일경로, 'query': 송신쿼리, 'xml': 송신XML},
                'recv': {'path': 수신파일경로, 'query': 수신쿼리, 'xml': 수신XML}
            }
        """
        results = {
            'send': {'path': None, 'query': None, 'xml': None},
            'recv': {'path': None, 'query': None, 'xml': None}
        }
        
        if not if_id:
            print("Warning: Empty IF_ID provided")
            return results
            
        try:
            # 디렉토리 내의 모든 XML 파일 검색
            for file in os.listdir(self.search_dir):
                if not file.startswith(if_id):
                    continue
                    
                file_path = os.path.join(self.search_dir, file)
                
                # 송신 파일 (.SND.xml)
                if file.endswith('.SND.xml'):
                    results['send']['path'] = file_path
                    query, xml = self.extract_from_xml(file_path)
                    if query and xml:
                        results['send']['query'] = query
                        results['send']['xml'] = xml
                    else:
                        print(f"Warning: Failed to extract query from send file: {file_path}")
                
                # 수신 파일 (.RCV.xml)
                elif file.endswith('.RCV.xml'):
                    results['recv']['path'] = file_path
                    query, xml = self.extract_from_xml(file_path)
                    if query and xml:
                        results['recv']['query'] = query
                        results['recv']['xml'] = xml
                    else:
                        print(f"Warning: Failed to extract query from receive file: {file_path}")
            
            # 파일을 찾았는지 확인
            if not results['send']['path'] and not results['recv']['path']:
                print(f"Warning: No interface files found for IF_ID: {if_id}")
            elif not results['send']['path']:
                print(f"Warning: No send file found for IF_ID: {if_id}")
            elif not results['recv']['path']:
                print(f"Warning: No receive file found for IF_ID: {if_id}")
            
            return results
            
        except Exception as e:
            print(f"Error finding interface files: {e}")
            return results
        
    def process_interface_block(self, start_col: int) -> Optional[Dict]:
        """
        Excel에서 하나의 인터페이스 블록을 처리합니다.
        
        Args:
            start_col (int): 인터페이스 블록이 시작되는 컬럼
            
        Returns:
            Optional[Dict]: 처리된 인터페이스 정보와 결과, 실패시 None
        """
        try:
            # Excel에서 인터페이스 정보 읽기
            interface_info = read_interface_block(self.worksheet, start_col)
            if not interface_info:
                print(f"Warning: Failed to read interface block at column {start_col}")
                return None
                
            # Excel에서 추출된 쿼리와 XML 얻기
            excel_results = process_interface(interface_info, self.mapper)
            if not excel_results:
                print(f"Warning: Failed to process interface at column {start_col}")
                return None
                
            # 송수신 파일 찾기
            file_results = self.find_interface_files(interface_info['interface_id'])
            if not file_results:
                print(f"Warning: No interface files found for IF_ID: {interface_info['interface_id']}")
                return None
            
            # 결과 초기화
            comparisons = {
                'send': None,
                'recv': None
            }
            warnings = {
                'send': [],
                'recv': []
            }
            
            # 송신 쿼리 처리
            if excel_results['send_sql'] and file_results['send']['query']:
                try:
                    comparisons['send'] = self.query_parser.compare_queries(
                        excel_results['send_sql'],
                        file_results['send']['query']
                    )
                    warnings['send'].extend(
                        self.query_parser.check_special_columns(
                            file_results['send']['query'],
                            'send'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing send queries: {e}")
                    print(f"Excel query: {excel_results['send_sql']}")
                    print(f"File query: {file_results['send']['query']}")
            
            # 수신 쿼리 처리
            if excel_results['recv_sql'] and file_results['recv']['query']:
                try:
                    comparisons['recv'] = self.query_parser.compare_queries(
                        excel_results['recv_sql'],
                        file_results['recv']['query']
                    )
                    warnings['recv'].extend(
                        self.query_parser.check_special_columns(
                            file_results['recv']['query'],
                            'recv'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing receive queries: {e}")
                    print(f"Excel query: {excel_results['recv_sql']}")
                    print(f"File query: {file_results['recv']['query']}")
            
            return {
                'if_id': interface_info['interface_id'],
                'interface_name': interface_info['interface_name'],
                'comparisons': comparisons,
                'warnings': warnings,
                'excel': excel_results,
                'files': file_results
            }
            
        except Exception as e:
            print(f"Error processing interface block at column {start_col}: {e}")
            return None
            
    def process_all_interfaces(self) -> List[Dict]:
        """
        Excel 파일의 모든 인터페이스를 처리합니다.
        B열부터 시작하여 3컬럼 단위로 처리합니다.
        
        Returns:
            List[Dict]: 각 인터페이스의 처리 결과 목록
        """
        results = []
        col = 2  # B열부터 시작
        
        while True:
            # 인터페이스 ID가 없으면 종료
            if not self.worksheet.cell(row=2, column=col).value:
                break
                
            result = self.process_interface_block(col)
            if result:
                results.append(result)
                
            col += 3  # 다음 인터페이스 블록으로 이동
            
        # 결과 출력
        for idx, result in enumerate(results, 1):
            print(f"\n=== 인터페이스 {idx} ===")
            print(f"ID: {result['if_id']}")
            print(f"이름: {result['interface_name']}")
            
            print("\n파일 검색 결과:")
            print(f"송신 파일: {result['files']['send']['path']}")
            print(f"수신 파일: {result['files']['recv']['path']}")
            
            print("\n쿼리 비교 결과:")
            if result['comparisons']['send']:
                print("송신 쿼리:")
                print(f"  {result['comparisons']['send']}")
            if result['comparisons']['recv']:
                print("수신 쿼리:")
                print(f"  {result['comparisons']['recv']}")
            
            # 경고가 있을 때만 경고 섹션 출력
            send_warnings = result['warnings']['send']
            recv_warnings = result['warnings']['recv']
            if send_warnings or recv_warnings:
                print("\n경고:")
                if send_warnings:
                    print("송신 쿼리 경고:")
                    for warning in send_warnings:
                        print(f"  - {warning}")
                if recv_warnings:
                    print("수신 쿼리 경고:")
                    for warning in recv_warnings:
                        print(f"  - {warning}")
                
        return results
        
    def close(self):
        """리소스 정리"""
        self.workbook.close()
        if self.mapper:
            self.mapper.close_connections()

    def find_bw_files(self) -> List[Dict[str, str]]:
        """
        엑셀의 인터페이스 정보에서 송신 테이블명을 추출하여 BW 파일을 검색합니다.
        
        Returns:
            List[Dict[str, str]]: [
                {
                    'interface_name': str,
                    'interface_id': str,
                    'send_table': str,
                    'bw_files': List[str]
                },
                ...
            ]
        """
        results = []
        file_searcher = FileSearcher()
        
        # 엑셀에서 인터페이스 정보 읽기
        for row in range(2, self.worksheet.max_row + 1, 3):  # 3행씩 건너뛰며 읽기
            interface_info = read_interface_block(self.worksheet, row)
            if not interface_info:
                continue
                
            # 송신 테이블명 추출 (스키마/오너 제외)
            send_table = interface_info['send'].get('table_name')
            if not send_table:
                continue
                
            # BW 파일 검색 - self.BW_SEARCH_DIR 사용
            bw_files = file_searcher.find_files_with_keywords(self.BW_SEARCH_DIR, [send_table])
            matching_files = bw_files.get(send_table, [])
            
            results.append({
                'interface_name': interface_info['interface_name'],
                'interface_id': interface_info['interface_id'],
                'send_table': send_table,
                'bw_files': matching_files
            })
            
        return results
        
    def print_bw_search_results(self, results: List[Dict[str, str]]):
        """
        BW 파일 검색 결과를 출력합니다.
        
        Args:
            results (List[Dict[str, str]]): find_bw_files()의 반환값
        """
        print("\nBW File Search Results:")
        print("-" * 80)
        print(f"{'Interface Name':<30} {'Interface ID':<15} {'Send Table':<20} {'BW Files'}")
        print("-" * 80)
        
        for result in results:
            bw_files_str = ', '.join(result['bw_files']) if result['bw_files'] else 'No matching files'
            print(f"{result['interface_name']:<30} {result['interface_id']:<15} {result['send_table']:<20} {bw_files_str}")

    def initialize_excel_output(self):
        """
        결과를 저장할 새 엑셀 파일 초기화
        """
        # ExcelManager를 통해 Excel 출력을 초기화
        self.excel_manager.initialize_excel_output()
        
    def save_excel_output(self, output_path=None):
        """
        처리된 결과를 엑셀 파일로 저장
        
        Args:
            output_path (str, optional): 출력 엑셀 파일 경로, 없으면 기본 경로 사용
            
        Returns:
            bool: 저장 성공 여부
        """
        # output_path 값을 인스턴스 변수에 저장
        if output_path:
            self.output_path = output_path
            
        # ExcelManager를 사용하여 파일 저장
        return self.excel_manager.save_excel_output(self.output_path)
        
    def create_interface_sheet(self, if_info, file_results, query_comparisons, bw_queries=None, bw_files=None):
        """
        인터페이스 정보와 비교 결과를 포함하는 엑셀 시트를 생성합니다.
        
        Args:
            if_info (dict): 인터페이스 정보
            file_results (dict): MQ 파일 결과 (송신/수신)
            query_comparisons (dict): 쿼리 비교 결과 (송신/수신)
            bw_queries (dict, optional): BW 쿼리 정보. Defaults to None.
            bw_files (list, optional): BW 매핑 파일 목록. Defaults to None.
        """
        # 기본값 설정
        bw_queries = bw_queries or {'send': '', 'recv': ''}
        bw_files = bw_files or []
        
        # 인터페이스 ID와 이름 확인
        if 'interface_id' not in if_info or not if_info['interface_id']:
            print("인터페이스 ID가 없습니다.")
            return
            
        # BW 파일 매핑
        bw_files_dict = {
            'send': bw_files[0] if bw_files and len(bw_files) > 0 else 'N/A',
            'recv': bw_files[1] if bw_files and len(bw_files) > 1 else 'N/A'
        }
        
        # MQ 파일 정보
        mq_files = {
            'send': file_results.get('send', {}),
            'recv': file_results.get('recv', {})
        }
        
        # 쿼리 정보 구성
        queries = {
            'mq_send': file_results.get('send', {}).get('query', 'N/A'),
            'bw_send': bw_queries.get('send', 'N/A'),
            'mq_recv': file_results.get('recv', {}).get('query', 'N/A'),
            'bw_recv': bw_queries.get('recv', 'N/A')
        }
        
        # 비교 결과 구성
        comparison_results = {
            'send': {
                'is_equal': query_comparisons.get('send', QueryDifference()).is_equal,
                'detail': self._get_difference_detail(query_comparisons.get('send', QueryDifference()))
            },
            'recv': {
                'is_equal': query_comparisons.get('recv', QueryDifference()).is_equal,
                'detail': self._get_difference_detail(query_comparisons.get('recv', QueryDifference()))
            }
        }
        
        # 인터페이스 시트 생성
        self.excel_manager.create_interface_sheet(if_info, mq_files, bw_files_dict, queries, comparison_results)
        
    def process_interface_with_bw(self, start_col: int, interface_info: Dict) -> Optional[Dict]:
        """
        하나의 인터페이스를 처리하고 BW 파일과 비교하여 결과 반환
        
        Args:
            start_col (int): 인터페이스 블록이 시작되는 컬럼
            interface_info (Dict): 인터페이스 정보
            
        Returns:
            Optional[Dict]: 처리된 인터페이스 정보와 결과, 실패시 None
        """
        try:
            # 표준 필드 생성
            # DB 정보에서 시스템 정보 추출
            if 'send' in interface_info and 'db_info' in interface_info['send'] and interface_info['send']['db_info']:
                interface_info['send_system'] = interface_info['send']['db_info'].get('system', 'N/A')
            else:
                interface_info['send_system'] = 'N/A'
                
            if 'recv' in interface_info and 'db_info' in interface_info['recv'] and interface_info['recv']['db_info']:
                interface_info['recv_system'] = interface_info['recv']['db_info'].get('system', 'N/A')
            else:
                interface_info['recv_system'] = 'N/A'
                
            # 테이블 정보 추출
            if 'send' in interface_info and 'table_name' in interface_info['send']:
                interface_info['send_table'] = interface_info['send']['table_name']
            else:
                interface_info['send_table'] = ''
                
            if 'recv' in interface_info and 'table_name' in interface_info['recv']:
                interface_info['recv_table'] = interface_info['recv']['table_name']
            else:
                interface_info['recv_table'] = ''
                
            # Excel에서 추출된 쿼리와 XML 얻기
            excel_results = process_interface(interface_info, self.mapper)
            if not excel_results:
                print(f"Warning: Failed to process interface at column {start_col}")
                return None
                
            # 송수신 파일 찾기
            file_results = self.find_interface_files(interface_info['interface_id'])
            if not file_results:
                print(f"Warning: No interface files found for IF_ID: {interface_info['interface_id']}")
                return None
            
            # BW 파일 찾기
            send_table = interface_info.get('send_table', '')
            if not send_table:
                print(f"Warning: No send table information for IF_ID: {interface_info['interface_id']}")
                bw_files = []
            else:
                # 송신 테이블로 BW 파일 검색
                bw_searcher = FileSearcher()
                bw_files = bw_searcher.find_files_with_keywords(
                    self.BW_SEARCH_DIR, 
                    [send_table]
                )
            
            # bw_files가 사전 형태이므로 send_table 키워드에 대한 결과를 가져옴
            matching_files = bw_files.get(send_table, [])
            
            # BW 쿼리 추출
            bw_queries = {
                'send': '',
                'recv': ''
            }
            extractor = BWQueryExtractor()
            for bw_file in matching_files:
                bw_file_path = os.path.join(self.BW_SEARCH_DIR, bw_file)
                if os.path.exists(bw_file_path):
                    # BWQueryExtractor의 extract_bw_queries 메서드를 사용하여 송신/수신 쿼리 모두 추출
                    queries = extractor.extract_bw_queries(bw_file_path)
                    
                    # 송신 쿼리가 없으면 첫 번째 송신 쿼리 저장
                    if not bw_queries['send'] and queries.get('send') and len(queries['send']) > 0:
                        bw_queries['send'] = queries['send'][0]
                    
                    # 수신 쿼리가 없으면 첫 번째 수신 쿼리 저장
                    if not bw_queries['recv'] and queries.get('recv') and len(queries['recv']) > 0:
                        bw_queries['recv'] = queries['recv'][0]
            
            # 결과 초기화
            comparisons = {
                'send': None,
                'recv': None
            }
            warnings = {
                'send': [],
                'recv': []
            }
            
            # 송신 쿼리 비교 (MQ XML vs BW XML)
            if file_results['send']['query'] and bw_queries['send']:
                try:
                    comparisons['send'] = self.query_parser.compare_queries(
                        file_results['send']['query'],
                        bw_queries['send']
                    )
                    warnings['send'].extend(
                        self.query_parser.check_special_columns(
                            file_results['send']['query'],
                            'send'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing send queries: {e}")
                    print(f"MQ query: {file_results['send']['query']}")
                    print(f"BW query: {bw_queries['send']}")
            
            # 수신 쿼리 비교 (MQ XML vs BW XML)
            if file_results['recv']['query'] and bw_queries['recv']:
                try:
                    comparisons['recv'] = self.query_parser.compare_queries(
                        file_results['recv']['query'],
                        bw_queries['recv']
                    )
                    warnings['recv'].extend(
                        self.query_parser.check_special_columns(
                            file_results['recv']['query'],
                            'recv'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing recv queries: {e}")
                    print(f"MQ query: {file_results['recv']['query']}")
                    print(f"BW query: {bw_queries['recv']}")
            
            # 결과 반환
            return {
                'interface_info': interface_info,
                'excel_results': excel_results,
                'file_results': file_results,
                'bw_queries': bw_queries,
                'comparisons': comparisons,
                'warnings': warnings,
                'bw_files': matching_files
            }
            
        except Exception as e:
            print(f"Error processing interface at column {start_col}: {e}")
            import traceback
            traceback.print_exc()
            return None

    def process_all_interfaces_with_bw(self):
        """
        모든 인터페이스를 처리하고 BW 파일과 비교하여 엑셀 파일로 결과 저장
        """
        # 엑셀 파일 초기화 - ExcelManager 사용
        self.excel_manager.initialize_excel_output()
        
        # 모든 열을 처리
        print("\n[인터페이스 처리 시작]")
        print("-" * 80)
        
        interface_count = 0
        processed_count = 0
        
        start_col = 2
        while True:
            interface_info = read_interface_block(self.worksheet, start_col)
            
            if not interface_info:
                break
                
            interface_count += 1
            
            # 인터페이스 ID와 이름 출력
            print(f"처리 중: [{interface_count}] {interface_info['interface_id']} - {interface_info['interface_name']}")
            
            # 인터페이스 처리 및 BW 비교
            result = self.process_interface_with_bw(start_col, interface_info)
            
            # 다음 인터페이스로 이동 (3칸씩)
            start_col += 3
            
            # 인터페이스 처리 결과가 있으면 엑셀에 저장
            if result:
                processed_count += 1
                
                # 결과를 저장할 인터페이스 시트 생성
                if_info = result['interface_info']
                
                # ExcelManager를 사용하여 인터페이스 시트 생성
                # MQ 파일 정보
                mq_files = {
                    'send': result['file_results']['send'],
                    'recv': result['file_results']['recv']
                }
                
                # BW 파일 정보
                bw_files = {
                    'send': result.get('bw_files', [])[0] if result.get('bw_files') and len(result.get('bw_files')) > 0 else 'N/A',
                    'recv': result.get('bw_files', [])[1] if result.get('bw_files') and len(result.get('bw_files')) > 1 else 'N/A'
                }
                
                # 쿼리 정보
                queries = {
                    'mq_send': result['file_results']['send']['query'],
                    'bw_send': result['bw_queries']['send'],
                    'mq_recv': result['file_results']['recv']['query'],
                    'bw_recv': result['bw_queries']['recv']
                }
                
                # 비교 결과
                comparison_results = {
                    'send': {
                        'is_equal': result['comparisons']['send'].is_equal if result['comparisons']['send'] else False,
                        'detail': self._get_difference_detail(result['comparisons']['send']) if result['comparisons']['send'] else '비교 불가'
                    },
                    'recv': {
                        'is_equal': result['comparisons']['recv'].is_equal if result['comparisons']['recv'] else False,
                        'detail': self._get_difference_detail(result['comparisons']['recv']) if result['comparisons']['recv'] else '비교 불가'
                    }
                }
                
                self.excel_manager.create_interface_sheet(
                    if_info, 
                    mq_files, 
                    bw_files, 
                    queries, 
                    comparison_results
                )
                
                # 요약 시트 업데이트
                self.update_summary_sheet(result, interface_count + 1)
        
        # 결과 저장
        self.save_excel_output()
        
        # 처리 결과 출력
        print("\n" + "=" * 80)
        print(f"처리 완료: 총 {interface_count}개 인터페이스 중 {processed_count}개 처리됨")
        print(f"결과 파일: {self.output_path}")
        print("=" * 80)
        
    def update_summary_sheet(self, result, row):
        """
        요약 시트에 현재 인터페이스 처리 결과를 추가합니다.
        
        Args:
            result (dict): 인터페이스 처리 결과
            row (int): 추가할 행 번호
        """
        # ExcelManager를 사용하여 요약 시트 업데이트
        self.excel_manager.update_summary_sheet(result, row)

    def extract_bw_queries(self, bw_results):
        """
        BW 파일에서 쿼리를 추출합니다.
        
        Args:
            bw_results (list): BW 파일 검색 결과 목록
            
        Returns:
            list: 인터페이스별 BW 쿼리 정보가 담긴 리스트
        """
        extractor = BWQueryExtractor()
        results = []
        
        for result in bw_results:
            if result['bw_files']:  # BW 파일이 있는 경우에만 처리
                print(f"\n인터페이스: {result['interface_name']} ({result['interface_id']})")
                print(f"송신 테이블: {result['send_table']}")
                print("찾은 BW 파일의 쿼리:")
                
                bw_queries = {'send': '', 'recv': ''}
                
                for bw_file in result['bw_files']:
                    bw_file_path = os.path.join(self.BW_SEARCH_DIR, bw_file)
                    if os.path.exists(bw_file_path):
                        # BWQueryExtractor의 extract_bw_queries 메서드를 사용하여 송신/수신 쿼리 모두 추출
                        queries = extractor.extract_bw_queries(bw_file_path)
                        
                        if queries['send'] and not bw_queries['send']:
                            bw_queries['send'] = queries['send'][0] if queries['send'] else ''
                            print(f"\nBW 송신 파일: {bw_file}")
                            print("-" * 40)
                            print(bw_queries['send'])
                            
                        if queries['recv'] and not bw_queries['recv']:
                            bw_queries['recv'] = queries['recv'][0] if queries['recv'] else ''
                            print(f"\nBW 수신 파일: {bw_file}")
                            print("-" * 40)
                            print(bw_queries['recv'])
                
                # 인터페이스 결과에 BW 쿼리 추가
                for interface_result in self.interface_results:
                    if interface_result['interface_info']['interface_id'] == result['interface_id']:
                        interface_result['bw_queries'] = bw_queries
                        interface_result['bw_files'] = result['bw_files']
                        break
                
                results.append({
                    'interface_id': result['interface_id'],
                    'bw_queries': bw_queries,
                    'bw_files': result['bw_files']
                })
        
        return results

    def _get_difference_detail(self, query_diff):
        """
        쿼리 차이점을 텍스트로 변환합니다.
        
        Args:
            query_diff (QueryDifference): 쿼리 차이점
        
        Returns:
            str: 차이점 텍스트
        """
        if not query_diff:
            return ''
        
        if query_diff.is_equal:
            return '일치 - 테이블과 칼럼이 모두 동일합니다.'
        
        # 차이점 텍스트 생성
        detail = '차이 - 다음과 같은 차이점이 발견되었습니다:\n'
        for diff in query_diff.differences:
            column = diff.get('column', 'N/A')
            query1_value = diff.get('query1_value', 'N/A')
            query2_value = diff.get('query2_value', 'N/A')
            detail += f'- 컬럼: {column}\n'
            detail += f'  · MQ: {query1_value}\n'
            detail += f'  · BW: {query2_value}\n'
        
        return detail

def main():
    # 고정된 경로 사용
    excel_path = 'C:\\work\\LT\\input_LT.xlsx' # 인터페이스 정보
    xml_dir = 'C:\\work\\LT\\xml' # MQ XML 파일 디렉토리
    bw_dir = 'C:\\work\\LT\\BW소스'  # BW XML 파일 디렉토리 경로
    output_path = 'C:\\work\\LT\\comp_mq_bw.xlsx'  # 출력 엑셀 파일 경로
    
    # BW 검색 디렉토리 설정
    XMLComparator.BW_SEARCH_DIR = bw_dir
    
    # XML 비교기 초기화
    comparator = XMLComparator(excel_path, xml_dir)
    
    # 명령행 인자가 있을 경우 처리
    if len(sys.argv) > 1:
        if sys.argv[1] == "excel":
            # 엑셀 출력 모드 실행
            print("\n[MQ XML과 BW XML 쿼리 비교 - 엑셀 출력 모드]")
            comparator.process_all_interfaces_with_bw()
            return
        elif len(sys.argv) > 2 and sys.argv[1] == "output":
            # 출력 경로 변경
            output_path = sys.argv[2]
            comparator.output_path = output_path
            print(f"\n[출력 경로 변경: {output_path}]")
    
    # 기본 모드 실행 - 기존 로직 유지
    print("\n[MQ XML 파일 검색 및 쿼리 비교 시작]")
    comparator.process_all_interfaces()
    
    # BW 파일 검색 및 결과 출력을 마지막으로 이동
    print("\n[BW 파일 검색 시작]")
    bw_results = comparator.find_bw_files()
    comparator.print_bw_search_results(bw_results)
    
    # BW 파일에서 쿼리 추출
    print("\n[BW 파일 쿼리 추출]")
    print("-" * 80)
    bw_queries = comparator.extract_bw_queries(bw_results)
    
    # 처리 결과를 Excel로 저장 (excel 모드가 아닌 경우)
    print("\n[결과를 Excel로 저장]")
    comparator.initialize_excel_output()
    
    # 인터페이스별 결과 처리
    for i, result in enumerate(comparator.interface_results):
        if_info = result['interface_info']
        
        # 인터페이스 시트 생성
        comparator.create_interface_sheet(
            if_info, 
            result['file_results'], 
            result['comparisons'],
            result.get('bw_queries', {'send': '', 'recv': ''}),
            result.get('bw_files', [])
        )
        
        # 요약 시트 업데이트
        comparator.update_summary_sheet(result, i + 2)
    
    # 결과 저장
    comparator.save_excel_output(output_path)
    print(f"\n[분석 완료] 결과가 저장되었습니다: {output_path}")
    
    print("\n[처리 완료]")
    print("엑셀 출력 모드로 실행하려면 'python comp_xml.py excel' 명령을 사용하세요.")

if __name__ == "__main__":
    main()

#############    
##comp_q.py
#############    
import xml.etree.ElementTree as ET
import re
from typing import Dict, List, Tuple, Optional
import os
import argparse

class QueryDifference:
    def __init__(self):
        self.is_equal = True
        self.differences = []
        self.query_type = None
        self.table_name = None
    
    def add_difference(self, column: str, value1: str, value2: str):
        self.is_equal = False
        self.differences.append({
            'column': column,
            'query1_value': value1,
            'query2_value': value2
        })

    def __str__(self) -> str:
        if self.is_equal:
            return "일치"
        
        return "불일치"

class QueryParser:
    # 특수 컬럼 정의를 클래스 변수로 변경
    special_columns = {
        'send': {
            'required': ['EAI_SEQ_ID', 'DATA_INTERFACE_TYPE_CODE'],
            'mappings': []  # 추가 매핑을 저장할 리스트
        },
        'recv': {
            'required': [
                'EAI_SEQ_ID',
                'DATA_INTERFACE_TYPE_CODE',
                'EAI_INTERFACE_DATE',
                'APPLICATION_TRANSFER_FLAG'
            ],
            'special_values': {
                'EAI_INTERFACE_DATE': 'SYSDATE',
                'APPLICATION_TRANSFER_FLAG': "'N'"
            }
        }
    }

    def __init__(self):
        self.select_queries = []
        self.insert_queries = []

    def normalize_query(self, query):
        """
        Normalize a SQL query by removing extra whitespace and standardizing format
        
        Args:
            query (str): SQL query to normalize
            
        Returns:
            str: Normalized query
        """
        print(f"Original query: {query}")
        
        # Remove comments if any
        query = re.sub(r'--.*$', '', query, flags=re.MULTILINE)
        
        # Replace newlines with spaces
        query = re.sub(r'\n', ' ', query)
        
        # Replace multiple whitespace with single space
        query = re.sub(r'\s+', ' ', query)
        
        # 핵심 SQL 키워드 주변에 공백 추가 (대소문자 구분 없이)
        keywords = ['SELECT', 'FROM', 'WHERE', 'ORDER BY', 'GROUP BY', 'HAVING', 
                   'JOIN', 'LEFT', 'RIGHT', 'INNER', 'OUTER', 'ON', 'AS']
        
        # 각 키워드를 공백으로 둘러싸기 (단어 경계 고려)
        for keyword in keywords:
            # \b는 단어 경계를 의미
            pattern = re.compile(r'\b' + keyword + r'\b', re.IGNORECASE)
            # 각 키워드를 찾아서 앞뒤에 공백 추가
            query = pattern.sub(' ' + keyword + ' ', query)
        
        # 다시 중복 공백 제거
        query = re.sub(r'\s+', ' ', query)
        
        result = query.strip()
        print(f"Normalized query: {result}")
        return result

    def parse_select_columns(self, query) -> Optional[Dict[str, str]]:
        """Extract columns from SELECT query and return as dictionary"""
        # 대소문자 구분 없이 정규화
        print(f"Parsing query: {query}")
        
        # SELECT 키워드 위치 찾기
        select_match = re.search(r'\bSELECT\b', query, re.IGNORECASE)
        from_match = re.search(r'\bFROM\b', query, re.IGNORECASE)
        
        if not select_match or not from_match:
            print(f"Could not find SELECT or FROM keywords in query: {query}")
            return None
        
        # SELECT와 FROM 사이의 부분 추출
        select_pos = select_match.end()
        from_pos = from_match.start()
        
        if select_pos >= from_pos:
            print(f"Invalid query structure (SELECT appears after FROM): {query}")
            return None
        
        # 컬럼 부분 추출
        column_part = query[select_pos:from_pos].strip()
        print(f"Extracted column part: {column_part}")
        
        # 컬럼 분리 및 처리
        columns = {}
        for col in self._parse_csv_with_functions(column_part):
            col = col.strip()
            if not col:
                continue
            
            print(f"Processing column: {col}")
            
            # to_char 함수 처리 (별칭 유무에 관계없이)
            if 'to_char(' in col.lower():
                # 함수 호출 이후에 별칭이 있는지 확인
                alias_match = re.search(r'(to_char\s*\([^)]+\))\s+([a-zA-Z0-9_]+)$', col, re.IGNORECASE)
                if alias_match:
                    # 별칭이 있는 경우
                    expr, alias = alias_match.groups()
                    print(f"Found to_char with alias: {expr} -> {alias}")
                    columns[expr.strip()] = {'expr': expr.strip(), 'alias': alias.strip(), 'full': col}
                else:
                    # 별칭이 없는 경우
                    print(f"Found to_char without alias: {col}")
                    columns[col] = {'expr': col, 'alias': None, 'full': col}
            else:
                # 일반 열 처리
                alias_match = re.search(r'(.+?)\s+(?:AS\s+)?([a-zA-Z0-9_]+)$', col, re.IGNORECASE)
                if alias_match:
                    expr, alias = alias_match.groups()
                    print(f"Found column with alias: {expr} -> {alias}")
                    columns[expr.strip()] = {'expr': expr.strip(), 'alias': alias.strip(), 'full': col}
                else:
                    print(f"Found column without alias: {col}")
                    columns[col] = {'expr': col, 'alias': None, 'full': col}
        
        print(f"Final parsed columns: {columns}")
        return columns if columns else None

    def _extract_values_with_balanced_parentheses(self, query, start_idx):
        """
        INSERT 쿼리에서 VALUES 절의 내용을 괄호 균형을 맞추며 추출
        
        Args:
            query (str): 전체 쿼리 문자열
            start_idx (int): VALUES 키워드 이후의 시작 인덱스
            
        Returns:
            str: 추출된 VALUES 절 내용 (괄호 포함)
        """
        paren_count = 0
        in_quotes = False
        quote_char = None
        idx = start_idx
        
        while idx < len(query):
            char = query[idx]
            
            # 따옴표 처리 ('나 " 내부에서는 괄호를 계산하지 않음)
            if char in ["'", '"'] and (idx == 0 or query[idx-1] != '\\'):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
            
            # 괄호 카운팅 (따옴표 밖에서만)
            if not in_quotes:
                if char == '(':
                    paren_count += 1
                    if paren_count == 1 and idx == start_idx:  # 시작 괄호
                        start_idx = idx
                elif char == ')':
                    paren_count -= 1
                    if paren_count == 0:  # 종료 괄호 도달
                        return query[start_idx:idx+1]
            
            idx += 1
        
        # 괄호가 맞지 않는 경우
        return None

    def _parse_csv_with_functions(self, csv_string):
        """
        함수 호출과 따옴표를 고려하여 CSV 문자열을 파싱합니다.
        
        Args:
            csv_string (str): 파싱할 CSV 문자열
            
        Returns:
            List[str]: 파싱된 값 목록
        """
        results = []
        current = ""
        paren_count = 0
        in_quotes = False
        quote_char = None
        
        for i, char in enumerate(csv_string):
            # 따옴표 처리
            if char in ["'", '"'] and (i == 0 or csv_string[i-1] != '\\'):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
            
            # 괄호 카운팅 (따옴표 안이 아닐 때)
            if not in_quotes:
                if char == '(':
                    paren_count += 1
                elif char == ')':
                    paren_count -= 1
                    
                # 값 구분자 처리
                if char == ',' and paren_count == 0:
                    results.append(current.strip())
                    current = ""
                    continue
            
            # 현재 문자 추가
            current += char
        
        # 마지막 값 추가
        if current:
            results.append(current.strip())
            
        return results

    def parse_insert_parts(self, query) -> Optional[Tuple[str, Dict[str, str]]]:
        """Extract and return table name and column-value pairs from INSERT query"""
        try:
            # 정규화된 쿼리 사용
            query = self.normalize_query(query)
            print(f"\nProcessing INSERT query:\n{query}")
            
            # INSERT INTO와 테이블 이름 추출
            table_match = re.search(r'INSERT\s+INTO\s+([A-Za-z0-9_$.]+)', query, flags=re.IGNORECASE)
            if not table_match:
                print("Failed to match INSERT INTO pattern")
                return None
                
            table_name = table_match.group(1)
            print(f"Found table name: {table_name}")
            
            # 컬럼 목록 추출
            columns_match = re.search(r'INSERT\s+INTO\s+[A-Za-z0-9_$.]+\s*\((.*?)\)', query, flags=re.IGNORECASE | re.DOTALL)
            if not columns_match:
                print("Failed to match columns pattern")
                return None
                
            col_names = [c.strip() for c in columns_match.group(1).split(',')]
            
            # VALUES 키워드 찾기
            values_match = re.search(r'VALUES\s*\(', query, flags=re.IGNORECASE)
            if not values_match:
                print("Failed to find VALUES keyword")
                return None
                
            # VALUES 절 추출 (괄호 균형 맞추며)
            values_start_idx = values_match.end() - 1  # '(' 위치
            values_part = self._extract_values_with_balanced_parentheses(query, values_start_idx)
            
            if not values_part:
                print("Failed to extract balanced VALUES part")
                return None
                
            # 괄호 제거하고 값만 추출
            values_str = values_part[1:-1]  # 시작과 끝 괄호 제거
            
            # 값 파싱 - 함수 호출을 고려한 파싱
            col_values = self._parse_csv_with_functions(values_str)
            
            print(f"Found columns: {col_names}")
            print(f"Found values: {col_values}")
            
            # 컬럼과 값의 개수가 일치하는지 확인
            if len(col_names) != len(col_values):
                print(f"Column count ({len(col_names)}) does not match value count ({len(col_values)})")
                return None
                
            # 빈 컬럼이나 값이 있는지 확인
            if not all(col_names) or not all(col_values):
                print("Found empty column names or values")
                return None
                
            columns = {}
            for name, value in zip(col_names, col_values):
                columns[name] = value
                
            print(f"Successfully parsed {len(columns)} columns")
            return (table_name, columns)
        except Exception as e:
            print(f"Error parsing INSERT parts: {str(e)}")
            return None

    def compare_queries(self, query1: str, query2: str) -> QueryDifference:
        """
        Compare two SQL queries and return detailed differences
        
        Args:
            query1 (str): First SQL query
            query2 (str): Second SQL query
            
        Returns:
            QueryDifference: Object containing comparison results and differences
        """
        result = QueryDifference()
        
        # 쿼리 정규화
        norm_query1 = self.normalize_query(query1)
        norm_query2 = self.normalize_query(query2)
        
        # 쿼리 타입 확인
        if re.search(r'SELECT', norm_query1, flags=re.IGNORECASE):
            result.query_type = 'SELECT'
            columns1 = self.parse_select_columns(query1)
            columns2 = self.parse_select_columns(query2)
            table1 = self.extract_table_name(query1)
            table2 = self.extract_table_name(query2)
            
            if columns1 is None or columns2 is None:
                raise ValueError("SELECT 쿼리 파싱 실패")
                
        elif re.search(r'INSERT', norm_query1, flags=re.IGNORECASE):
            result.query_type = 'INSERT'
            insert_result1 = self.parse_insert_parts(query1)
            insert_result2 = self.parse_insert_parts(query2)
            
            if insert_result1 is None or insert_result2 is None:
                raise ValueError("INSERT 쿼리 파싱 실패")
                
            table1, columns1 = insert_result1
            table2, columns2 = insert_result2
        else:
            raise ValueError("지원하지 않는 쿼리 타입입니다.")
            
        result.table_name = table1
        
        # 특수 컬럼 제외
        direction = 'recv' if result.query_type == 'INSERT' else 'send'
        special_cols = set(self.special_columns[direction]['required'])
        
        # 정규화된 비교를 위해 컬럼 처리
        if result.query_type == 'SELECT':
            # 결과가 동일한지 계산
            result.is_equal = self._compare_select_columns(columns1, columns2, special_cols, result)
        else:
            # 일반 컬럼만 비교 (대소문자 구분 없이 비교하되 원본 케이스 유지)
            columns1_filtered = {k: v for k, v in columns1.items() if k.upper() not in special_cols}
            columns2_filtered = {k: v for k, v in columns2.items() if k.upper() not in special_cols}
            
            # 컬럼 비교
            all_columns = set(columns1_filtered.keys()) | set(columns2_filtered.keys())
            is_equal = True
            for col in all_columns:
                if col not in columns1_filtered:
                    result.add_difference(col, None, columns2_filtered[col])
                    is_equal = False
                elif col not in columns2_filtered:
                    result.add_difference(col, columns1_filtered[col], None)
                    is_equal = False
                else:
                    # 값 비교 시 to_char 함수의 포맷 차이를 무시
                    val1 = columns1_filtered[col]
                    val2 = columns2_filtered[col]
                    
                    # to_char 또는 to_date 함수를 포함하는 값이면 정규화 적용
                    if 'to_char(' in val1.lower() or 'to_char(' in val2.lower() or 'to_date(' in val1.lower() or 'to_date(' in val2.lower():
                        norm_val1 = self._normalize_tochar_format(val1)
                        norm_val2 = self._normalize_tochar_format(val2)
                        if norm_val1 != norm_val2:
                            result.add_difference(col, val1, val2)
                            is_equal = False
                    else:
                        # 일반 값은 그대로 비교
                        if val1 != val2:
                            result.add_difference(col, val1, val2)
                            is_equal = False
            
            result.is_equal = is_equal
                
        return result
    
    def _compare_select_columns(self, columns1, columns2, special_cols, result):
        """
        SELECT 쿼리의 컬럼을 비교하는 보조 메소드
        
        Args:
            columns1: 첫 번째 쿼리의 컬럼 정보
            columns2: 두 번째 쿼리의 컬럼 정보
            special_cols: 특수 컬럼 집합
            result: 결과를 저장할 QueryDifference 객체
            
        Returns:
            bool: 두 쿼리의 컬럼이 동일한지 여부
        """
        # 일반 컬럼만 필터링 (특수 컬럼 제외)
        columns1_filtered = {k: v for k, v in columns1.items() 
                            if k.upper() not in special_cols and 
                              (v['alias'] is None or v['alias'].upper() not in special_cols)}
        columns2_filtered = {k: v for k, v in columns2.items() 
                            if k.upper() not in special_cols and 
                              (v['alias'] is None or v['alias'].upper() not in special_cols)}
        
        # 두 쿼리의 모든 컬럼 표현식 목록 생성
        expr1_set = {info['expr'].strip() for info in columns1_filtered.values()}
        expr2_set = {info['expr'].strip() for info in columns2_filtered.values()}
        
        # 정규화된 표현식 매핑 생성 - 공백 차이 등을 무시
        norm_expr1_map = {}
        for info in columns1_filtered.values():
            expr = info['expr'].strip()
            # to_char 함수의 포맷 문자열을 정규화 (포맷 차이 무시)
            norm_expr = self._normalize_tochar_format(expr)
            norm_expr1_map[norm_expr] = expr
            
        norm_expr2_map = {}
        for info in columns2_filtered.values():
            expr = info['expr'].strip()
            # to_char 함수의 포맷 문자열을 정규화 (포맷 차이 무시)
            norm_expr = self._normalize_tochar_format(expr)
            norm_expr2_map[norm_expr] = expr
            
        # 정규화된 표현식 세트
        norm_expr1_set = set(norm_expr1_map.keys())
        norm_expr2_set = set(norm_expr2_map.keys())
        
        # 정규화된 표현식으로 비교 (별칭과 공백 차이 무시)
        only_in_query1 = norm_expr1_set - norm_expr2_set
        only_in_query2 = norm_expr2_set - norm_expr1_set
        
        is_equal = True
        
        # 첫 번째 쿼리에만 있는 표현식 처리
        for norm_expr in only_in_query1:
            orig_expr = norm_expr1_map[norm_expr]
            # 이 표현식을 포함하는 컬럼 정보 찾기
            for col, info in columns1_filtered.items():
                norm_col_expr = self._normalize_tochar_format(info['expr'].strip())
                if norm_col_expr == norm_expr:
                    result.add_difference(col, info['full'], None)
                    is_equal = False
                    break
        
        # 두 번째 쿼리에만 있는 표현식 처리
        for norm_expr in only_in_query2:
            orig_expr = norm_expr2_map[norm_expr]
            # 이 표현식을 포함하는 컬럼 정보 찾기
            for col, info in columns2_filtered.items():
                norm_col_expr = self._normalize_tochar_format(info['expr'].strip())
                if norm_col_expr == norm_expr:
                    result.add_difference(col, None, info['full'])
                    is_equal = False
                    break
        
        return is_equal
    
    def _normalize_tochar_format(self, expr):
        """
        to_char 함수와 to_date 함수의 포맷 문자열을 정규화합니다.
        포맷 문자열의 차이를 무시하고 함수와 인자 패턴만 비교합니다.
        
        Args:
            expr (str): SQL 표현식
            
        Returns:
            str: 정규화된 표현식
        """
        # 기본 공백 정규화
        norm_expr = re.sub(r'\s+', ' ', expr).strip().lower()
        
        # 함수 내부의 공백 정규화 (특히 콤마와 따옴표 사이의 공백)
        # 콤마 다음 공백 정규화
        norm_expr = re.sub(r',\s+', ',', norm_expr)
        # 콤마 이전 공백 정규화
        norm_expr = re.sub(r'\s+,', ',', norm_expr)
        # 괄호와 인자 사이의 공백 정규화
        norm_expr = re.sub(r'\(\s+', '(', norm_expr)
        norm_expr = re.sub(r'\s+\)', ')', norm_expr)
        
        # to_char 함수의 포맷 부분 정규화
        # to_char(column, 'FORMAT') 패턴에서 'FORMAT' 부분을 일반화
        to_char_pattern = r"""
            (to_char\s*\(\s*[^,]+\s*,\s*)  # 함수 이름과 첫 인자
            (?:\'[^\']*\'|\"[^\"]*\")      # 포맷 문자열
            (\s*\))                        # 닫는 괄호
        """
        to_char_pattern = re.compile(to_char_pattern, flags=re.IGNORECASE | re.VERBOSE)
        
        # to_date 함수의 포맷 부분 정규화
        # to_date(column, 'FORMAT') 패턴에서 'FORMAT' 부분을 일반화
        to_date_pattern = r"""
            (to_date\s*\(\s*[^,]+\s*,\s*)  # 함수 이름과 첫 인자
            (?:\'[^\']*\'|\"[^\"]*\")      # 포맷 문자열
            (\s*\))                        # 닫는 괄호
        """
        to_date_pattern = re.compile(to_date_pattern, flags=re.IGNORECASE | re.VERBOSE)
        
        # to_char와 to_date 함수의 포맷 문자열을 'FORMAT'으로 일반화
        norm_expr = to_char_pattern.sub(r'\1\'FORMAT\'\2', norm_expr)
        norm_expr = to_date_pattern.sub(r'\1\'FORMAT\'\2', norm_expr)
        
        return norm_expr

    def check_special_columns(self, query: str, direction: str) -> List[str]:
        """
        특수 컬럼의 존재 여부와 값을 체크합니다.
        
        Args:
            query (str): 검사할 쿼리
            direction (str): 송신('send') 또는 수신('recv')
            
        Returns:
            List[str]: 경고 메시지 리스트
        """
        warnings = []
        
        if direction == 'send':
            columns = self.parse_select_columns(query)
        else:
            _, columns = self.parse_insert_parts(query)
            
        if not columns:
            return warnings
            
        # 대소문자 구분 없이 컬럼 비교를 위한 매핑 생성
        if direction == 'send':
            # SELECT 쿼리의 경우 새 구조에 맞게 처리
            columns_upper = {}
            for k, v in columns.items():
                # 별칭이 있는 경우 별칭을 키로 사용
                if v['alias'] is not None:
                    columns_upper[v['alias'].upper()] = (v['alias'], v)
                else:
                    # 별칭이 없는 경우 표현식을 키로 사용
                    columns_upper[k.upper()] = (k, v)
        else:
            # INSERT 쿼리는 기존과 동일하게 처리
            columns_upper = {k.upper(): (k, v) for k, v in columns.items()}
        
        # 필수 특수 컬럼 체크
        for col in self.special_columns[direction]['required']:
            if col not in columns_upper:
                warnings.append(f"필수 특수 컬럼 '{col}'이(가) {direction} 쿼리에 없습니다.")
        
        # 수신 쿼리의 특수 값 체크
        if direction == 'recv':
            for col, expected_value in self.special_columns[direction]['special_values'].items():
                if col in columns_upper:
                    col_name, col_value = columns_upper[col]
                    if col_value != expected_value:
                        warnings.append(f"특수 컬럼 '{col}'의 값이 기대값과 다릅니다. 기대값: {expected_value}, 실제값: {col_value}")
                        
        return warnings

    def clean_select_query(self, query):
        """
        Clean SELECT query by removing WHERE clause
        """
        # Find the position of WHERE (case insensitive)
        where_match = re.search(r'\sWHERE\s', query, flags=re.IGNORECASE)
        if where_match:
            # Return only the part before WHERE
            return query[:where_match.start()].strip()
        return query.strip()

    def clean_insert_query(self, query: str) -> str:
        """
        Clean INSERT query by removing PL/SQL blocks
        """
        # PL/SQL 블록에서 INSERT 문 추출
        pattern = r"""
            (?:BEGIN\s+)?          # BEGIN (optional)
            (INSERT\s+INTO\s+      # INSERT INTO
            [^;]+                  # everything until semicolon
            )                      # capture this part
            (?:\s*;)?             # optional semicolon
            (?:\s*EXCEPTION\s+     # EXCEPTION block (optional)
            .*?                    # everything until END
            END;?)?                # END with optional semicolon
        """
        insert_match = re.search(
            pattern,
            query,
            flags=re.IGNORECASE | re.MULTILINE | re.DOTALL | re.VERBOSE
        )
        
        if insert_match:
            return insert_match.group(1).strip()
        return query.strip()

    def is_meaningful_query(self, query: str) -> bool:
        """
        Check if a query is meaningful (not just a simple existence check or count)
        
        Args:
            query (str): SQL query to analyze
            
        Returns:
            bool: True if the query is meaningful, False otherwise
        """
        query = query.lower()
        
        # Remove comments and normalize whitespace
        query = re.sub(r'--.*$', '', query, flags=re.MULTILINE)
        query = ' '.join(query.split())
        
        # Patterns for meaningless queries
        meaningless_patterns = [
            r'select\s+1\s+from',  # SELECT 1 FROM ...
            r'select\s+count\s*\(\s*\*\s*\)\s+from',  # SELECT COUNT(*) FROM ...
            r'select\s+count\s*\(\s*1\s*\)\s+from',  # SELECT COUNT(1) FROM ...
            r'select\s+null\s+from',  # SELECT NULL FROM ...
            r'select\s+\'[^\']*\'\s+from',  # SELECT 'constant' FROM ...
            r'select\s+\d+\s+from',  # SELECT {number} FROM ...
        ]
        
        # Check if query matches any meaningless pattern
        for pattern in meaningless_patterns:
            if re.search(pattern, query):
                return False
                
        # For SELECT queries, check if it's selecting actual columns
        if query.startswith('select'):
            # Extract the SELECT clause (between SELECT and FROM)
            select_match = re.match(r'select\s+(.+?)\s+from', query)
            if select_match:
                select_clause = select_match.group(1)
                # If only selecting literals or simple expressions, consider it meaningless
                if re.match(r'^[\d\'\"\s,]+$', select_clause):
                    return False
        
        return True

    def find_files_by_table(self, folder_path: str, table_name: str, skip_meaningless: bool = True) -> dict:
        """
        Find files containing queries that reference the specified table
        
        Args:
            folder_path (str): Path to the folder to search in
            table_name (str): Name of the DB table to search for
            skip_meaningless (bool): If True, skip queries that appear to be meaningless
            
        Returns:
            dict: Dictionary with 'select' and 'insert' as keys, each containing a list of tuples
                 where each tuple contains (file_path, query)
        """
        import os
        
        results = {
            'select': [],
            'insert': []
        }
        
        # Normalize table name for comparison
        table_name = table_name.lower()
        
        # Create parser instance for processing files
        parser = self
        
        # Walk through all files in the folder
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                
                # Skip non-XML files silently
                if not file_path.lower().endswith('.xml'):
                    continue
                    
                try:
                    # Try to parse queries from the file
                    select_queries, insert_queries = self.parse_xml_file(file_path)
                    
                    # Check SELECT queries
                    for query in select_queries:
                        if self.extract_table_name(query).lower() == table_name:
                            # Skip meaningless queries if requested
                            if skip_meaningless and not self.is_meaningful_query(query):
                                continue
                            rel_path = os.path.relpath(file_path, folder_path)
                            results['select'].append((rel_path, query))
                    
                    # Check INSERT queries
                    for query in insert_queries:
                        if self.extract_table_name(query).lower() == table_name:
                            rel_path = os.path.relpath(file_path, folder_path)
                            results['insert'].append((rel_path, query))
                    
                except Exception:
                    # Skip any errors silently
                    continue
        
        return results

    def parse_xml_file(self, filename):
        """
        Parse XML file and extract SQL queries
        
        Args:
            filename (str): Path to the XML file
            
        Returns:
            tuple: Lists of (select_queries, insert_queries)
        """
        try:
            # Clear previous queries
            self.select_queries = []
            self.insert_queries = []
            
            # Parse XML file
            tree = ET.parse(filename)
            root = tree.getroot()
            
            # Find all text content in the XML
            for elem in root.iter():
                if elem.text:
                    text = elem.text.strip()
                    # Extract SELECT queries
                    if re.search(r'SELECT', text, flags=re.IGNORECASE):
                        cleaned_query = self.clean_select_query(text)
                        self.select_queries.append(cleaned_query)
                    # Extract INSERT queries
                    elif re.search(r'INSERT', text, flags=re.IGNORECASE):
                        cleaned_query = self.clean_insert_query(text)
                        self.insert_queries.append(cleaned_query)
            
            return self.select_queries, self.insert_queries
            
        except ET.ParseError:
            return [], []
        except Exception:
            return [], []
    
    def get_select_queries(self):
        """Return list of extracted SELECT queries"""
        return self.select_queries
    
    def get_insert_queries(self):
        """Return list of extracted INSERT queries"""
        return self.insert_queries
    
    def print_queries(self):
        """Print all extracted queries"""
        print("\nSELECT Queries:")
        for i, query in enumerate(self.select_queries, 1):
            print(f"{i}. {query}\n")
            
        print("\nINSERT Queries:")
        for i, query in enumerate(self.insert_queries, 1):
            print(f"{i}. {query}\n")

    def print_query_differences(self, diff: QueryDifference):
        """Print the differences between two queries in a readable format"""
        print(f"\nQuery Type: {diff.query_type}")
        if diff.is_equal:
            print("Queries are equivalent")
        else:
            print("Differences found:")
            for d in diff.differences:
                print(f"- Column '{d['column']}':")
                print(f"  Query 1: {d['query1_value']}")
                print(f"  Query 2: {d['query2_value']}")

    def extract_table_name(self, query: str) -> str:
        """
        Extract table name from a SQL query
        
        Args:
            query (str): SQL query to analyze
            
        Returns:
            str: Table name or empty string if not found
        """
        query = self.normalize_query(query)
        
        # For SELECT queries
        select_match = re.search(r'from\s+([a-zA-Z0-9_$.]+)', query, flags=re.IGNORECASE)
        if select_match:
            return select_match.group(1)
            
        # For INSERT queries
        insert_match = re.search(r'insert\s+into\s+([a-zA-Z0-9_$.]+)', query, flags=re.IGNORECASE)
        if insert_match:
            return insert_match.group(1)
            
        return ""

    def print_table_search_results(self, results: dict, table_name: str):
        """
        Print table search results in a formatted way
        
        Args:
            results (dict): Dictionary with search results
            table_name (str): Name of the DB table that was searched
        """
        print(f"\nFiles and queries referencing table: {table_name}")
        print("=" * 50)
        
        print("\nSELECT queries found in:")
        if results['select']:
            for i, (file, query) in enumerate(results['select'], 1):
                print(f"\n{i}. File: {file}")
                print("Query:")
                print(query)
        else:
            print("  No files found with SELECT queries")
            
        print("\nINSERT queries found in:")
        if results['insert']:
            for i, (file, query) in enumerate(results['insert'], 1):
                print(f"\n{i}. File: {file}")
                print("Query:")
                print(query)
        else:
            print("  No files found with INSERT queries")
        
        print("\n" + "=" * 50)

    def compare_mq_bw_queries(self, mq_xml_path: str, bw_xml_path: str) -> Dict[str, List[QueryDifference]]:
        """
        MQ XML과 BW XML 파일에서 추출한 송신/수신 쿼리를 비교합니다.
        
        Args:
            mq_xml_path (str): MQ XML 파일 경로
            bw_xml_path (str): BW XML 파일 경로
            
        Returns:
            Dict[str, List[QueryDifference]]: 송신/수신별 쿼리 비교 결과
                {
                    'send': [송신 쿼리 비교 결과 목록],
                    'recv': [수신 쿼리 비교 결과 목록]
                }
        """
        results = {
            'send': [],
            'recv': []
        }
        
        # MQ XML 파싱
        mq_queries = self.parse_xml_file(mq_xml_path)
        if not mq_queries:
            print(f"Failed to parse MQ XML file: {mq_xml_path}")
            return results
            
        mq_select_queries, mq_insert_queries = mq_queries
        
        # BW XML 파싱
        bw_extractor = BWQueryExtractor()
        bw_queries = bw_extractor.extract_bw_queries(bw_xml_path)
        if not bw_queries:
            print(f"Failed to parse BW XML file: {bw_xml_path}")
            return results
        
        bw_send_queries = bw_queries.get('send', [])
        bw_recv_queries = bw_queries.get('recv', [])
        
        # 송신 쿼리 비교 (SELECT)
        if mq_select_queries and bw_send_queries:
            print("\n===== 송신 쿼리 비교 (SELECT) =====")
            for mq_query in mq_select_queries:
                for bw_query in bw_send_queries:
                    diff = self.compare_queries(mq_query, bw_query)
                    if diff:
                        results['send'].append(diff)
                        self.print_query_differences(diff)
        else:
            print("송신 쿼리 비교를 위한 데이터가 부족합니다.")
            
        # 수신 쿼리 비교 (INSERT)
        if mq_insert_queries and bw_recv_queries:
            print("\n===== 수신 쿼리 비교 (INSERT) =====")
            for mq_query in mq_insert_queries:
                for bw_query in bw_recv_queries:
                    diff = self.compare_queries(mq_query, bw_query)
                    if diff:
                        results['recv'].append(diff)
                        self.print_query_differences(diff)
        else:
            print("수신 쿼리 비교를 위한 데이터가 부족합니다.")
            
        return results

    def compare_mq_bw_queries_by_interface_id(self, interface_id: str, mq_folder_path: str, bw_folder_path: str) -> Dict[str, List[QueryDifference]]:
        """
        인터페이스 ID를 기준으로 MQ XML과 BW XML 파일을 찾아 쿼리를 비교합니다.
        
        Args:
            interface_id (str): 인터페이스 ID
            mq_folder_path (str): MQ XML 파일이 있는 폴더 경로
            bw_folder_path (str): BW XML 파일이 있는 폴더 경로
            
        Returns:
            Dict[str, List[QueryDifference]]: 송신/수신별 쿼리 비교 결과
        """
        results = {
            'send': [],
            'recv': []
        }
        
        # MQ XML 파일 찾기
        searcher = FileSearcher()
        mq_files = searcher.find_files_with_keywords(mq_folder_path, [interface_id])
        
        if not mq_files or not mq_files.get(interface_id):
            print(f"No MQ XML files found for interface ID: {interface_id}")
            return results
            
        mq_xml_path = mq_files[interface_id][0] if mq_files[interface_id] else None
        if not mq_xml_path:
            print(f"No MQ XML file found for interface ID: {interface_id}")
            return results
            
        # MQ XML에서 테이블 이름 추출
        mq_queries = self.parse_xml_file(mq_xml_path)
        if not mq_queries or not mq_queries[0]:
            print(f"Failed to parse MQ XML file: {mq_xml_path}")
            return results
            
        table_name = self.extract_table_name(mq_queries[0][0]) if mq_queries[0] else None
        if not table_name:
            print(f"Failed to extract table name from MQ XML SELECT query")
            return results
            
        print(f"Found table name from MQ XML: {table_name}")
        
        # BW XML 파일 찾기
        bw_files = searcher.find_files_with_keywords(bw_folder_path, [table_name])
        
        if not bw_files or not bw_files.get(table_name):
            print(f"No BW XML files found for table name: {table_name}")
            return results
            
        bw_xml_paths = bw_files[table_name] if bw_files.get(table_name) else []
        if not bw_xml_paths:
            print(f"No BW XML files found for table name: {table_name}")
            return results
            
        # 각 BW XML 파일과 비교
        all_results = {
            'send': [],
            'recv': []
        }
        
        for bw_xml_path in bw_xml_paths:
            print(f"\nComparing MQ XML ({mq_xml_path}) with BW XML ({bw_xml_path}):")
            curr_results = self.compare_mq_bw_queries(mq_xml_path, bw_xml_path)
            
            if curr_results['send']:
                all_results['send'].extend(curr_results['send'])
                
            if curr_results['recv']:
                all_results['recv'].extend(curr_results['recv'])
                
        return all_results

class BWQueryExtractor:
    """TIBCO BW XML 파일에서 특정 태그 구조에 따라 SQL 쿼리를 추출하는 클래스"""
    
    def __init__(self):
        self.ns = {
            'pd': 'http://xmlns.tibco.com/bw/process/2003',
            'xsl': 'http://www.w3.org/1999/XSL/Transform',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
        }

    def _remove_oracle_hints(self, query: str) -> str:
        """
        SQL 쿼리에서 Oracle 힌트(/*+ ... */) 제거
        
        Args:
            query (str): 원본 SQL 쿼리
            
        Returns:
            str: 힌트가 제거된 SQL 쿼리
        """
        import re
        # /*+ ... */ 패턴의 힌트 제거
        cleaned_query = re.sub(r'/\*\+[^*]*\*/', '', query)
        # 불필요한 공백 정리 (여러 개의 공백을 하나로)
        cleaned_query = re.sub(r'\s+', ' ', cleaned_query).strip()
        
        if cleaned_query != query:
            print("\n=== Oracle 힌트 제거 ===")
            print(f"원본 쿼리: {query}")
            print(f"정리된 쿼리: {cleaned_query}")
            
        return cleaned_query

    def _get_parameter_names(self, activity) -> List[str]:
        """
        Prepared_Param_DataType에서 파라미터 이름 목록 추출
        
        Args:
            activity: JDBC 액티비티 XML 요소
            
        Returns:
            List[str]: 파라미터 이름 목록
        """
        param_names = []
        print("\n=== XML 구조 디버깅 ===")
        print("activity 태그:", activity.tag)
        print("activity의 자식 태그들:", [child.tag for child in activity])
        
        # 대소문자를 맞춰서 수정
        prepared_params = activity.find('.//Prepared_Param_DataType', self.ns)
        if prepared_params is not None:
            print("\n=== Prepared_Param_DataType 태그 발견 ===")
            print("prepared_params 태그:", prepared_params.tag)
            print("prepared_params의 자식 태그들:", [child.tag for child in prepared_params])
            
            for param in prepared_params.findall('./parameter', self.ns):
                param_name = param.find('./parameterName', self.ns)
                if param_name is not None and param_name.text:
                    name = param_name.text.strip()
                    param_names.append(name)
                    print(f"파라미터 이름 추출: {name}")
        else:
            print("\n=== Prepared_Param_DataType 태그를 찾을 수 없음 ===")
            # 전체 XML 구조를 재귀적으로 출력하여 디버깅
            def print_element_tree(element, level=0):
                print("  " * level + f"- {element.tag}")
                for child in element:
                    print_element_tree(child, level + 1)
            print("\n=== 전체 XML 구조 ===")
            print_element_tree(activity)
        
        return param_names

    def _replace_with_param_names(self, query: str, param_names: List[str]) -> str:
        """
        1단계: SQL 쿼리의 ? 플레이스홀더를 prepared_Param_DataType의 파라미터 이름으로 대체
        
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
            if i < len(parts):
                result += f":{param_name}" + parts[i+1]
                
        print("\n=== 1단계: prepared_Param_DataType 매핑 결과 ===")
        print(f"원본 쿼리: {query}")
        print(f"매핑된 쿼리: {result}")
        return result

    def _get_record_mappings(self, activity, param_names: List[str]) -> Dict[str, str]:
        """
        2단계: Record 태그에서 실제 값 매핑 정보 추출
        
        Args:
            activity: JDBC 액티비티 XML 요소
            param_names: prepared_Param_DataType에서 추출한 파라미터 이름 목록
            
        Returns:
            Dict[str, str]: 파라미터 이름과 매핑된 실제 값의 딕셔너리
        """
        mappings = {}
        # 이미 매핑된 실제 컬럼 값을 추적하는 집합
        mapped_values = set()
        
        input_bindings = activity.find('.//pd:inputBindings', self.ns)
        if input_bindings is None:
            print("\n=== inputBindings 태그를 찾을 수 없음 ===")
            return mappings

        print("\n=== Record 매핑 검색 시작 ===")
        
        # jdbcUpdateActivityInput/Record 찾기
        jdbc_input = input_bindings.find('.//jdbcUpdateActivityInput')
        if jdbc_input is None:
            print("jdbcUpdateActivityInput을 찾을 수 없음")
            return mappings

        # for-each/Record 찾기
        for_each = jdbc_input.find('.//xsl:for-each', self.ns)
        record = for_each.find('./Record') if for_each is not None else jdbc_input
        
        if record is not None:
            print("Record 태그 발견")
            # 각 파라미터 이름에 대해 매핑 찾기
            for param_name in param_names:
                print(f"\n파라미터 '{param_name}' 매핑 검색:")
                param_element = record.find(f'.//{param_name}')
                if param_element is not None:
                    # 매핑 타입별로 값을 추출하되, 중복 매핑을 방지
                    mapping_found = False
                    
                    # value-of 체크 (우선 순위 1)
                    value_of = param_element.find('.//xsl:value-of', self.ns)
                    if value_of is not None:
                        select_attr = value_of.get('select', '')
                        if select_attr:
                            # select="BANANA"와 같은 형식에서 실제 값 추출
                            value = select_attr.split('/')[-1]
                            mappings[param_name] = value
                            print(f"value-of 매핑 발견: {param_name} -> {value}")
                            mapping_found = True
                    
                    # choose/when 체크 (우선 순위 2, value-of가 없을 경우만)
                    if not mapping_found:
                        choose = param_element.find('.//xsl:choose', self.ns)
                        if choose is not None:
                            when = choose.find('.//xsl:when', self.ns)
                            if when is not None:
                                test_attr = when.get('test', '')
                                if 'exists(' in test_attr:
                                    # exists(BANANA)와 같은 형식에서 변수 이름 추출
                                    value = test_attr[test_attr.find('(')+1:test_attr.find(')')]
                                    mappings[param_name] = value
                                    print(f"choose/when 매핑 발견: {param_name} -> {value}")
                else:
                    print(f"'{param_name}'에 대한 매핑을 찾을 수 없음")

        return mappings

    def _replace_with_actual_values(self, query: str, mappings: Dict[str, str]) -> str:
        """
        2단계: 파라미터 이름을 Record에서 찾은 실제 값으로 대체
        
        Args:
            query (str): 1단계에서 파라미터 이름이 대체된 쿼리
            mappings (Dict[str, str]): 파라미터 이름과 실제 값의 매핑
            
        Returns:
            str: 실제 값이 대체된 SQL 쿼리
        """
        # 순차적 치환 문제 해결을 위해 모든 대체를 한 번에 수행
        # 1. 대체될 모든 패턴을 고유한 임시 패턴으로 먼저 변환 (충돌 방지)
        result = query
        temp_replacements = {}
        
        import re
        
        for i, (param_name, actual_value) in enumerate(mappings.items()):
            # 고유한 임시 패턴 생성 (절대 원본 쿼리에 존재할 수 없는 패턴)
            temp_pattern = f"__TEMP_PLACEHOLDER_{i}__"
            
            # 정규 표현식을 사용하여 정확한 파라미터 이름만 대체
            # 단어 경계(\b)를 사용하여 정확한 파라미터 이름만 매칭
            result = re.sub(f":{param_name}\\b", temp_pattern, result)
            
            # 임시 패턴을 최종 값으로 매핑
            temp_replacements[temp_pattern] = f":{actual_value}"
        
        # 2. 모든 임시 패턴을 최종 값으로 한 번에 변환
        for temp_pattern, final_value in temp_replacements.items():
            result = result.replace(temp_pattern, final_value)
            
        print("\n=== 2단계: Record 매핑 결과 ===")
        print(f"1단계 쿼리: {query}")
        print(f"최종 쿼리: {result}")
        return result

    def extract_recv_query(self, xml_path: str) -> List[Tuple[str, str, str]]:
        """
        수신용 XML에서 SQL 쿼리와 파라미터가 매핑된 쿼리를 추출
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            List[Tuple[str, str, str]]: (원본 쿼리, 1단계 매핑 쿼리, 2단계 매핑 쿼리) 목록
        """
        queries = []
        try:
            print(f"\n=== XML 파일 처리 시작: {xml_path} ===")
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # JDBC 액티비티 찾기
            activities = root.findall('.//pd:activity', self.ns)
            
            for activity in activities:
                # JDBC 액티비티 타입 확인
                activity_type = activity.find('./pd:type', self.ns)
                if activity_type is None or 'jdbc' not in activity_type.text.lower():
                    continue
                    
                print(f"\nJDBC 액티비티 발견: {activity.get('name', 'Unknown')}")
                
                # statement 추출
                statement = activity.find('.//config/statement')
                if statement is not None and statement.text:
                    query = statement.text.strip()
                    print(f"\n발견된 쿼리:\n{query}")
                    
                    # SELECT 쿼리인 경우
                    if query.lower().startswith('select'):
                        # FROM DUAL 쿼리 제외
                        if not self._is_valid_query(query):
                            print("=> FROM DUAL 쿼리이므로 제외")
                            continue
                        # Oracle 힌트 제거
                        query = self._remove_oracle_hints(query)
                        print(f"=> Oracle 힌트 제거 후 쿼리:\n{query}")
                        queries.append((query, query, query))  # SELECT는 파라미터 매핑 없음
                    
                    # INSERT, UPDATE, DELETE 쿼리인 경우
                    elif query.lower().startswith(('insert', 'update', 'delete')):
                        # 1단계: prepared_Param_DataType의 파라미터 이름으로 매핑
                        param_names = self._get_parameter_names(activity)
                        stage1_query = self._replace_with_param_names(query, param_names)
                        
                        # 2단계: Record의 실제 값으로 매핑
                        mappings = self._get_record_mappings(activity, param_names)
                        stage2_query = self._replace_with_actual_values(stage1_query, mappings)
                        
                        queries.append((query, stage1_query, stage2_query))
                        print(f"=> 최종 처리된 쿼리:\n{stage2_query}")
            
            print(f"\n=== 처리된 유효한 쿼리 수: {len(queries)} ===")
            
        except ET.ParseError as e:
            print(f"\n=== XML 파싱 오류: {e} ===")
        except Exception as e:
            print(f"\n=== 쿼리 추출 중 오류 발생: {e} ===")
            
        return queries

    def _is_valid_query(self, query: str) -> bool:
        """
        분석 대상이 되는 유효한 쿼리인지 확인
        
        Args:
            query (str): SQL 쿼리
            
        Returns:
            bool: 유효한 쿼리이면 True
        """
        # 소문자로 변환하여 검사
        query_lower = query.lower()
        
        # SELECT FROM DUAL 패턴 체크
        if query_lower.startswith('select') and 'from dual' in query_lower:
            print(f"\n=== 단순 쿼리 제외 ===")
            print(f"제외된 쿼리: {query}")
            return False
            
        return True

    def extract_send_query(self, xml_path: str) -> List[str]:
        """
        송신용 XML에서 SQL 쿼리 추출
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            List[str]: SQL 쿼리 목록
        """
        queries = []
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # 송신 쿼리 추출 (Group 내의 SelectP 활동)
            select_activities = root.findall('.//pd:group[@name="Group"]//pd:activity[@name="SelectP"]', self.ns)
            
            print(f"\n=== 송신용 XML 처리 시작: {xml_path} ===")
            print(f"발견된 SelectP 활동 수: {len(select_activities)}")
            
            for activity in select_activities:
                statement = activity.find('.//config/statement')
                if statement is not None and statement.text:
                    query = statement.text.strip()
                    print(f"\n발견된 쿼리:\n{query}")
                    
                    # 1. 유효한 쿼리인지 먼저 확인
                    if not self._is_valid_query(query):
                        print("=> FROM DUAL 쿼리이므로 제외")
                        continue
                        
                    # 2. 유효한 쿼리에 대해서만 Oracle 힌트 제거
                    cleaned_query = self._remove_oracle_hints(query)
                    print(f"=> 최종 처리된 쿼리:\n{cleaned_query}")
                    queries.append(cleaned_query)
            
            print(f"\n=== 처리된 유효한 쿼리 수: {len(queries)} ===")
            
        except ET.ParseError as e:
            print(f"XML 파싱 오류: {e}")
        except Exception as e:
            print(f"쿼리 추출 중 오류 발생: {e}")
            
        return queries

    def extract_bw_queries(self, xml_path: str) -> Dict[str, List[str]]:
        """
        TIBCO BW XML 파일에서 송신/수신 쿼리를 모두 추출
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            Dict[str, List[str]]: 송신/수신 쿼리 목록
                {
                    'send': [select 쿼리 목록],
                    'recv': [insert 쿼리 목록]
                }
        """
        # 모든 쿼리를 추출
        send_queries = self.extract_send_query(xml_path)
        recv_queries_full = self.extract_recv_query(xml_path)
        
        # 수신 쿼리 중 INSERT 문만 필터링
        recv_queries = []
        for orig_query, _, mapped_query in recv_queries_full:
            if orig_query.lower().startswith('insert'):
                recv_queries.append(mapped_query)
        
        return {
            'send': [query for query in send_queries if query.lower().startswith('select')],
            'recv': recv_queries
        }
    def get_single_query(self, xml_path: str) -> str:
        """
        BW XML 파일에서 SQL 쿼리를 추출하여 단일 문자열로 반환
        송신(send)과 수신(recv) 쿼리 중 존재하는 것을 반환
        둘 다 없는 경우 빈 문자열 반환
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            str: 추출된 SQL 쿼리 문자열. 쿼리가 없으면 빈 문자열
        """
        try:
            # 기존 extract_bw_queries 메소드 활용
            queries = self.extract_bw_queries(xml_path)
            
            # 송신 쿼리 확인
            if queries.get('send') and len(queries['send']) > 0:
                return queries['send'][0]  # 첫 번째 송신 쿼리 반환
                
            # 수신 쿼리 확인
            if queries.get('recv') and len(queries['recv']) > 0:
                return queries['recv'][0]  # 첫 번째 수신 쿼리 반환
                
            # 쿼리가 없는 경우
            return ""
            
        except Exception as e:
            print(f"쿼리 추출 중 오류 발생: {e}")
            return ""  # 오류 발생 시 빈 문자열 반환        

class FileSearcher:
    @staticmethod
    def find_files_with_keywords(folder_path: str, keywords: list) -> dict:
        """
        Search for files in the given folder that contain any of the specified keywords
        
        Args:
            folder_path (str): Path to the folder to search in
            keywords (list): List of keywords to search for
            
        Returns:
            dict: Dictionary with keyword as key and list of matching files as value
        """
        import os
        
        # Initialize results dictionary
        results = {keyword: [] for keyword in keywords}
        
        # Walk through all files in the folder
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                
                try:
                    # Try to read file content
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                        
                    # Check for each keyword
                    for keyword in keywords:
                        if keyword in content:
                            # Store relative path instead of full path
                            rel_path = os.path.relpath(file_path, folder_path)
                            results[keyword].append(rel_path)
                            
                except (UnicodeDecodeError, IOError):
                    # Skip files that can't be read as text
                    continue
        
        return results

    @staticmethod
    def print_search_results(results: dict):
        """
        Print search results in a formatted way
        
        Args:
            results (dict): Dictionary with keyword as key and list of matching files as value
        """
        print("\nSearch Results:")
        print("=" * 50)
        
        for keyword, files in results.items():
            print(f"\nKeyword: {keyword}")
            if files:
                print("Found in files:")
                for i, file in enumerate(files, 1):
                    print(f"  {i}. {file}")
            else:
                print("No files found containing this keyword")
        
        print("\n" + "=" * 50)

# Test the query comparison
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Compare SQL queries in MQ and BW XML files")
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    
    # 테이블 검색 명령
    table_parser = subparsers.add_parser("find_table", help="Find files containing a specific table")
    table_parser.add_argument("folder_path", help="Folder path to search in")
    table_parser.add_argument("table_name", help="Table name to search for")
    
    # 쿼리 비교 명령
    compare_parser = subparsers.add_parser("compare", help="Compare MQ and BW queries")
    compare_parser.add_argument("mq_xml", help="MQ XML file path")
    compare_parser.add_argument("bw_xml", help="BW XML file path")
    
    # 인터페이스 ID로 비교 명령
    interface_parser = subparsers.add_parser("compare_by_id", help="Compare by interface ID")
    interface_parser.add_argument("interface_id", help="Interface ID")
    interface_parser.add_argument("mq_folder", help="MQ XML folder path")
    interface_parser.add_argument("bw_folder", help="BW XML folder path")
    
    args = parser.parse_args()
    
    query_parser = QueryParser()
    
    if args.command == "find_table":
        table_results = query_parser.find_files_by_table(args.folder_path, args.table_name)
        query_parser.print_table_search_results(table_results, args.table_name)
    elif args.command == "compare":
        comparison_results = query_parser.compare_mq_bw_queries(args.mq_xml, args.bw_xml)
        print("\nComparison Results Summary:")
        print(f"송신 쿼리 비교 결과: {len(comparison_results['send'])} 개의 차이점 발견")
        print(f"수신 쿼리 비교 결과: {len(comparison_results['recv'])} 개의 차이점 발견")
    elif args.command == "compare_by_id":
        comparison_results = query_parser.compare_mq_bw_queries_by_interface_id(
            args.interface_id, args.mq_folder, args.bw_folder
        )
        print("\nComparison Results Summary:")
        print(f"송신 쿼리 비교 결과: {len(comparison_results['send'])} 개의 차이점 발견")
        print(f"수신 쿼리 비교 결과: {len(comparison_results['recv'])} 개의 차이점 발견")
    else:
        parser.print_help()

LLM 성능 벤치마크를 위한 프롬프트 제안

제공된 파이썬 소스 코드에서 특정 기능을 발췌하여 독립적인 모듈을 생성하는 과업을 LLM에게 요청하고, 그 결과물을 통해 모델의 성능을 벤치마킹하기 위한 프롬프트입니다.

이 프롬프트는 단순히 코드 생성을 요청하는 것을 넘어, 코드 분석 능력, 설계 능력, 설명 능력, 예외 처리 능력 등을 종합적으로 평가할 수 있도록 구조화되었습니다.
LLM에게 전달할 프롬프트

[역할 및 목표]
당신은 코드 리팩토링 및 모듈화에 특화된 전문가 파이썬 개발자입니다. 지금부터 여러 개의 파이썬 소스 파일을 제공할 것입니다. 당신의 임무는 이 파일들에서 특정 기능을 정확히 식별하고, 이를 재사용 가능하며 독립적인 단일 파이썬 모듈로 만들어내는 것입니다.

당신의 답변은 여러 LLM의 성능을 벤치마킹하는 데 사용될 것이며, 아래 평가 기준에 따라 분석될 것입니다. 따라서 코드의 정확성뿐만 아니라, 설계의 우수성과 설명의 명확성까지 모두 고려하여 답변해 주십시오.

[제공되는 소스 코드]
(아래 코드는 분석 대상이며, 이 코드들의 의존성을 파악하여 필요한 로직을 추출해야 합니다.)
Python

#############
## test24.py
#############

import sys
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import xml.etree.ElementTree as ET
import ast
from typing import Dict, List, Tuple, Optional

# comp_xml.py와 comp_q.py에서 필요한 클래스와 함수 import
from comp_xml import read_interface_block, XMLComparator
from comp_q import QueryParser

class InterfaceXMLToExcel:
    def __init__(self, excel_path: str, xml_dir: str, output_path: str = 'test24.xlsx'):
        """
        XML 파일에서 추출한 쿼리의 컬럼과 VALUES를 매핑하여 Excel 파일을 생성하는 클래스

        Args:
            excel_path (str): 인터페이스 정보가 있는 Excel 파일 경로
            xml_dir (str): XML 파일이 있는 디렉토리 경로
            output_path (str): 출력할 Excel 파일 경로
        """
        self.excel_path = excel_path
        self.xml_dir = xml_dir
        self.output_path = output_path
        self.query_parser = QueryParser()

        # Excel 파일 로드
        self.input_workbook = openpyxl.load_workbook(excel_path)
        self.input_worksheet = self.input_workbook.active

        # 출력 Excel 파일 생성
        self.output_workbook = openpyxl.Workbook()
        self.output_worksheet = self.output_workbook.active

        # 스타일 정의
        self.header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        self.header_font = Font(color='FFFFFF', bold=True, size=9)
        self.normal_font = Font(name='맑은 고딕', size=9)
        self.bold_font = Font(bold=True, size=9)
        self.center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        self.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

    def find_rcv_file(self, if_id: str) -> str:
        """
        주어진 인터페이스 ID에 해당하는 수신(RCV) XML 파일을 찾습니다.

        Args:
            if_id (str): 인터페이스 ID

        Returns:
            str: 찾은 파일의 경로, 없으면 None
        """
        if not if_id:
            print(f"Warning: Empty IF_ID provided")
            return None

        try:
            # 디렉토리 내의 모든 XML 파일 검색
            for file in os.listdir(self.xml_dir):
                if not file.startswith(if_id):
                    continue

                # 수신 파일 (.RCV.xml)
                if file.endswith('.RCV.xml'):
                    file_path = os.path.join(self.xml_dir, file)
                    return file_path

            print(f"Warning: No receive file found for IF_ID: {if_id}")
            return None

        except Exception as e:
            print(f"Error finding interface files: {e}")
            return None

    def extract_query_from_xml(self, xml_path: str) -> str:
        """
        XML 파일에서 SQL 쿼리를 추출합니다.

        Args:
            xml_path (str): XML 파일 경로

        Returns:
            str: 추출된 SQL 쿼리, 없으면 None
        """
        try:
            # XML 파일이 제대로 로드되었는지 확인
            if not os.path.exists(xml_path):
                print(f"Warning: XML file not found: {xml_path}")
                return None

            tree = ET.parse(xml_path)
            root = tree.getroot()

            # XML 내용이 유효한지 확인
            if root is None:
                print(f"Warning: Invalid XML content in file: {xml_path}")
                return None

            # SQL 노드 찾기
            sql_node = root.find(".//SQL")
            if sql_node is None or not sql_node.text:
                print(f"Warning: No SQL content found in file: {xml_path}")
                return None

            query = sql_node.text.strip()

            # 추출된 쿼리가 유효한지 확인
            if not query:
                print(f"Warning: Empty SQL query in file: {xml_path}")
                return None

            return query

        except ET.ParseError as e:
            print(f"Error parsing XML file {xml_path}: {e}")
            return None
        except Exception as e:
            print(f"Unexpected error processing file {xml_path}: {e}")
            return None

    def clean_value(self, value: str) -> str:
        """
        VALUES 항목에서 콜론(:)을 제거하고 TO_DATE 함수 내의 실제 값만 추출합니다.

        Args:
            value (str): 원본 값

        Returns:
            str: 정제된 값
        """
        if not value:
            return ""

        # 콜론(:) 제거
        cleaned_value = value.replace(':', '')

        # TO_DATE 함수 처리
        to_date_pattern = r'TO_DATE\(\s*:?([A-Za-z0-9_]+)\s*,\s*[\'"](.*?)[\'"](\s*\))'
        to_date_match = re.search(to_date_pattern, value, re.IGNORECASE)

        if to_date_match:
            # TO_DATE 함수에서 첫 번째 인자(실제 값)만 추출
            param_name = to_date_match.group(1)
            return param_name

        return cleaned_value

    def get_column_value_mapping(self, query: str) -> Dict[str, str]:
        """
        INSERT 쿼리에서 컬럼과 VALUES를 매핑합니다.

        Args:
            query (str): INSERT SQL 쿼리

        Returns:
            Dict[str, str]: 컬럼과 값의 매핑 딕셔너리
        """
        if not query:
            return {}

        # QueryParser를 사용하여 INSERT 쿼리 파싱
        insert_parts = self.query_parser.parse_insert_parts(query)
        if not insert_parts:
            print(f"Failed to parse INSERT query: {query}")
            return {}

        # 테이블 이름과 컬럼-값 매핑 추출
        table_name, columns = insert_parts

        # 값 정제 처리
        cleaned_columns = {}
        for col, val in columns.items():
            cleaned_columns[col] = self.clean_value(val)

        return cleaned_columns

    def process_interfaces(self):
        """
        Excel 파일에서 인터페이스 정보를 읽고, XML 파일에서 쿼리를 추출하여 매핑 후 출력 Excel 파일에 작성합니다.
        """
        # 헤더 행 복사 및 스타일 적용
        for col in range(1, self.input_worksheet.max_column + 1):
            self.output_worksheet.cell(row=1, column=col).value = self.input_worksheet.cell(row=1, column=col).value
            self.output_worksheet.cell(row=1, column=col).font = self.normal_font

            self.output_worksheet.cell(row=2, column=col).value = self.input_worksheet.cell(row=2, column=col).value
            self.output_worksheet.cell(row=2, column=col).font = self.normal_font

            self.output_worksheet.cell(row=3, column=col).value = self.input_worksheet.cell(row=3, column=col).value
            self.output_worksheet.cell(row=3, column=col).font = self.normal_font

            self.output_worksheet.cell(row=4, column=col).value = self.input_worksheet.cell(row=4, column=col).value
            self.output_worksheet.cell(row=4, column=col).font = self.normal_font

        # 인터페이스 블록 처리
        current_col = 2  # B열부터 시작
        interface_count = 0

        while current_col <= self.input_worksheet.max_column:
            try:
                # 인터페이스 정보 읽기
                interface_info = read_interface_block(self.input_worksheet, current_col)
                if not interface_info:
                    break

                interface_count += 1
                interface_id = interface_info.get('interface_id', '')
                interface_name = interface_info.get('interface_name', f'Interface_{interface_count}')

                print(f"\n처리 중인 인터페이스: {interface_name} (ID: {interface_id})")

                # 인터페이스 기본 정보 복사 및 스타일 적용
                self.output_worksheet.cell(row=1, column=current_col).value = interface_name
                self.output_worksheet.cell(row=1, column=current_col).font = self.normal_font

                self.output_worksheet.cell(row=2, column=current_col).value = interface_id
                self.output_worksheet.cell(row=2, column=current_col).font = self.normal_font

                self.output_worksheet.cell(row=3, column=current_col).value = self.input_worksheet.cell(row=3, column=current_col).value
                self.output_worksheet.cell(row=3, column=current_col).font = self.normal_font
                self.output_worksheet.cell(row=3, column=current_col + 1).value = self.input_worksheet.cell(row=3, column=current_col + 1).value
                self.output_worksheet.cell(row=3, column=current_col + 1).font = self.normal_font

                self.output_worksheet.cell(row=4, column=current_col).value = self.input_worksheet.cell(row=4, column=current_col).value
                self.output_worksheet.cell(row=4, column=current_col).font = self.normal_font
                self.output_worksheet.cell(row=4, column=current_col + 1).value = self.input_worksheet.cell(row=4, column=current_col + 1).value
                self.output_worksheet.cell(row=4, column=current_col + 1).font = self.normal_font

                # 수신 XML 파일 찾기
                rcv_file_path = self.find_rcv_file(interface_id)
                if not rcv_file_path:
                    print(f"Warning: No receive file found for interface {interface_name} (ID: {interface_id})")
                    current_col += 3  # 다음 인터페이스로 이동
                    continue

                # XML 파일에서 쿼리 추출
                query = self.extract_query_from_xml(rcv_file_path)
                if not query:
                    print(f"Warning: Failed to extract query from file {rcv_file_path}")
                    current_col += 3  # 다음 인터페이스로 이동
                    continue

                # 쿼리에서 컬럼-값 매핑 추출
                column_value_mapping = self.get_column_value_mapping(query)
                if not column_value_mapping:
                    print(f"Warning: Failed to extract column-value mapping from query")
                    current_col += 3  # 다음 인터페이스로 이동
                    continue

                # 특수 컬럼 제외
                special_columns = set(self.query_parser.special_columns['recv']['required'])
                filtered_mapping = {k: v for k, v in column_value_mapping.items() if k.upper() not in special_columns}

                # 매핑 정보를 Excel에 작성
                row = 5  # 5행부터 컬럼 매핑 시작
                for column, value in filtered_mapping.items():
                    # 수신 컬럼을 첫 번째 열(B열)에 배치
                    self.output_worksheet.cell(row=row, column=current_col).value = column  # 수신 컬럼을 첫 번째 열에 배치
                    self.output_worksheet.cell(row=row, column=current_col).font = self.normal_font

                    # VALUES 항목을 오른쪽 열(C열)에 배치 - 이미 정제된 값 사용
                    self.output_worksheet.cell(row=row, column=current_col + 1).value = value  # VALUES 항목 (콜론 제거와 TO_DATE 함수 처리가 적용됨)
                    self.output_worksheet.cell(row=row, column=current_col + 1).font = self.normal_font

                    row += 1

                print(f"인터페이스 {interface_name} (ID: {interface_id}) 처리 완료")

            except Exception as e:
                print(f"Error processing interface at column {current_col}: {str(e)}")

            current_col += 3  # 다음 인터페이스로 이동

        # 출력 파일 저장
        self.output_workbook.save(self.output_path)
        self.input_workbook.close()
        self.output_workbook.close()

        print(f"\n=== 처리 완료 ===")
        print(f"총 처리된 인터페이스 수: {interface_count}")
        print(f"출력 파일 저장 완료: {self.output_path}")

def main():
    try:
        # 하드코딩된 경로 설정 (python test24.py만 실행해도 작동하도록)
        # 현재 스크립트가 있는 디렉토리 기준으로 상대 경로 설정
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # 기본 경로 설정 (실제 환경에 맞게 수정 필요)
        excel_path = os.path.join(current_dir, 'C:\\work\\LT\\input_W7.xlsx')  # 인터페이스 정보 파일
        xml_dir = os.path.join(current_dir, 'C:\\work\\LT\\W7xml')  # XML 파일 디렉토리
        output_path = os.path.join(current_dir, 'C:\\work\\LT\\test24.xlsx')  # 출력 파일

        # 명령행 인수가 있으면 덮어쓰기
        if len(sys.argv) > 1:
            excel_path = sys.argv[1]
        if len(sys.argv) > 2:
            xml_dir = sys.argv[2]
        if len(sys.argv) > 3:
            output_path = sys.argv[3]

        print(f"사용할 파일 경로:")
        print(f"- 인터페이스 정보 파일: {excel_path}")
        print(f"- XML 파일 디렉토리: {xml_dir}")
        print(f"- 출력 파일: {output_path}")

        # 인터페이스 처리 실행
        processor = InterfaceXMLToExcel(excel_path, xml_dir, output_path)
        processor.process_interfaces()

    except Exception as e:
        print(f"\n[심각한 오류] 프로그램 실행 중 오류 발생: {str(e)}")
        raise

if __name__ == "__main__":
    main()

#############
## comp_xml.py
#############
import sys
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from typing import Dict, List, Tuple, Optional
import xml.etree.ElementTree as ET
from comp_excel import ExcelManager, read_interface_block
from xltest import process_interface, read_interface_block
from comp_q import QueryParser, QueryDifference, FileSearcher, BWQueryExtractor
from maptest import ColumnMapper
import datetime
import ast

def read_interface_block(ws, start_col):
    """Excel에서 3컬럼 단위로 하나의 인터페이스 정보를 읽습니다.
    이 함수는 xltest.py의 동일한 함수를 대체하지 않고, 가져오지 못한 경우의 백업 역할만 합니다.
    """
    try:
        interface_info = {
            'interface_name': ws.cell(row=1, column=start_col).value or '',  # IF NAME (1행)
            'interface_id': ws.cell(row=2, column=start_col).value or '',    # IF ID (2행)
            'send': {'owner': None, 'table_name': None, 'columns': [], 'db_info': None},
            'recv': {'owner': None, 'table_name': None, 'columns': [], 'db_info': None}
        }
        
        # 인터페이스 ID가 없으면 빈 인터페이스로 간주
        if not interface_info['interface_id']:
            return None
            
        # DB 연결 정보 (3행에서 읽기)
        try:
            send_db_value = ws.cell(row=3, column=start_col).value
            send_db_info = ast.literal_eval(send_db_value) if send_db_value else {}
            
            recv_db_value = ws.cell(row=3, column=start_col + 1).value
            recv_db_info = ast.literal_eval(recv_db_value) if recv_db_value else {}
        except (SyntaxError, ValueError):
            # 데이터 형식 오류 시 빈 딕셔너리로 설정
            send_db_info = {}
            recv_db_info = {}
            
        interface_info['send']['db_info'] = send_db_info
        interface_info['recv']['db_info'] = recv_db_info
        
        # 테이블 정보 (4행에서 읽기)
        try:
            send_table_value = ws.cell(row=4, column=start_col).value
            send_table_info = ast.literal_eval(send_table_value) if send_table_value else {}
            
            recv_table_value = ws.cell(row=4, column=start_col + 1).value
            recv_table_info = ast.literal_eval(recv_table_value) if recv_table_value else {}
        except (SyntaxError, ValueError):
            # 데이터 형식 오류 시 빈 딕셔너리로 설정
            send_table_info = {}
            recv_table_info = {}
        
        interface_info['send']['owner'] = send_table_info.get('owner')
        interface_info['send']['table_name'] = send_table_info.get('table_name')
        interface_info['recv']['owner'] = recv_table_info.get('owner')
        interface_info['recv']['table_name'] = recv_table_info.get('table_name')
        
        # 컬럼 매핑 정보 (5행부터)
        row = 5
        while True:
            send_col = ws.cell(row=row, column=start_col).value
            recv_col = ws.cell(row=row, column=start_col + 1).value
            
            if not send_col and not recv_col:
                break
                
            interface_info['send']['columns'].append(send_col if send_col else '')
            interface_info['recv']['columns'].append(recv_col if recv_col else '')
            row += 1
            
    except Exception as e:
        print(f'인터페이스 정보 읽기 중 오류 발생: {str(e)}')
        return None
    
    return interface_info

class XMLComparator:
    # 클래스 변수로 BW_SEARCH_DIR 정의
    BW_SEARCH_DIR = "C:\\work\\LT\\BW소스"

    def __init__(self, excel_path: str, search_dir: str):
        """
        XML 비교를 위한 클래스 초기화
        
        Args:
            excel_path (str): 인터페이스 정보가 있는 Excel 파일 경로
            search_dir (str): XML 파일을 검색할 디렉토리 경로
        """
        self.excel_path = excel_path
        self.search_dir = search_dir
        self.workbook = openpyxl.load_workbook(excel_path)
        self.worksheet = self.workbook.active
        self.mapper = ColumnMapper()
        self.query_parser = QueryParser()  # QueryParser 인스턴스 생성
        self.excel_manager = ExcelManager()  # ExcelManager 인스턴스 생성
        self.interface_results = []  # 모든 인터페이스 처리 결과 저장
        self.output_path = 'C:\\work\\LT\\comp_mq_bw.xlsx'  # 기본 출력 경로

    def extract_from_xml(self, xml_path: str) -> Tuple[str, str]:
        """
        XML 파일에서 쿼리와 XML 내용을 추출합니다.
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            Tuple[str, str]: (쿼리, XML 내용)
        """
        try:
            # XML 파일이 제대로 로드되었는지 확인
            if not os.path.exists(xml_path):
                print(f"Warning: XML file not found: {xml_path}")
                return None, None
                
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # XML 내용이 유효한지 확인
            if root is None:
                print(f"Warning: Invalid XML content in file: {xml_path}")
                return None, None
            
            # SQL 노드 찾기
            sql_node = root.find(".//SQL")
            if sql_node is None or not sql_node.text:
                print(f"Warning: No SQL content found in file: {xml_path}")
                return None, None
                
            query = sql_node.text.strip()
            xml_content = ET.tostring(root, encoding='unicode')
            
            # 추출된 쿼리가 유효한지 확인
            if not query:
                print(f"Warning: Empty SQL query in file: {xml_path}")
                return None, None
                
            return query, xml_content
            
        except ET.ParseError as e:
            print(f"Error parsing XML file {xml_path}: {e}")
            return None, None
        except Exception as e:
            print(f"Unexpected error processing file {xml_path}: {e}")
            return None, None
            
    def compare_queries(self, query1: str, query2: str) -> QueryDifference:
        """
        두 쿼리를 비교합니다.
        
        Args:
            query1 (str): 첫 번째 쿼리
            query2 (str): 두 번째 쿼리
            
        Returns:
            QueryDifference: 쿼리 비교 결과
        """
        if not query1 or not query2:
            return None
        return self.query_parser.compare_queries(query1, query2)
        
    def find_interface_files(self, if_id: str) -> Dict[str, Dict]:
        """
        주어진 IF ID에 해당하는 송수신 XML 파일을 찾고 쿼리를 추출합니다.
        파일명 패턴: {if_id}로 시작하고 .SND.xml 또는 .RCV.xml로 끝나는 파일
        
        Args:
            if_id (str): 인터페이스 ID
            
        Returns:
            Dict[str, Dict]: {
                'send': {'path': 송신파일경로, 'query': 송신쿼리, 'xml': 송신XML},
                'recv': {'path': 수신파일경로, 'query': 수신쿼리, 'xml': 수신XML}
            }
        """
        results = {
            'send': {'path': None, 'query': None, 'xml': None},
            'recv': {'path': None, 'query': None, 'xml': None}
        }
        
        if not if_id:
            print("Warning: Empty IF_ID provided")
            return results
            
        try:
            # 디렉토리 내의 모든 XML 파일 검색
            for file in os.listdir(self.search_dir):
                if not file.startswith(if_id):
                    continue
                    
                file_path = os.path.join(self.search_dir, file)
                
                # 송신 파일 (.SND.xml)
                if file.endswith('.SND.xml'):
                    results['send']['path'] = file_path
                    query, xml = self.extract_from_xml(file_path)
                    if query and xml:
                        results['send']['query'] = query
                        results['send']['xml'] = xml
                    else:
                        print(f"Warning: Failed to extract query from send file: {file_path}")
                
                # 수신 파일 (.RCV.xml)
                elif file.endswith('.RCV.xml'):
                    results['recv']['path'] = file_path
                    query, xml = self.extract_from_xml(file_path)
                    if query and xml:
                        results['recv']['query'] = query
                        results['recv']['xml'] = xml
                    else:
                        print(f"Warning: Failed to extract query from receive file: {file_path}")
            
            # 파일을 찾았는지 확인
            if not results['send']['path'] and not results['recv']['path']:
                print(f"Warning: No interface files found for IF_ID: {if_id}")
            elif not results['send']['path']:
                print(f"Warning: No send file found for IF_ID: {if_id}")
            elif not results['recv']['path']:
                print(f"Warning: No receive file found for IF_ID: {if_id}")
            
            return results
            
        except Exception as e:
            print(f"Error finding interface files: {e}")
            return results
        
    def process_interface_block(self, start_col: int) -> Optional[Dict]:
        """
        Excel에서 하나의 인터페이스 블록을 처리합니다.
        
        Args:
            start_col (int): 인터페이스 블록이 시작되는 컬럼
            
        Returns:
            Optional[Dict]: 처리된 인터페이스 정보와 결과, 실패시 None
        """
        try:
            # Excel에서 인터페이스 정보 읽기
            interface_info = read_interface_block(self.worksheet, start_col)
            if not interface_info:
                print(f"Warning: Failed to read interface block at column {start_col}")
                return None
                
            # Excel에서 추출된 쿼리와 XML 얻기
            excel_results = process_interface(interface_info, self.mapper)
            if not excel_results:
                print(f"Warning: Failed to process interface at column {start_col}")
                return None
                
            # 송수신 파일 찾기
            file_results = self.find_interface_files(interface_info['interface_id'])
            if not file_results:
                print(f"Warning: No interface files found for IF_ID: {interface_info['interface_id']}")
                return None
            
            # 결과 초기화
            comparisons = {
                'send': None,
                'recv': None
            }
            warnings = {
                'send': [],
                'recv': []
            }
            
            # 송신 쿼리 처리
            if excel_results['send_sql'] and file_results['send']['query']:
                try:
                    comparisons['send'] = self.query_parser.compare_queries(
                        excel_results['send_sql'],
                        file_results['send']['query']
                    )
                    warnings['send'].extend(
                        self.query_parser.check_special_columns(
                            file_results['send']['query'],
                            'send'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing send queries: {e}")
                    print(f"Excel query: {excel_results['send_sql']}")
                    print(f"File query: {file_results['send']['query']}")
            
            # 수신 쿼리 처리
            if excel_results['recv_sql'] and file_results['recv']['query']:
                try:
                    comparisons['recv'] = self.query_parser.compare_queries(
                        excel_results['recv_sql'],
                        file_results['recv']['query']
                    )
                    warnings['recv'].extend(
                        self.query_parser.check_special_columns(
                            file_results['recv']['query'],
                            'recv'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing receive queries: {e}")
                    print(f"Excel query: {excel_results['recv_sql']}")
                    print(f"File query: {file_results['recv']['query']}")
            
            return {
                'if_id': interface_info['interface_id'],
                'interface_name': interface_info['interface_name'],
                'comparisons': comparisons,
                'warnings': warnings,
                'excel': excel_results,
                'files': file_results
            }
            
        except Exception as e:
            print(f"Error processing interface block at column {start_col}: {e}")
            return None
            
    def process_all_interfaces(self) -> List[Dict]:
        """
        Excel 파일의 모든 인터페이스를 처리합니다.
        B열부터 시작하여 3컬럼 단위로 처리합니다.
        
        Returns:
            List[Dict]: 각 인터페이스의 처리 결과 목록
        """
        results = []
        col = 2  # B열부터 시작
        
        while True:
            # 인터페이스 ID가 없으면 종료
            if not self.worksheet.cell(row=2, column=col).value:
                break
                
            result = self.process_interface_block(col)
            if result:
                results.append(result)
                
            col += 3  # 다음 인터페이스 블록으로 이동
            
        # 결과 출력
        for idx, result in enumerate(results, 1):
            print(f"\n=== 인터페이스 {idx} ===")
            print(f"ID: {result['if_id']}")
            print(f"이름: {result['interface_name']}")
            
            print("\n파일 검색 결과:")
            print(f"송신 파일: {result['files']['send']['path']}")
            print(f"수신 파일: {result['files']['recv']['path']}")
            
            print("\n쿼리 비교 결과:")
            if result['comparisons']['send']:
                print("송신 쿼리:")
                print(f"  {result['comparisons']['send']}")
            if result['comparisons']['recv']:
                print("수신 쿼리:")
                print(f"  {result['comparisons']['recv']}")
            
            # 경고가 있을 때만 경고 섹션 출력
            send_warnings = result['warnings']['send']
            recv_warnings = result['warnings']['recv']
            if send_warnings or recv_warnings:
                print("\n경고:")
                if send_warnings:
                    print("송신 쿼리 경고:")
                    for warning in send_warnings:
                        print(f"  - {warning}")
                if recv_warnings:
                    print("수신 쿼리 경고:")
                    for warning in recv_warnings:
                        print(f"  - {warning}")
                
        return results
        
    def close(self):
        """리소스 정리"""
        self.workbook.close()
        if self.mapper:
            self.mapper.close_connections()

    def find_bw_files(self) -> List[Dict[str, str]]:
        """
        엑셀의 인터페이스 정보에서 송신 테이블명을 추출하여 BW 파일을 검색합니다.
        
        Returns:
            List[Dict[str, str]]: [
                {
                    'interface_name': str,
                    'interface_id': str,
                    'send_table': str,
                    'bw_files': List[str]
                },
                ...
            ]
        """
        results = []
        file_searcher = FileSearcher()
        
        # 엑셀에서 인터페이스 정보 읽기
        for row in range(2, self.worksheet.max_row + 1, 3):  # 3행씩 건너뛰며 읽기
            interface_info = read_interface_block(self.worksheet, row)
            if not interface_info:
                continue
                
            # 송신 테이블명 추출 (스키마/오너 제외)
            send_table = interface_info['send'].get('table_name')
            if not send_table:
                continue
                
            # BW 파일 검색 - self.BW_SEARCH_DIR 사용
            bw_files = file_searcher.find_files_with_keywords(self.BW_SEARCH_DIR, [send_table])
            matching_files = bw_files.get(send_table, [])
            
            results.append({
                'interface_name': interface_info['interface_name'],
                'interface_id': interface_info['interface_id'],
                'send_table': send_table,
                'bw_files': matching_files
            })
            
        return results
        
    def print_bw_search_results(self, results: List[Dict[str, str]]):
        """
        BW 파일 검색 결과를 출력합니다.
        
        Args:
            results (List[Dict[str, str]]): find_bw_files()의 반환값
        """
        print("\nBW File Search Results:")
        print("-" * 80)
        print(f"{'Interface Name':<30} {'Interface ID':<15} {'Send Table':<20} {'BW Files'}")
        print("-" * 80)
        
        for result in results:
            bw_files_str = ', '.join(result['bw_files']) if result['bw_files'] else 'No matching files'
            print(f"{result['interface_name']:<30} {result['interface_id']:<15} {result['send_table']:<20} {bw_files_str}")

    def initialize_excel_output(self):
        """
        결과를 저장할 새 엑셀 파일 초기화
        """
        # ExcelManager를 통해 Excel 출력을 초기화
        self.excel_manager.initialize_excel_output()
        
    def save_excel_output(self, output_path=None):
        """
        처리된 결과를 엑셀 파일로 저장
        
        Args:
            output_path (str, optional): 출력 엑셀 파일 경로, 없으면 기본 경로 사용
            
        Returns:
            bool: 저장 성공 여부
        """
        # output_path 값을 인스턴스 변수에 저장
        if output_path:
            self.output_path = output_path
            
        # ExcelManager를 사용하여 파일 저장
        return self.excel_manager.save_excel_output(self.output_path)
        
    def create_interface_sheet(self, if_info, file_results, query_comparisons, bw_queries=None, bw_files=None):
        """
        인터페이스 정보와 비교 결과를 포함하는 엑셀 시트를 생성합니다.
        
        Args:
            if_info (dict): 인터페이스 정보
            file_results (dict): MQ 파일 결과 (송신/수신)
            query_comparisons (dict): 쿼리 비교 결과 (송신/수신)
            bw_queries (dict, optional): BW 쿼리 정보. Defaults to None.
            bw_files (list, optional): BW 매핑 파일 목록. Defaults to None.
        """
        # 기본값 설정
        bw_queries = bw_queries or {'send': '', 'recv': ''}
        bw_files = bw_files or []
        
        # 인터페이스 ID와 이름 확인
        if 'interface_id' not in if_info or not if_info['interface_id']:
            print("인터페이스 ID가 없습니다.")
            return
            
        # BW 파일 매핑
        bw_files_dict = {
            'send': bw_files[0] if bw_files and len(bw_files) > 0 else 'N/A',
            'recv': bw_files[1] if bw_files and len(bw_files) > 1 else 'N/A'
        }
        
        # MQ 파일 정보
        mq_files = {
            'send': file_results.get('send', {}),
            'recv': file_results.get('recv', {})
        }
        
        # 쿼리 정보 구성
        queries = {
            'mq_send': file_results.get('send', {}).get('query', 'N/A'),
            'bw_send': bw_queries.get('send', 'N/A'),
            'mq_recv': file_results.get('recv', {}).get('query', 'N/A'),
            'bw_recv': bw_queries.get('recv', 'N/A')
        }
        
        # 비교 결과 구성
        comparison_results = {
            'send': {
                'is_equal': query_comparisons.get('send', QueryDifference()).is_equal,
                'detail': self._get_difference_detail(query_comparisons.get('send', QueryDifference()))
            },
            'recv': {
                'is_equal': query_comparisons.get('recv', QueryDifference()).is_equal,
                'detail': self._get_difference_detail(query_comparisons.get('recv', QueryDifference()))
            }
        }
        
        # 인터페이스 시트 생성
        self.excel_manager.create_interface_sheet(if_info, mq_files, bw_files_dict, queries, comparison_results)
        
    def process_interface_with_bw(self, start_col: int, interface_info: Dict) -> Optional[Dict]:
        """
        하나의 인터페이스를 처리하고 BW 파일과 비교하여 결과 반환
        
        Args:
            start_col (int): 인터페이스 블록이 시작되는 컬럼
            interface_info (Dict): 인터페이스 정보
            
        Returns:
            Optional[Dict]: 처리된 인터페이스 정보와 결과, 실패시 None
        """
        try:
            # 표준 필드 생성
            # DB 정보에서 시스템 정보 추출
            if 'send' in interface_info and 'db_info' in interface_info['send'] and interface_info['send']['db_info']:
                interface_info['send_system'] = interface_info['send']['db_info'].get('system', 'N/A')
            else:
                interface_info['send_system'] = 'N/A'
                
            if 'recv' in interface_info and 'db_info' in interface_info['recv'] and interface_info['recv']['db_info']:
                interface_info['recv_system'] = interface_info['recv']['db_info'].get('system', 'N/A')
            else:
                interface_info['recv_system'] = 'N/A'
                
            # 테이블 정보 추출
            if 'send' in interface_info and 'table_name' in interface_info['send']:
                interface_info['send_table'] = interface_info['send']['table_name']
            else:
                interface_info['send_table'] = ''
                
            if 'recv' in interface_info and 'table_name' in interface_info['recv']:
                interface_info['recv_table'] = interface_info['recv']['table_name']
            else:
                interface_info['recv_table'] = ''
                
            # Excel에서 추출된 쿼리와 XML 얻기
            excel_results = process_interface(interface_info, self.mapper)
            if not excel_results:
                print(f"Warning: Failed to process interface at column {start_col}")
                return None
                
            # 송수신 파일 찾기
            file_results = self.find_interface_files(interface_info['interface_id'])
            if not file_results:
                print(f"Warning: No interface files found for IF_ID: {interface_info['interface_id']}")
                return None
            
            # BW 파일 찾기
            send_table = interface_info.get('send_table', '')
            if not send_table:
                print(f"Warning: No send table information for IF_ID: {interface_info['interface_id']}")
                bw_files = []
            else:
                # 송신 테이블로 BW 파일 검색
                bw_searcher = FileSearcher()
                bw_files = bw_searcher.find_files_with_keywords(
                    self.BW_SEARCH_DIR, 
                    [send_table]
                )
            
            # bw_files가 사전 형태이므로 send_table 키워드에 대한 결과를 가져옴
            matching_files = bw_files.get(send_table, [])
            
            # BW 쿼리 추출
            bw_queries = {
                'send': '',
                'recv': ''
            }
            extractor = BWQueryExtractor()
            for bw_file in matching_files:
                bw_file_path = os.path.join(self.BW_SEARCH_DIR, bw_file)
                if os.path.exists(bw_file_path):
                    # BWQueryExtractor의 extract_bw_queries 메서드를 사용하여 송신/수신 쿼리 모두 추출
                    queries = extractor.extract_bw_queries(bw_file_path)
                    
                    # 송신 쿼리가 없으면 첫 번째 송신 쿼리 저장
                    if not bw_queries['send'] and queries.get('send') and len(queries['send']) > 0:
                        bw_queries['send'] = queries['send'][0]
                    
                    # 수신 쿼리가 없으면 첫 번째 수신 쿼리 저장
                    if not bw_queries['recv'] and queries.get('recv') and len(queries['recv']) > 0:
                        bw_queries['recv'] = queries['recv'][0]
            
            # 결과 초기화
            comparisons = {
                'send': None,
                'recv': None
            }
            warnings = {
                'send': [],
                'recv': []
            }
            
            # 송신 쿼리 비교 (MQ XML vs BW XML)
            if file_results['send']['query'] and bw_queries['send']:
                try:
                    comparisons['send'] = self.query_parser.compare_queries(
                        file_results['send']['query'],
                        bw_queries['send']
                    )
                    warnings['send'].extend(
                        self.query_parser.check_special_columns(
                            file_results['send']['query'],
                            'send'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing send queries: {e}")
                    print(f"MQ query: {file_results['send']['query']}")
                    print(f"BW query: {bw_queries['send']}")
            
            # 수신 쿼리 비교 (MQ XML vs BW XML)
            if file_results['recv']['query'] and bw_queries['recv']:
                try:
                    comparisons['recv'] = self.query_parser.compare_queries(
                        file_results['recv']['query'],
                        bw_queries['recv']
                    )
                    warnings['recv'].extend(
                        self.query_parser.check_special_columns(
                            file_results['recv']['query'],
                            'recv'
                        )
                    )
                except Exception as e:
                    print(f"Error comparing recv queries: {e}")
                    print(f"MQ query: {file_results['recv']['query']}")
                    print(f"BW query: {bw_queries['recv']}")
            
            # 결과 반환
            return {
                'interface_info': interface_info,
                'excel_results': excel_results,
                'file_results': file_results,
                'bw_queries': bw_queries,
                'comparisons': comparisons,
                'warnings': warnings,
                'bw_files': matching_files
            }
            
        except Exception as e:
            print(f"Error processing interface at column {start_col}: {e}")
            import traceback
            traceback.print_exc()
            return None

    def process_all_interfaces_with_bw(self):
        """
        모든 인터페이스를 처리하고 BW 파일과 비교하여 엑셀 파일로 결과 저장
        """
        # 엑셀 파일 초기화 - ExcelManager 사용
        self.excel_manager.initialize_excel_output()
        
        # 모든 열을 처리
        print("\n[인터페이스 처리 시작]")
        print("-" * 80)
        
        interface_count = 0
        processed_count = 0
        
        start_col = 2
        while True:
            interface_info = read_interface_block(self.worksheet, start_col)
            
            if not interface_info:
                break
                
            interface_count += 1
            
            # 인터페이스 ID와 이름 출력
            print(f"처리 중: [{interface_count}] {interface_info['interface_id']} - {interface_info['interface_name']}")
            
            # 인터페이스 처리 및 BW 비교
            result = self.process_interface_with_bw(start_col, interface_info)
            
            # 다음 인터페이스로 이동 (3칸씩)
            start_col += 3
            
            # 인터페이스 처리 결과가 있으면 엑셀에 저장
            if result:
                processed_count += 1
                
                # 결과를 저장할 인터페이스 시트 생성
                if_info = result['interface_info']
                
                # ExcelManager를 사용하여 인터페이스 시트 생성
                # MQ 파일 정보
                mq_files = {
                    'send': result['file_results']['send'],
                    'recv': result['file_results']['recv']
                }
                
                # BW 파일 정보
                bw_files = {
                    'send': result.get('bw_files', [])[0] if result.get('bw_files') and len(result.get('bw_files')) > 0 else 'N/A',
                    'recv': result.get('bw_files', [])[1] if result.get('bw_files') and len(result.get('bw_files')) > 1 else 'N/A'
                }
                
                # 쿼리 정보
                queries = {
                    'mq_send': result['file_results']['send']['query'],
                    'bw_send': result['bw_queries']['send'],
                    'mq_recv': result['file_results']['recv']['query'],
                    'bw_recv': result['bw_queries']['recv']
                }
                
                # 비교 결과
                comparison_results = {
                    'send': {
                        'is_equal': result['comparisons']['send'].is_equal if result['comparisons']['send'] else False,
                        'detail': self._get_difference_detail(result['comparisons']['send']) if result['comparisons']['send'] else '비교 불가'
                    },
                    'recv': {
                        'is_equal': result['comparisons']['recv'].is_equal if result['comparisons']['recv'] else False,
                        'detail': self._get_difference_detail(result['comparisons']['recv']) if result['comparisons']['recv'] else '비교 불가'
                    }
                }
                
                self.excel_manager.create_interface_sheet(
                    if_info, 
                    mq_files, 
                    bw_files, 
                    queries, 
                    comparison_results
                )
                
                # 요약 시트 업데이트
                self.update_summary_sheet(result, interface_count + 1)
        
        # 결과 저장
        self.save_excel_output()
        
        # 처리 결과 출력
        print("\n" + "=" * 80)
        print(f"처리 완료: 총 {interface_count}개 인터페이스 중 {processed_count}개 처리됨")
        print(f"결과 파일: {self.output_path}")
        print("=" * 80)
        
    def update_summary_sheet(self, result, row):
        """
        요약 시트에 현재 인터페이스 처리 결과를 추가합니다.
        
        Args:
            result (dict): 인터페이스 처리 결과
            row (int): 추가할 행 번호
        """
        # ExcelManager를 사용하여 요약 시트 업데이트
        self.excel_manager.update_summary_sheet(result, row)

    def extract_bw_queries(self, bw_results):
        """
        BW 파일에서 쿼리를 추출합니다.
        
        Args:
            bw_results (list): BW 파일 검색 결과 목록
            
        Returns:
            list: 인터페이스별 BW 쿼리 정보가 담긴 리스트
        """
        extractor = BWQueryExtractor()
        results = []
        
        for result in bw_results:
            if result['bw_files']:  # BW 파일이 있는 경우에만 처리
                print(f"\n인터페이스: {result['interface_name']} ({result['interface_id']})")
                print(f"송신 테이블: {result['send_table']}")
                print("찾은 BW 파일의 쿼리:")
                
                bw_queries = {'send': '', 'recv': ''}
                
                for bw_file in result['bw_files']:
                    bw_file_path = os.path.join(self.BW_SEARCH_DIR, bw_file)
                    if os.path.exists(bw_file_path):
                        # BWQueryExtractor의 extract_bw_queries 메서드를 사용하여 송신/수신 쿼리 모두 추출
                        queries = extractor.extract_bw_queries(bw_file_path)
                        
                        if queries['send'] and not bw_queries['send']:
                            bw_queries['send'] = queries['send'][0] if queries['send'] else ''
                            print(f"\nBW 송신 파일: {bw_file}")
                            print("-" * 40)
                            print(bw_queries['send'])
                            
                        if queries['recv'] and not bw_queries['recv']:
                            bw_queries['recv'] = queries['recv'][0] if queries['recv'] else ''
                            print(f"\nBW 수신 파일: {bw_file}")
                            print("-" * 40)
                            print(bw_queries['recv'])
                
                # 인터페이스 결과에 BW 쿼리 추가
                for interface_result in self.interface_results:
                    if interface_result['interface_info']['interface_id'] == result['interface_id']:
                        interface_result['bw_queries'] = bw_queries
                        interface_result['bw_files'] = result['bw_files']
                        break
                
                results.append({
                    'interface_id': result['interface_id'],
                    'bw_queries': bw_queries,
                    'bw_files': result['bw_files']
                })
        
        return results

    def _get_difference_detail(self, query_diff):
        """
        쿼리 차이점을 텍스트로 변환합니다.
        
        Args:
            query_diff (QueryDifference): 쿼리 차이점
        
        Returns:
            str: 차이점 텍스트
        """
        if not query_diff:
            return ''
        
        if query_diff.is_equal:
            return '일치 - 테이블과 칼럼이 모두 동일합니다.'
        
        # 차이점 텍스트 생성
        detail = '차이 - 다음과 같은 차이점이 발견되었습니다:\n'
        for diff in query_diff.differences:
            column = diff.get('column', 'N/A')
            query1_value = diff.get('query1_value', 'N/A')
            query2_value = diff.get('query2_value', 'N/A')
            detail += f'- 컬럼: {column}\n'
            detail += f'  · MQ: {query1_value}\n'
            detail += f'  · BW: {query2_value}\n'
        
        return detail

def main():
    # 고정된 경로 사용
    excel_path = 'C:\\work\\LT\\input_LT.xlsx' # 인터페이스 정보
    xml_dir = 'C:\\work\\LT\\xml' # MQ XML 파일 디렉토리
    bw_dir = 'C:\\work\\LT\\BW소스'  # BW XML 파일 디렉토리 경로
    output_path = 'C:\\work\\LT\\comp_mq_bw.xlsx'  # 출력 엑셀 파일 경로
    
    # BW 검색 디렉토리 설정
    XMLComparator.BW_SEARCH_DIR = bw_dir
    
    # XML 비교기 초기화
    comparator = XMLComparator(excel_path, xml_dir)
    
    # 명령행 인자가 있을 경우 처리
    if len(sys.argv) > 1:
        if sys.argv[1] == "excel":
            # 엑셀 출력 모드 실행
            print("\n[MQ XML과 BW XML 쿼리 비교 - 엑셀 출력 모드]")
            comparator.process_all_interfaces_with_bw()
            return
        elif len(sys.argv) > 2 and sys.argv[1] == "output":
            # 출력 경로 변경
            output_path = sys.argv[2]
            comparator.output_path = output_path
            print(f"\n[출력 경로 변경: {output_path}]")
    
    # 기본 모드 실행 - 기존 로직 유지
    print("\n[MQ XML 파일 검색 및 쿼리 비교 시작]")
    comparator.process_all_interfaces()
    
    # BW 파일 검색 및 결과 출력을 마지막으로 이동
    print("\n[BW 파일 검색 시작]")
    bw_results = comparator.find_bw_files()
    comparator.print_bw_search_results(bw_results)
    
    # BW 파일에서 쿼리 추출
    print("\n[BW 파일 쿼리 추출]")
    print("-" * 80)
    bw_queries = comparator.extract_bw_queries(bw_results)
    
    # 처리 결과를 Excel로 저장 (excel 모드가 아닌 경우)
    print("\n[결과를 Excel로 저장]")
    comparator.initialize_excel_output()
    
    # 인터페이스별 결과 처리
    for i, result in enumerate(comparator.interface_results):
        if_info = result['interface_info']
        
        # 인터페이스 시트 생성
        comparator.create_interface_sheet(
            if_info, 
            result['file_results'], 
            result['comparisons'],
            result.get('bw_queries', {'send': '', 'recv': ''}),
            result.get('bw_files', [])
        )
        
        # 요약 시트 업데이트
        comparator.update_summary_sheet(result, i + 2)
    
    # 결과 저장
    comparator.save_excel_output(output_path)
    print(f"\n[분석 완료] 결과가 저장되었습니다: {output_path}")
    
    print("\n[처리 완료]")
    print("엑셀 출력 모드로 실행하려면 'python comp_xml.py excel' 명령을 사용하세요.")

if __name__ == "__main__":
    main()

#############
## comp_q.py
#############
import xml.etree.ElementTree as ET
import re
from typing import Dict, List, Tuple, Optional
import os
import argparse

class QueryDifference:
    def __init__(self):
        self.is_equal = True
        self.differences = []
        self.query_type = None
        self.table_name = None
    
    def add_difference(self, column: str, value1: str, value2: str):
        self.is_equal = False
        self.differences.append({
            'column': column,
            'query1_value': value1,
            'query2_value': value2
        })

    def __str__(self) -> str:
        if self.is_equal:
            return "일치"
        
        return "불일치"

class QueryParser:
    # 특수 컬럼 정의를 클래스 변수로 변경
    special_columns = {
        'send': {
            'required': ['EAI_SEQ_ID', 'DATA_INTERFACE_TYPE_CODE'],
            'mappings': []  # 추가 매핑을 저장할 리스트
        },
        'recv': {
            'required': [
                'EAI_SEQ_ID',
                'DATA_INTERFACE_TYPE_CODE',
                'EAI_INTERFACE_DATE',
                'APPLICATION_TRANSFER_FLAG'
            ],
            'special_values': {
                'EAI_INTERFACE_DATE': 'SYSDATE',
                'APPLICATION_TRANSFER_FLAG': "'N'"
            }
        }
    }

    def __init__(self):
        self.select_queries = []
        self.insert_queries = []

    def normalize_query(self, query):
        """
        Normalize a SQL query by removing extra whitespace and standardizing format
        
        Args:
            query (str): SQL query to normalize
            
        Returns:
            str: Normalized query
        """
        print(f"Original query: {query}")
        
        # Remove comments if any
        query = re.sub(r'--.*$', '', query, flags=re.MULTILINE)
        
        # Replace newlines with spaces
        query = re.sub(r'\n', ' ', query)
        
        # Replace multiple whitespace with single space
        query = re.sub(r'\s+', ' ', query)
        
        # 핵심 SQL 키워드 주변에 공백 추가 (대소문자 구분 없이)
        keywords = ['SELECT', 'FROM', 'WHERE', 'ORDER BY', 'GROUP BY', 'HAVING', 
                   'JOIN', 'LEFT', 'RIGHT', 'INNER', 'OUTER', 'ON', 'AS']
        
        # 각 키워드를 공백으로 둘러싸기 (단어 경계 고려)
        for keyword in keywords:
            # \b는 단어 경계를 의미
            pattern = re.compile(r'\b' + keyword + r'\b', re.IGNORECASE)
            # 각 키워드를 찾아서 앞뒤에 공백 추가
            query = pattern.sub(' ' + keyword + ' ', query)
        
        # 다시 중복 공백 제거
        query = re.sub(r'\s+', ' ', query)
        
        result = query.strip()
        print(f"Normalized query: {result}")
        return result

    def parse_select_columns(self, query) -> Optional[Dict[str, str]]:
        """Extract columns from SELECT query and return as dictionary"""
        # 대소문자 구분 없이 정규화
        print(f"Parsing query: {query}")
        
        # SELECT 키워드 위치 찾기
        select_match = re.search(r'\bSELECT\b', query, re.IGNORECASE)
        from_match = re.search(r'\bFROM\b', query, re.IGNORECASE)
        
        if not select_match or not from_match:
            print(f"Could not find SELECT or FROM keywords in query: {query}")
            return None
        
        # SELECT와 FROM 사이의 부분 추출
        select_pos = select_match.end()
        from_pos = from_match.start()
        
        if select_pos >= from_pos:
            print(f"Invalid query structure (SELECT appears after FROM): {query}")
            return None
        
        # 컬럼 부분 추출
        column_part = query[select_pos:from_pos].strip()
        print(f"Extracted column part: {column_part}")
        
        # 컬럼 분리 및 처리
        columns = {}
        for col in self._parse_csv_with_functions(column_part):
            col = col.strip()
            if not col:
                continue
            
            print(f"Processing column: {col}")
            
            # to_char 함수 처리 (별칭 유무에 관계없이)
            if 'to_char(' in col.lower():
                # 함수 호출 이후에 별칭이 있는지 확인
                alias_match = re.search(r'(to_char\s*\([^)]+\))\s+([a-zA-Z0-9_]+)$', col, re.IGNORECASE)
                if alias_match:
                    # 별칭이 있는 경우
                    expr, alias = alias_match.groups()
                    print(f"Found to_char with alias: {expr} -> {alias}")
                    columns[expr.strip()] = {'expr': expr.strip(), 'alias': alias.strip(), 'full': col}
                else:
                    # 별칭이 없는 경우
                    print(f"Found to_char without alias: {col}")
                    columns[col] = {'expr': col, 'alias': None, 'full': col}
            else:
                # 일반 열 처리
                alias_match = re.search(r'(.+?)\s+(?:AS\s+)?([a-zA-Z0-9_]+)$', col, re.IGNORECASE)
                if alias_match:
                    expr, alias = alias_match.groups()
                    print(f"Found column with alias: {expr} -> {alias}")
                    columns[expr.strip()] = {'expr': expr.strip(), 'alias': alias.strip(), 'full': col}
                else:
                    print(f"Found column without alias: {col}")
                    columns[col] = {'expr': col, 'alias': None, 'full': col}
        
        print(f"Final parsed columns: {columns}")
        return columns if columns else None

    def _extract_values_with_balanced_parentheses(self, query, start_idx):
        """
        INSERT 쿼리에서 VALUES 절의 내용을 괄호 균형을 맞추며 추출
        
        Args:
            query (str): 전체 쿼리 문자열
            start_idx (int): VALUES 키워드 이후의 시작 인덱스
            
        Returns:
            str: 추출된 VALUES 절 내용 (괄호 포함)
        """
        paren_count = 0
        in_quotes = False
        quote_char = None
        idx = start_idx
        
        while idx < len(query):
            char = query[idx]
            
            # 따옴표 처리 ('나 " 내부에서는 괄호를 계산하지 않음)
            if char in ["'", '"'] and (idx == 0 or query[idx-1] != '\\'):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
            
            # 괄호 카운팅 (따옴표 밖에서만)
            if not in_quotes:
                if char == '(':
                    paren_count += 1
                    if paren_count == 1 and idx == start_idx:  # 시작 괄호
                        start_idx = idx
                elif char == ')':
                    paren_count -= 1
                    if paren_count == 0:  # 종료 괄호 도달
                        return query[start_idx:idx+1]
            
            idx += 1
        
        # 괄호가 맞지 않는 경우
        return None

    def _parse_csv_with_functions(self, csv_string):
        """
        함수 호출과 따옴표를 고려하여 CSV 문자열을 파싱합니다.
        
        Args:
            csv_string (str): 파싱할 CSV 문자열
            
        Returns:
            List[str]: 파싱된 값 목록
        """
        results = []
        current = ""
        paren_count = 0
        in_quotes = False
        quote_char = None
        
        for i, char in enumerate(csv_string):
            # 따옴표 처리
            if char in ["'", '"'] and (i == 0 or csv_string[i-1] != '\\'):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
            
            # 괄호 카운팅 (따옴표 안이 아닐 때)
            if not in_quotes:
                if char == '(':
                    paren_count += 1
                elif char == ')':
                    paren_count -= 1
                    
                # 값 구분자 처리
                if char == ',' and paren_count == 0:
                    results.append(current.strip())
                    current = ""
                    continue
            
            # 현재 문자 추가
            current += char
        
        # 마지막 값 추가
        if current:
            results.append(current.strip())
            
        return results

    def parse_insert_parts(self, query) -> Optional[Tuple[str, Dict[str, str]]]:
        """Extract and return table name and column-value pairs from INSERT query"""
        try:
            # 정규화된 쿼리 사용
            query = self.normalize_query(query)
            print(f"\nProcessing INSERT query:\n{query}")
            
            # INSERT INTO와 테이블 이름 추출
            table_match = re.search(r'INSERT\s+INTO\s+([A-Za-z0-9_$.]+)', query, flags=re.IGNORECASE)
            if not table_match:
                print("Failed to match INSERT INTO pattern")
                return None
                
            table_name = table_match.group(1)
            print(f"Found table name: {table_name}")
            
            # 컬럼 목록 추출
            columns_match = re.search(r'INSERT\s+INTO\s+[A-Za-z0-9_$.]+\s*\((.*?)\)', query, flags=re.IGNORECASE | re.DOTALL)
            if not columns_match:
                print("Failed to match columns pattern")
                return None
                
            col_names = [c.strip() for c in columns_match.group(1).split(',')]
            
            # VALUES 키워드 찾기
            values_match = re.search(r'VALUES\s*\(', query, flags=re.IGNORECASE)
            if not values_match:
                print("Failed to find VALUES keyword")
                return None
                
            # VALUES 절 추출 (괄호 균형 맞추며)
            values_start_idx = values_match.end() - 1  # '(' 위치
            values_part = self._extract_values_with_balanced_parentheses(query, values_start_idx)
            
            if not values_part:
                print("Failed to extract balanced VALUES part")
                return None
                
            # 괄호 제거하고 값만 추출
            values_str = values_part[1:-1]  # 시작과 끝 괄호 제거
            
            # 값 파싱 - 함수 호출을 고려한 파싱
            col_values = self._parse_csv_with_functions(values_str)
            
            print(f"Found columns: {col_names}")
            print(f"Found values: {col_values}")
            
            # 컬럼과 값의 개수가 일치하는지 확인
            if len(col_names) != len(col_values):
                print(f"Column count ({len(col_names)}) does not match value count ({len(col_values)})")
                return None
                
            # 빈 컬럼이나 값이 있는지 확인
            if not all(col_names) or not all(col_values):
                print("Found empty column names or values")
                return None
                
            columns = {}
            for name, value in zip(col_names, col_values):
                columns[name] = value
                
            print(f"Successfully parsed {len(columns)} columns")
            return (table_name, columns)
        except Exception as e:
            print(f"Error parsing INSERT parts: {str(e)}")
            return None

    def compare_queries(self, query1: str, query2: str) -> QueryDifference:
        """
        Compare two SQL queries and return detailed differences
        
        Args:
            query1 (str): First SQL query
            query2 (str): Second SQL query
            
        Returns:
            QueryDifference: Object containing comparison results and differences
        """
        result = QueryDifference()
        
        # 쿼리 정규화
        norm_query1 = self.normalize_query(query1)
        norm_query2 = self.normalize_query(query2)
        
        # 쿼리 타입 확인
        if re.search(r'SELECT', norm_query1, flags=re.IGNORECASE):
            result.query_type = 'SELECT'
            columns1 = self.parse_select_columns(query1)
            columns2 = self.parse_select_columns(query2)
            table1 = self.extract_table_name(query1)
            table2 = self.extract_table_name(query2)
            
            if columns1 is None or columns2 is None:
                raise ValueError("SELECT 쿼리 파싱 실패")
                
        elif re.search(r'INSERT', norm_query1, flags=re.IGNORECASE):
            result.query_type = 'INSERT'
            insert_result1 = self.parse_insert_parts(query1)
            insert_result2 = self.parse_insert_parts(query2)
            
            if insert_result1 is None or insert_result2 is None:
                raise ValueError("INSERT 쿼리 파싱 실패")
                
            table1, columns1 = insert_result1
            table2, columns2 = insert_result2
        else:
            raise ValueError("지원하지 않는 쿼리 타입입니다.")
            
        result.table_name = table1
        
        # 특수 컬럼 제외
        direction = 'recv' if result.query_type == 'INSERT' else 'send'
        special_cols = set(self.special_columns[direction]['required'])
        
        # 정규화된 비교를 위해 컬럼 처리
        if result.query_type == 'SELECT':
            # 결과가 동일한지 계산
            result.is_equal = self._compare_select_columns(columns1, columns2, special_cols, result)
        else:
            # 일반 컬럼만 비교 (대소문자 구분 없이 비교하되 원본 케이스 유지)
            columns1_filtered = {k: v for k, v in columns1.items() if k.upper() not in special_cols}
            columns2_filtered = {k: v for k, v in columns2.items() if k.upper() not in special_cols}
            
            # 컬럼 비교
            all_columns = set(columns1_filtered.keys()) | set(columns2_filtered.keys())
            is_equal = True
            for col in all_columns:
                if col not in columns1_filtered:
                    result.add_difference(col, None, columns2_filtered[col])
                    is_equal = False
                elif col not in columns2_filtered:
                    result.add_difference(col, columns1_filtered[col], None)
                    is_equal = False
                else:
                    # 값 비교 시 to_char 함수의 포맷 차이를 무시
                    val1 = columns1_filtered[col]
                    val2 = columns2_filtered[col]
                    
                    # to_char 또는 to_date 함수를 포함하는 값이면 정규화 적용
                    if 'to_char(' in val1.lower() or 'to_char(' in val2.lower() or 'to_date(' in val1.lower() or 'to_date(' in val2.lower():
                        norm_val1 = self._normalize_tochar_format(val1)
                        norm_val2 = self._normalize_tochar_format(val2)
                        if norm_val1 != norm_val2:
                            result.add_difference(col, val1, val2)
                            is_equal = False
                    else:
                        # 일반 값은 그대로 비교
                        if val1 != val2:
                            result.add_difference(col, val1, val2)
                            is_equal = False
            
            result.is_equal = is_equal
                
        return result
    
    def _compare_select_columns(self, columns1, columns2, special_cols, result):
        """
        SELECT 쿼리의 컬럼을 비교하는 보조 메소드
        
        Args:
            columns1: 첫 번째 쿼리의 컬럼 정보
            columns2: 두 번째 쿼리의 컬럼 정보
            special_cols: 특수 컬럼 집합
            result: 결과를 저장할 QueryDifference 객체
            
        Returns:
            bool: 두 쿼리의 컬럼이 동일한지 여부
        """
        # 일반 컬럼만 필터링 (특수 컬럼 제외)
        columns1_filtered = {k: v for k, v in columns1.items() 
                            if k.upper() not in special_cols and 
                              (v['alias'] is None or v['alias'].upper() not in special_cols)}
        columns2_filtered = {k: v for k, v in columns2.items() 
                            if k.upper() not in special_cols and 
                              (v['alias'] is None or v['alias'].upper() not in special_cols)}
        
        # 두 쿼리의 모든 컬럼 표현식 목록 생성
        expr1_set = {info['expr'].strip() for info in columns1_filtered.values()}
        expr2_set = {info['expr'].strip() for info in columns2_filtered.values()}
        
        # 정규화된 표현식 매핑 생성 - 공백 차이 등을 무시
        norm_expr1_map = {}
        for info in columns1_filtered.values():
            expr = info['expr'].strip()
            # to_char 함수의 포맷 문자열을 정규화 (포맷 차이 무시)
            norm_expr = self._normalize_tochar_format(expr)
            norm_expr1_map[norm_expr] = expr
            
        norm_expr2_map = {}
        for info in columns2_filtered.values():
            expr = info['expr'].strip()
            # to_char 함수의 포맷 문자열을 정규화 (포맷 차이 무시)
            norm_expr = self._normalize_tochar_format(expr)
            norm_expr2_map[norm_expr] = expr
            
        # 정규화된 표현식 세트
        norm_expr1_set = set(norm_expr1_map.keys())
        norm_expr2_set = set(norm_expr2_map.keys())
        
        # 정규화된 표현식으로 비교 (별칭과 공백 차이 무시)
        only_in_query1 = norm_expr1_set - norm_expr2_set
        only_in_query2 = norm_expr2_set - norm_expr1_set
        
        is_equal = True
        
        # 첫 번째 쿼리에만 있는 표현식 처리
        for norm_expr in only_in_query1:
            orig_expr = norm_expr1_map[norm_expr]
            # 이 표현식을 포함하는 컬럼 정보 찾기
            for col, info in columns1_filtered.items():
                norm_col_expr = self._normalize_tochar_format(info['expr'].strip())
                if norm_col_expr == norm_expr:
                    result.add_difference(col, info['full'], None)
                    is_equal = False
                    break
        
        # 두 번째 쿼리에만 있는 표현식 처리
        for norm_expr in only_in_query2:
            orig_expr = norm_expr2_map[norm_expr]
            # 이 표현식을 포함하는 컬럼 정보 찾기
            for col, info in columns2_filtered.items():
                norm_col_expr = self._normalize_tochar_format(info['expr'].strip())
                if norm_col_expr == norm_expr:
                    result.add_difference(col, None, info['full'])
                    is_equal = False
                    break
        
        return is_equal
    
    def _normalize_tochar_format(self, expr):
        """
        to_char 함수와 to_date 함수의 포맷 문자열을 정규화합니다.
        포맷 문자열의 차이를 무시하고 함수와 인자 패턴만 비교합니다.
        
        Args:
            expr (str): SQL 표현식
            
        Returns:
            str: 정규화된 표현식
        """
        # 기본 공백 정규화
        norm_expr = re.sub(r'\s+', ' ', expr).strip().lower()
        
        # 함수 내부의 공백 정규화 (특히 콤마와 따옴표 사이의 공백)
        # 콤마 다음 공백 정규화
        norm_expr = re.sub(r',\s+', ',', norm_expr)
        # 콤마 이전 공백 정규화
        norm_expr = re.sub(r'\s+,', ',', norm_expr)
        # 괄호와 인자 사이의 공백 정규화
        norm_expr = re.sub(r'\(\s+', '(', norm_expr)
        norm_expr = re.sub(r'\s+\)', ')', norm_expr)
        
        # to_char 함수의 포맷 부분 정규화
        # to_char(column, 'FORMAT') 패턴에서 'FORMAT' 부분을 일반화
        to_char_pattern = r"""
            (to_char\s*\(\s*[^,]+\s*,\s*)  # 함수 이름과 첫 인자
            (?:\'[^\']*\'|\"[^\"]*\")      # 포맷 문자열
            (\s*\))                        # 닫는 괄호
        """
        to_char_pattern = re.compile(to_char_pattern, flags=re.IGNORECASE | re.VERBOSE)
        
        # to_date 함수의 포맷 부분 정규화
        # to_date(column, 'FORMAT') 패턴에서 'FORMAT' 부분을 일반화
        to_date_pattern = r"""
            (to_date\s*\(\s*[^,]+\s*,\s*)  # 함수 이름과 첫 인자
            (?:\'[^\']*\'|\"[^\"]*\")      # 포맷 문자열
            (\s*\))                        # 닫는 괄호
        """
        to_date_pattern = re.compile(to_date_pattern, flags=re.IGNORECASE | re.VERBOSE)
        
        # to_char와 to_date 함수의 포맷 문자열을 'FORMAT'으로 일반화
        norm_expr = to_char_pattern.sub(r'\1\'FORMAT\'\2', norm_expr)
        norm_expr = to_date_pattern.sub(r'\1\'FORMAT\'\2', norm_expr)
        
        return norm_expr

    def check_special_columns(self, query: str, direction: str) -> List[str]:
        """
        특수 컬럼의 존재 여부와 값을 체크합니다.
        
        Args:
            query (str): 검사할 쿼리
            direction (str): 송신('send') 또는 수신('recv')
            
        Returns:
            List[str]: 경고 메시지 리스트
        """
        warnings = []
        
        if direction == 'send':
            columns = self.parse_select_columns(query)
        else:
            _, columns = self.parse_insert_parts(query)
            
        if not columns:
            return warnings
            
        # 대소문자 구분 없이 컬럼 비교를 위한 매핑 생성
        if direction == 'send':
            # SELECT 쿼리의 경우 새 구조에 맞게 처리
            columns_upper = {}
            for k, v in columns.items():
                # 별칭이 있는 경우 별칭을 키로 사용
                if v['alias'] is not None:
                    columns_upper[v['alias'].upper()] = (v['alias'], v)
                else:
                    # 별칭이 없는 경우 표현식을 키로 사용
                    columns_upper[k.upper()] = (k, v)
        else:
            # INSERT 쿼리는 기존과 동일하게 처리
            columns_upper = {k.upper(): (k, v) for k, v in columns.items()}
        
        # 필수 특수 컬럼 체크
        for col in self.special_columns[direction]['required']:
            if col not in columns_upper:
                warnings.append(f"필수 특수 컬럼 '{col}'이(가) {direction} 쿼리에 없습니다.")
        
        # 수신 쿼리의 특수 값 체크
        if direction == 'recv':
            for col, expected_value in self.special_columns[direction]['special_values'].items():
                if col in columns_upper:
                    col_name, col_value = columns_upper[col]
                    if col_value != expected_value:
                        warnings.append(f"특수 컬럼 '{col}'의 값이 기대값과 다릅니다. 기대값: {expected_value}, 실제값: {col_value}")
                        
        return warnings

    def clean_select_query(self, query):
        """
        Clean SELECT query by removing WHERE clause
        """
        # Find the position of WHERE (case insensitive)
        where_match = re.search(r'\sWHERE\s', query, flags=re.IGNORECASE)
        if where_match:
            # Return only the part before WHERE
            return query[:where_match.start()].strip()
        return query.strip()

    def clean_insert_query(self, query: str) -> str:
        """
        Clean INSERT query by removing PL/SQL blocks
        """
        # PL/SQL 블록에서 INSERT 문 추출
        pattern = r"""
            (?:BEGIN\s+)?          # BEGIN (optional)
            (INSERT\s+INTO\s+      # INSERT INTO
            [^;]+                  # everything until semicolon
            )                      # capture this part
            (?:\s*;)?             # optional semicolon
            (?:\s*EXCEPTION\s+     # EXCEPTION block (optional)
            .*?                    # everything until END
            END;?)?                # END with optional semicolon
        """
        insert_match = re.search(
            pattern,
            query,
            flags=re.IGNORECASE | re.MULTILINE | re.DOTALL | re.VERBOSE
        )
        
        if insert_match:
            return insert_match.group(1).strip()
        return query.strip()

    def is_meaningful_query(self, query: str) -> bool:
        """
        Check if a query is meaningful (not just a simple existence check or count)
        
        Args:
            query (str): SQL query to analyze
            
        Returns:
            bool: True if the query is meaningful, False otherwise
        """
        query = query.lower()
        
        # Remove comments and normalize whitespace
        query = re.sub(r'--.*$', '', query, flags=re.MULTILINE)
        query = ' '.join(query.split())
        
        # Patterns for meaningless queries
        meaningless_patterns = [
            r'select\s+1\s+from',  # SELECT 1 FROM ...
            r'select\s+count\s*\(\s*\*\s*\)\s+from',  # SELECT COUNT(*) FROM ...
            r'select\s+count\s*\(\s*1\s*\)\s+from',  # SELECT COUNT(1) FROM ...
            r'select\s+null\s+from',  # SELECT NULL FROM ...
            r'select\s+\'[^\']*\'\s+from',  # SELECT 'constant' FROM ...
            r'select\s+\d+\s+from',  # SELECT {number} FROM ...
        ]
        
        # Check if query matches any meaningless pattern
        for pattern in meaningless_patterns:
            if re.search(pattern, query):
                return False
                
        # For SELECT queries, check if it's selecting actual columns
        if query.startswith('select'):
            # Extract the SELECT clause (between SELECT and FROM)
            select_match = re.match(r'select\s+(.+?)\s+from', query)
            if select_match:
                select_clause = select_match.group(1)
                # If only selecting literals or simple expressions, consider it meaningless
                if re.match(r'^[\d\'\"\s,]+$', select_clause):
                    return False
        
        return True

    def find_files_by_table(self, folder_path: str, table_name: str, skip_meaningless: bool = True) -> dict:
        """
        Find files containing queries that reference the specified table
        
        Args:
            folder_path (str): Path to the folder to search in
            table_name (str): Name of the DB table to search for
            skip_meaningless (bool): If True, skip queries that appear to be meaningless
            
        Returns:
            dict: Dictionary with 'select' and 'insert' as keys, each containing a list of tuples
                 where each tuple contains (file_path, query)
        """
        import os
        
        results = {
            'select': [],
            'insert': []
        }
        
        # Normalize table name for comparison
        table_name = table_name.lower()
        
        # Create parser instance for processing files
        parser = self
        
        # Walk through all files in the folder
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                
                # Skip non-XML files silently
                if not file_path.lower().endswith('.xml'):
                    continue
                    
                try:
                    # Try to parse queries from the file
                    select_queries, insert_queries = self.parse_xml_file(file_path)
                    
                    # Check SELECT queries
                    for query in select_queries:
                        if self.extract_table_name(query).lower() == table_name:
                            # Skip meaningless queries if requested
                            if skip_meaningless and not self.is_meaningful_query(query):
                                continue
                            rel_path = os.path.relpath(file_path, folder_path)
                            results['select'].append((rel_path, query))
                    
                    # Check INSERT queries
                    for query in insert_queries:
                        if self.extract_table_name(query).lower() == table_name:
                            rel_path = os.path.relpath(file_path, folder_path)
                            results['insert'].append((rel_path, query))
                    
                except Exception:
                    # Skip any errors silently
                    continue
        
        return results

    def parse_xml_file(self, filename):
        """
        Parse XML file and extract SQL queries
        
        Args:
            filename (str): Path to the XML file
            
        Returns:
            tuple: Lists of (select_queries, insert_queries)
        """
        try:
            # Clear previous queries
            self.select_queries = []
            self.insert_queries = []
            
            # Parse XML file
            tree = ET.parse(filename)
            root = tree.getroot()
            
            # Find all text content in the XML
            for elem in root.iter():
                if elem.text:
                    text = elem.text.strip()
                    # Extract SELECT queries
                    if re.search(r'SELECT', text, flags=re.IGNORECASE):
                        cleaned_query = self.clean_select_query(text)
                        self.select_queries.append(cleaned_query)
                    # Extract INSERT queries
                    elif re.search(r'INSERT', text, flags=re.IGNORECASE):
                        cleaned_query = self.clean_insert_query(text)
                        self.insert_queries.append(cleaned_query)
            
            return self.select_queries, self.insert_queries
            
        except ET.ParseError:
            return [], []
        except Exception:
            return [], []
    
    def get_select_queries(self):
        """Return list of extracted SELECT queries"""
        return self.select_queries
    
    def get_insert_queries(self):
        """Return list of extracted INSERT queries"""
        return self.insert_queries
    
    def print_queries(self):
        """Print all extracted queries"""
        print("\nSELECT Queries:")
        for i, query in enumerate(self.select_queries, 1):
            print(f"{i}. {query}\n")
            
        print("\nINSERT Queries:")
        for i, query in enumerate(self.insert_queries, 1):
            print(f"{i}. {query}\n")

    def print_query_differences(self, diff: QueryDifference):
        """Print the differences between two queries in a readable format"""
        print(f"\nQuery Type: {diff.query_type}")
        if diff.is_equal:
            print("Queries are equivalent")
        else:
            print("Differences found:")
            for d in diff.differences:
                print(f"- Column '{d['column']}':")
                print(f"  Query 1: {d['query1_value']}")
                print(f"  Query 2: {d['query2_value']}")

    def extract_table_name(self, query: str) -> str:
        """
        Extract table name from a SQL query
        
        Args:
            query (str): SQL query to analyze
            
        Returns:
            str: Table name or empty string if not found
        """
        query = self.normalize_query(query)
        
        # For SELECT queries
        select_match = re.search(r'from\s+([a-zA-Z0-9_$.]+)', query, flags=re.IGNORECASE)
        if select_match:
            return select_match.group(1)
            
        # For INSERT queries
        insert_match = re.search(r'insert\s+into\s+([a-zA-Z0-9_$.]+)', query, flags=re.IGNORECASE)
        if insert_match:
            return insert_match.group(1)
            
        return ""

    def print_table_search_results(self, results: dict, table_name: str):
        """
        Print table search results in a formatted way
        
        Args:
            results (dict): Dictionary with search results
            table_name (str): Name of the DB table that was searched
        """
        print(f"\nFiles and queries referencing table: {table_name}")
        print("=" * 50)
        
        print("\nSELECT queries found in:")
        if results['select']:
            for i, (file, query) in enumerate(results['select'], 1):
                print(f"\n{i}. File: {file}")
                print("Query:")
                print(query)
        else:
            print("  No files found with SELECT queries")
            
        print("\nINSERT queries found in:")
        if results['insert']:
            for i, (file, query) in enumerate(results['insert'], 1):
                print(f"\n{i}. File: {file}")
                print("Query:")
                print(query)
        else:
            print("  No files found with INSERT queries")
        
        print("\n" + "=" * 50)

    def compare_mq_bw_queries(self, mq_xml_path: str, bw_xml_path: str) -> Dict[str, List[QueryDifference]]:
        """
        MQ XML과 BW XML 파일에서 추출한 송신/수신 쿼리를 비교합니다.
        
        Args:
            mq_xml_path (str): MQ XML 파일 경로
            bw_xml_path (str): BW XML 파일 경로
            
        Returns:
            Dict[str, List[QueryDifference]]: 송신/수신별 쿼리 비교 결과
                {
                    'send': [송신 쿼리 비교 결과 목록],
                    'recv': [수신 쿼리 비교 결과 목록]
                }
        """
        results = {
            'send': [],
            'recv': []
        }
        
        # MQ XML 파싱
        mq_queries = self.parse_xml_file(mq_xml_path)
        if not mq_queries:
            print(f"Failed to parse MQ XML file: {mq_xml_path}")
            return results
            
        mq_select_queries, mq_insert_queries = mq_queries
        
        # BW XML 파싱
        bw_extractor = BWQueryExtractor()
        bw_queries = bw_extractor.extract_bw_queries(bw_xml_path)
        if not bw_queries:
            print(f"Failed to parse BW XML file: {bw_xml_path}")
            return results
        
        bw_send_queries = bw_queries.get('send', [])
        bw_recv_queries = bw_queries.get('recv', [])
        
        # 송신 쿼리 비교 (SELECT)
        if mq_select_queries and bw_send_queries:
            print("\n===== 송신 쿼리 비교 (SELECT) =====")
            for mq_query in mq_select_queries:
                for bw_query in bw_send_queries:
                    diff = self.compare_queries(mq_query, bw_query)
                    if diff:
                        results['send'].append(diff)
                        self.print_query_differences(diff)
        else:
            print("송신 쿼리 비교를 위한 데이터가 부족합니다.")
            
        # 수신 쿼리 비교 (INSERT)
        if mq_insert_queries and bw_recv_queries:
            print("\n===== 수신 쿼리 비교 (INSERT) =====")
            for mq_query in mq_insert_queries:
                for bw_query in bw_recv_queries:
                    diff = self.compare_queries(mq_query, bw_query)
                    if diff:
                        results['recv'].append(diff)
                        self.print_query_differences(diff)
        else:
            print("수신 쿼리 비교를 위한 데이터가 부족합니다.")
            
        return results

    def compare_mq_bw_queries_by_interface_id(self, interface_id: str, mq_folder_path: str, bw_folder_path: str) -> Dict[str, List[QueryDifference]]:
        """
        인터페이스 ID를 기준으로 MQ XML과 BW XML 파일을 찾아 쿼리를 비교합니다.
        
        Args:
            interface_id (str): 인터페이스 ID
            mq_folder_path (str): MQ XML 파일이 있는 폴더 경로
            bw_folder_path (str): BW XML 파일이 있는 폴더 경로
            
        Returns:
            Dict[str, List[QueryDifference]]: 송신/수신별 쿼리 비교 결과
        """
        results = {
            'send': [],
            'recv': []
        }
        
        # MQ XML 파일 찾기
        searcher = FileSearcher()
        mq_files = searcher.find_files_with_keywords(mq_folder_path, [interface_id])
        
        if not mq_files or not mq_files.get(interface_id):
            print(f"No MQ XML files found for interface ID: {interface_id}")
            return results
            
        mq_xml_path = mq_files[interface_id][0] if mq_files[interface_id] else None
        if not mq_xml_path:
            print(f"No MQ XML file found for interface ID: {interface_id}")
            return results
            
        # MQ XML에서 테이블 이름 추출
        mq_queries = self.parse_xml_file(mq_xml_path)
        if not mq_queries or not mq_queries[0]:
            print(f"Failed to parse MQ XML file: {mq_xml_path}")
            return results
            
        table_name = self.extract_table_name(mq_queries[0][0]) if mq_queries[0] else None
        if not table_name:
            print(f"Failed to extract table name from MQ XML SELECT query")
            return results
            
        print(f"Found table name from MQ XML: {table_name}")
        
        # BW XML 파일 찾기
        bw_files = searcher.find_files_with_keywords(bw_folder_path, [table_name])
        
        if not bw_files or not bw_files.get(table_name):
            print(f"No BW XML files found for table name: {table_name}")
            return results
            
        bw_xml_paths = bw_files[table_name] if bw_files.get(table_name) else []
        if not bw_xml_paths:
            print(f"No BW XML files found for table name: {table_name}")
            return results
            
        # 각 BW XML 파일과 비교
        all_results = {
            'send': [],
            'recv': []
        }
        
        for bw_xml_path in bw_xml_paths:
            print(f"\nComparing MQ XML ({mq_xml_path}) with BW XML ({bw_xml_path}):")
            curr_results = self.compare_mq_bw_queries(mq_xml_path, bw_xml_path)
            
            if curr_results['send']:
                all_results['send'].extend(curr_results['send'])
                
            if curr_results['recv']:
                all_results['recv'].extend(curr_results['recv'])
                
        return all_results

class BWQueryExtractor:
    """TIBCO BW XML 파일에서 특정 태그 구조에 따라 SQL 쿼리를 추출하는 클래스"""
    
    def __init__(self):
        self.ns = {
            'pd': 'http://xmlns.tibco.com/bw/process/2003',
            'xsl': 'http://www.w3.org/1999/XSL/Transform',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
        }

    def _remove_oracle_hints(self, query: str) -> str:
        """
        SQL 쿼리에서 Oracle 힌트(/*+ ... */) 제거
        
        Args:
            query (str): 원본 SQL 쿼리
            
        Returns:
            str: 힌트가 제거된 SQL 쿼리
        """
        import re
        # /*+ ... */ 패턴의 힌트 제거
        cleaned_query = re.sub(r'/\*\+[^*]*\*/', '', query)
        # 불필요한 공백 정리 (여러 개의 공백을 하나로)
        cleaned_query = re.sub(r'\s+', ' ', cleaned_query).strip()
        
        if cleaned_query != query:
            print("\n=== Oracle 힌트 제거 ===")
            print(f"원본 쿼리: {query}")
            print(f"정리된 쿼리: {cleaned_query}")
            
        return cleaned_query

    def _get_parameter_names(self, activity) -> List[str]:
        """
        Prepared_Param_DataType에서 파라미터 이름 목록 추출
        
        Args:
            activity: JDBC 액티비티 XML 요소
            
        Returns:
            List[str]: 파라미터 이름 목록
        """
        param_names = []
        print("\n=== XML 구조 디버깅 ===")
        print("activity 태그:", activity.tag)
        print("activity의 자식 태그들:", [child.tag for child in activity])
        
        # 대소문자를 맞춰서 수정
        prepared_params = activity.find('.//Prepared_Param_DataType', self.ns)
        if prepared_params is not None:
            print("\n=== Prepared_Param_DataType 태그 발견 ===")
            print("prepared_params 태그:", prepared_params.tag)
            print("prepared_params의 자식 태그들:", [child.tag for child in prepared_params])
            
            for param in prepared_params.findall('./parameter', self.ns):
                param_name = param.find('./parameterName', self.ns)
                if param_name is not None and param_name.text:
                    name = param_name.text.strip()
                    param_names.append(name)
                    print(f"파라미터 이름 추출: {name}")
        else:
            print("\n=== Prepared_Param_DataType 태그를 찾을 수 없음 ===")
            # 전체 XML 구조를 재귀적으로 출력하여 디버깅
            def print_element_tree(element, level=0):
                print("  " * level + f"- {element.tag}")
                for child in element:
                    print_element_tree(child, level + 1)
            print("\n=== 전체 XML 구조 ===")
            print_element_tree(activity)
        
        return param_names

    def _replace_with_param_names(self, query: str, param_names: List[str]) -> str:
        """
        1단계: SQL 쿼리의 ? 플레이스홀더를 prepared_Param_DataType의 파라미터 이름으로 대체
        
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
            if i < len(parts):
                result += f":{param_name}" + parts[i+1]
                
        print("\n=== 1단계: prepared_Param_DataType 매핑 결과 ===")
        print(f"원본 쿼리: {query}")
        print(f"매핑된 쿼리: {result}")
        return result

    def _get_record_mappings(self, activity, param_names: List[str]) -> Dict[str, str]:
        """
        2단계: Record 태그에서 실제 값 매핑 정보 추출
        
        Args:
            activity: JDBC 액티비티 XML 요소
            param_names: prepared_Param_DataType에서 추출한 파라미터 이름 목록
            
        Returns:
            Dict[str, str]: 파라미터 이름과 매핑된 실제 값의 딕셔너리
        """
        mappings = {}
        # 이미 매핑된 실제 컬럼 값을 추적하는 집합
        mapped_values = set()
        
        input_bindings = activity.find('.//pd:inputBindings', self.ns)
        if input_bindings is None:
            print("\n=== inputBindings 태그를 찾을 수 없음 ===")
            return mappings

        print("\n=== Record 매핑 검색 시작 ===")
        
        # jdbcUpdateActivityInput/Record 찾기
        jdbc_input = input_bindings.find('.//jdbcUpdateActivityInput')
        if jdbc_input is None:
            print("jdbcUpdateActivityInput을 찾을 수 없음")
            return mappings

        # for-each/Record 찾기
        for_each = jdbc_input.find('.//xsl:for-each', self.ns)
        record = for_each.find('./Record') if for_each is not None else jdbc_input
        
        if record is not None:
            print("Record 태그 발견")
            # 각 파라미터 이름에 대해 매핑 찾기
            for param_name in param_names:
                print(f"\n파라미터 '{param_name}' 매핑 검색:")
                param_element = record.find(f'.//{param_name}')
                if param_element is not None:
                    # 매핑 타입별로 값을 추출하되, 중복 매핑을 방지
                    mapping_found = False
                    
                    # value-of 체크 (우선 순위 1)
                    value_of = param_element.find('.//xsl:value-of', self.ns)
                    if value_of is not None:
                        select_attr = value_of.get('select', '')
                        if select_attr:
                            # select="BANANA"와 같은 형식에서 실제 값 추출
                            value = select_attr.split('/')[-1]
                            mappings[param_name] = value
                            print(f"value-of 매핑 발견: {param_name} -> {value}")
                            mapping_found = True
                    
                    # choose/when 체크 (우선 순위 2, value-of가 없을 경우만)
                    if not mapping_found:
                        choose = param_element.find('.//xsl:choose', self.ns)
                        if choose is not None:
                            when = choose.find('.//xsl:when', self.ns)
                            if when is not None:
                                test_attr = when.get('test', '')
                                if 'exists(' in test_attr:
                                    # exists(BANANA)와 같은 형식에서 변수 이름 추출
                                    value = test_attr[test_attr.find('(')+1:test_attr.find(')')]
                                    mappings[param_name] = value
                                    print(f"choose/when 매핑 발견: {param_name} -> {value}")
                else:
                    print(f"'{param_name}'에 대한 매핑을 찾을 수 없음")

        return mappings

    def _replace_with_actual_values(self, query: str, mappings: Dict[str, str]) -> str:
        """
        2단계: 파라미터 이름을 Record에서 찾은 실제 값으로 대체
        
        Args:
            query (str): 1단계에서 파라미터 이름이 대체된 쿼리
            mappings (Dict[str, str]): 파라미터 이름과 실제 값의 매핑
            
        Returns:
            str: 실제 값이 대체된 SQL 쿼리
        """
        # 순차적 치환 문제 해결을 위해 모든 대체를 한 번에 수행
        # 1. 대체될 모든 패턴을 고유한 임시 패턴으로 먼저 변환 (충돌 방지)
        result = query
        temp_replacements = {}
        
        import re
        
        for i, (param_name, actual_value) in enumerate(mappings.items()):
            # 고유한 임시 패턴 생성 (절대 원본 쿼리에 존재할 수 없는 패턴)
            temp_pattern = f"__TEMP_PLACEHOLDER_{i}__"
            
            # 정규 표현식을 사용하여 정확한 파라미터 이름만 대체
            # 단어 경계(\b)를 사용하여 정확한 파라미터 이름만 매칭
            result = re.sub(f":{param_name}\\b", temp_pattern, result)
            
            # 임시 패턴을 최종 값으로 매핑
            temp_replacements[temp_pattern] = f":{actual_value}"
        
        # 2. 모든 임시 패턴을 최종 값으로 한 번에 변환
        for temp_pattern, final_value in temp_replacements.items():
            result = result.replace(temp_pattern, final_value)
            
        print("\n=== 2단계: Record 매핑 결과 ===")
        print(f"1단계 쿼리: {query}")
        print(f"최종 쿼리: {result}")
        return result

    def extract_recv_query(self, xml_path: str) -> List[Tuple[str, str, str]]:
        """
        수신용 XML에서 SQL 쿼리와 파라미터가 매핑된 쿼리를 추출
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            List[Tuple[str, str, str]]: (원본 쿼리, 1단계 매핑 쿼리, 2단계 매핑 쿼리) 목록
        """
        queries = []
        try:
            print(f"\n=== XML 파일 처리 시작: {xml_path} ===")
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # JDBC 액티비티 찾기
            activities = root.findall('.//pd:activity', self.ns)
            
            for activity in activities:
                # JDBC 액티비티 타입 확인
                activity_type = activity.find('./pd:type', self.ns)
                if activity_type is None or 'jdbc' not in activity_type.text.lower():
                    continue
                    
                print(f"\nJDBC 액티비티 발견: {activity.get('name', 'Unknown')}")
                
                # statement 추출
                statement = activity.find('.//config/statement')
                if statement is not None and statement.text:
                    query = statement.text.strip()
                    print(f"\n발견된 쿼리:\n{query}")
                    
                    # SELECT 쿼리인 경우
                    if query.lower().startswith('select'):
                        # FROM DUAL 쿼리 제외
                        if not self._is_valid_query(query):
                            print("=> FROM DUAL 쿼리이므로 제외")
                            continue
                        # Oracle 힌트 제거
                        query = self._remove_oracle_hints(query)
                        print(f"=> Oracle 힌트 제거 후 쿼리:\n{query}")
                        queries.append((query, query, query))  # SELECT는 파라미터 매핑 없음
                    
                    # INSERT, UPDATE, DELETE 쿼리인 경우
                    elif query.lower().startswith(('insert', 'update', 'delete')):
                        # 1단계: prepared_Param_DataType의 파라미터 이름으로 매핑
                        param_names = self._get_parameter_names(activity)
                        stage1_query = self._replace_with_param_names(query, param_names)
                        
                        # 2단계: Record의 실제 값으로 매핑
                        mappings = self._get_record_mappings(activity, param_names)
                        stage2_query = self._replace_with_actual_values(stage1_query, mappings)
                        
                        queries.append((query, stage1_query, stage2_query))
                        print(f"=> 최종 처리된 쿼리:\n{stage2_query}")
            
            print(f"\n=== 처리된 유효한 쿼리 수: {len(queries)} ===")
            
        except ET.ParseError as e:
            print(f"\n=== XML 파싱 오류: {e} ===")
        except Exception as e:
            print(f"\n=== 쿼리 추출 중 오류 발생: {e} ===")
            
        return queries

    def _is_valid_query(self, query: str) -> bool:
        """
        분석 대상이 되는 유효한 쿼리인지 확인
        
        Args:
            query (str): SQL 쿼리
            
        Returns:
            bool: 유효한 쿼리이면 True
        """
        # 소문자로 변환하여 검사
        query_lower = query.lower()
        
        # SELECT FROM DUAL 패턴 체크
        if query_lower.startswith('select') and 'from dual' in query_lower:
            print(f"\n=== 단순 쿼리 제외 ===")
            print(f"제외된 쿼리: {query}")
            return False
            
        return True

    def extract_send_query(self, xml_path: str) -> List[str]:
        """
        송신용 XML에서 SQL 쿼리 추출
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            List[str]: SQL 쿼리 목록
        """
        queries = []
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # 송신 쿼리 추출 (Group 내의 SelectP 활동)
            select_activities = root.findall('.//pd:group[@name="Group"]//pd:activity[@name="SelectP"]', self.ns)
            
            print(f"\n=== 송신용 XML 처리 시작: {xml_path} ===")
            print(f"발견된 SelectP 활동 수: {len(select_activities)}")
            
            for activity in select_activities:
                statement = activity.find('.//config/statement')
                if statement is not None and statement.text:
                    query = statement.text.strip()
                    print(f"\n발견된 쿼리:\n{query}")
                    
                    # 1. 유효한 쿼리인지 먼저 확인
                    if not self._is_valid_query(query):
                        print("=> FROM DUAL 쿼리이므로 제외")
                        continue
                        
                    # 2. 유효한 쿼리에 대해서만 Oracle 힌트 제거
                    cleaned_query = self._remove_oracle_hints(query)
                    print(f"=> 최종 처리된 쿼리:\n{cleaned_query}")
                    queries.append(cleaned_query)
            
            print(f"\n=== 처리된 유효한 쿼리 수: {len(queries)} ===")
            
        except ET.ParseError as e:
            print(f"XML 파싱 오류: {e}")
        except Exception as e:
            print(f"쿼리 추출 중 오류 발생: {e}")
            
        return queries

    def extract_bw_queries(self, xml_path: str) -> Dict[str, List[str]]:
        """
        TIBCO BW XML 파일에서 송신/수신 쿼리를 모두 추출
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            Dict[str, List[str]]: 송신/수신 쿼리 목록
                {
                    'send': [select 쿼리 목록],
                    'recv': [insert 쿼리 목록]
                }
        """
        # 모든 쿼리를 추출
        send_queries = self.extract_send_query(xml_path)
        recv_queries_full = self.extract_recv_query(xml_path)
        
        # 수신 쿼리 중 INSERT 문만 필터링
        recv_queries = []
        for orig_query, _, mapped_query in recv_queries_full:
            if orig_query.lower().startswith('insert'):
                recv_queries.append(mapped_query)
        
        return {
            'send': [query for query in send_queries if query.lower().startswith('select')],
            'recv': recv_queries
        }
    def get_single_query(self, xml_path: str) -> str:
        """
        BW XML 파일에서 SQL 쿼리를 추출하여 단일 문자열로 반환
        송신(send)과 수신(recv) 쿼리 중 존재하는 것을 반환
        둘 다 없는 경우 빈 문자열 반환
        
        Args:
            xml_path (str): XML 파일 경로
            
        Returns:
            str: 추출된 SQL 쿼리 문자열. 쿼리가 없으면 빈 문자열
        """
        try:
            # 기존 extract_bw_queries 메소드 활용
            queries = self.extract_bw_queries(xml_path)
            
            # 송신 쿼리 확인
            if queries.get('send') and len(queries['send']) > 0:
                return queries['send'][0]  # 첫 번째 송신 쿼리 반환
                
            # 수신 쿼리 확인
            if queries.get('recv') and len(queries['recv']) > 0:
                return queries['recv'][0]  # 첫 번째 수신 쿼리 반환
                
            # 쿼리가 없는 경우
            return ""
            
        except Exception as e:
            print(f"쿼리 추출 중 오류 발생: {e}")
            return ""  # 오류 발생 시 빈 문자열 반환        

class FileSearcher:
    @staticmethod
    def find_files_with_keywords(folder_path: str, keywords: list) -> dict:
        """
        Search for files in the given folder that contain any of the specified keywords
        
        Args:
            folder_path (str): Path to the folder to search in
            keywords (list): List of keywords to search for
            
        Returns:
            dict: Dictionary with keyword as key and list of matching files as value
        """
        import os
        
        # Initialize results dictionary
        results = {keyword: [] for keyword in keywords}
        
        # Walk through all files in the folder
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                
                try:
                    # Try to read file content
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                        
                    # Check for each keyword
                    for keyword in keywords:
                        if keyword in content:
                            # Store relative path instead of full path
                            rel_path = os.path.relpath(file_path, folder_path)
                            results[keyword].append(rel_path)
                            
                except (UnicodeDecodeError, IOError):
                    # Skip files that can't be read as text
                    continue
        
        return results

    @staticmethod
    def print_search_results(results: dict):
        """
        Print search results in a formatted way
        
        Args:
            results (dict): Dictionary with keyword as key and list of matching files as value
        """
        print("\nSearch Results:")
        print("=" * 50)
        
        for keyword, files in results.items():
            print(f"\nKeyword: {keyword}")
            if files:
                print("Found in files:")
                for i, file in enumerate(files, 1):
                    print(f"  {i}. {file}")
            else:
                print("No files found containing this keyword")
        
        print("\n" + "=" * 50)

# Test the query comparison
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Compare SQL queries in MQ and BW XML files")
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    
    # 테이블 검색 명령
    table_parser = subparsers.add_parser("find_table", help="Find files containing a specific table")
    table_parser.add_argument("folder_path", help="Folder path to search in")
    table_parser.add_argument("table_name", help="Table name to search for")
    
    # 쿼리 비교 명령
    compare_parser = subparsers.add_parser("compare", help="Compare MQ and BW queries")
    compare_parser.add_argument("mq_xml", help="MQ XML file path")
    compare_parser.add_argument("bw_xml", help="BW XML file path")
    
    # 인터페이스 ID로 비교 명령
    interface_parser = subparsers.add_parser("compare_by_id", help="Compare by interface ID")
    interface_parser.add_argument("interface_id", help="Interface ID")
    interface_parser.add_argument("mq_folder", help="MQ XML folder path")
    interface_parser.add_argument("bw_folder", help="BW XML folder path")
    
    args = parser.parse_args()
    
    query_parser = QueryParser()
    
    if args.command == "find_table":
        table_results = query_parser.find_files_by_table(args.folder_path, args.table_name)
        query_parser.print_table_search_results(table_results, args.table_name)
    elif args.command == "compare":
        comparison_results = query_parser.compare_mq_bw_queries(args.mq_xml, args.bw_xml)
        print("\nComparison Results Summary:")
        print(f"송신 쿼리 비교 결과: {len(comparison_results['send'])} 개의 차이점 발견")
        print(f"수신 쿼리 비교 결과: {len(comparison_results['recv'])} 개의 차이점 발견")
    elif args.command == "compare_by_id":
        comparison_results = query_parser.compare_mq_bw_queries_by_interface_id(
            args.interface_id, args.mq_folder, args.bw_folder
        )
        print("\nComparison Results Summary:")
        print(f"송신 쿼리 비교 결과: {len(comparison_results['send'])} 개의 차이점 발견")
        print(f"수신 쿼리 비교 결과: {len(comparison_results['recv'])} 개의 차이점 발견")
    else:
        parser.print_help()