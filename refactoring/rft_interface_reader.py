"""
인터페이스 정보 읽기 모듈

Excel 파일에서 인터페이스 정보를 읽어 파이썬 자료구조로 변환하고,
TIBCO BW .process 파일에서 수신용 INSERT 쿼리를 추출합니다.
"""

import os
import ast
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any
import pandas as pd
import datetime


class InterfaceExcelReader:
    """
    인터페이스 정보가 담긴 Excel 파일을 읽어 파이썬 자료구조로 변환하는 클래스
    
    Excel 파일 구조:
    - B열부터 3컬럼 단위로 하나의 인터페이스 블록
    - 1행: 인터페이스명
    - 2행: 인터페이스ID  
    - 3행: DB 연결 정보 (문자열로 저장된 딕셔너리)
    - 4행: 테이블 정보 (문자열로 저장된 딕셔너리)
    - 5행부터: 컬럼 매핑 정보
    """
    
    def __init__(self, replacer_excel_path: Optional[str] = None):
        """
        InterfaceExcelReader 클래스 초기화
        
        Args:
            replacer_excel_path: string_replacer용 Excel 파일 경로
        """
        self.processed_count = 0
        self.error_count = 0
        self.replacer_excel_path = replacer_excel_path or 'rft_interface_processed.csv'
        self.interfaces = {}
        
    def read_excel(self, file_path: str) -> Dict[str, Dict[str, Any]]:
        """
        Excel 파일에서 인터페이스 정보를 읽어 딕셔너리로 반환
        
        Args:
            file_path: Excel 파일 경로
            
        Returns:
            인터페이스 정보를 담은 딕셔너리
        """
        try:
            print(f"Excel 파일 읽기 시작: {file_path}")
            
            if not os.path.exists(file_path):
                print(f"오류: Excel 파일을 찾을 수 없습니다 - {file_path}")
                return {}
            
            # Excel 파일 읽기
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file_path, engine='openpyxl', header=None)
            
            print(f"Excel 데이터 로드 완료: {len(df)}행 x {len(df.columns)}열")
            
            interfaces = {}
            
            # B열부터 3컬럼씩 처리 (B=1, C=2, D=3이면 1,4,7,10... 순서)
            col_start = 1  # B열
            interface_count = 0
            
            while col_start < len(df.columns):
                try:
                    # 3컬럼 추출
                    if col_start + 2 >= len(df.columns):
                        break
                    
                    # 인터페이스 기본 정보 추출
                    interface_name = df.iloc[0, col_start] if pd.notna(df.iloc[0, col_start]) else f"Interface_{interface_count+1}"
                    interface_id = df.iloc[1, col_start] if pd.notna(df.iloc[1, col_start]) else f"ID_{interface_count+1}"
                    
                    # 빈 인터페이스명이면 건너뛰기
                    if not str(interface_name).strip() or str(interface_name).strip().lower() in ['nan', '']:
                        col_start += 3
                        continue
                    
                    print(f"인터페이스 처리 중: {interface_name} (ID: {interface_id})")
                    
                    # DB 연결 정보 파싱
                    db_info_str = df.iloc[2, col_start] if pd.notna(df.iloc[2, col_start]) else "{}"
                    try:
                        db_info = ast.literal_eval(str(db_info_str)) if str(db_info_str).strip() else {}
                    except:
                        db_info = {}
                    
                    # 테이블 정보 파싱
                    table_info_str = df.iloc[3, col_start] if pd.notna(df.iloc[3, col_start]) else "{}"
                    try:
                        table_info = ast.literal_eval(str(table_info_str)) if str(table_info_str).strip() else {}
                    except:
                        table_info = {}
                    
                    # 컬럼 매핑 정보 추출 (5행부터)
                    column_mappings = []
                    for row_idx in range(4, len(df)):  # 5행부터 (0-based이므로 4부터)
                        source_col = df.iloc[row_idx, col_start] if pd.notna(df.iloc[row_idx, col_start]) else ""
                        target_col = df.iloc[row_idx, col_start + 1] if pd.notna(df.iloc[row_idx, col_start + 1]) else ""
                        data_type = df.iloc[row_idx, col_start + 2] if pd.notna(df.iloc[row_idx, col_start + 2]) else ""
                        
                        if str(source_col).strip() and str(target_col).strip():
                            column_mappings.append({
                                'source': str(source_col).strip(),
                                'target': str(target_col).strip(),
                                'type': str(data_type).strip()
                            })
                    
                    # 인터페이스 정보 저장
                    interfaces[interface_name] = {
                        'id': interface_id,
                        'db_info': db_info,
                        'table_info': table_info,
                        'column_mappings': column_mappings,
                        'column_count': len(column_mappings)
                    }
                    
                    interface_count += 1
                    self.processed_count += 1
                    
                except Exception as e:
                    print(f"인터페이스 처리 중 오류 발생 (컬럼 {col_start}): {str(e)}")
                    self.error_count += 1
                
                # 다음 인터페이스로 이동 (3컬럼씩)
                col_start += 3
            
            print(f"인터페이스 정보 읽기 완료: {interface_count}개 처리됨")
            if self.error_count > 0:
                print(f"오류 발생: {self.error_count}건")
            
            self.interfaces = interfaces
            return interfaces
            
        except Exception as e:
            print(f"Excel 파일 읽기 중 오류 발생: {str(e)}")
            self.error_count += 1
            return {}
    
    def get_interface_summary(self) -> Dict[str, Any]:
        """
        인터페이스 정보 요약을 반환
        
        Returns:
            인터페이스 요약 정보
        """
        total_interfaces = len(self.interfaces)
        total_columns = sum(info['column_count'] for info in self.interfaces.values())
        
        summary = {
            'total_interfaces': total_interfaces,
            'total_columns': total_columns,
            'processed_count': self.processed_count,
            'error_count': self.error_count,
            'interfaces': list(self.interfaces.keys())
        }
        
        return summary
    
    def export_to_csv(self, output_path: str) -> bool:
        """
        인터페이스 정보를 CSV 파일로 내보내기
        
        Args:
            output_path: 출력 CSV 파일 경로
            
        Returns:
            내보내기 성공 여부
        """
        try:
            if not self.interfaces:
                print("내보낼 인터페이스 정보가 없습니다.")
                return False
            
            # CSV용 데이터 준비
            rows = []
            for interface_name, info in self.interfaces.items():
                base_row = {
                    'interface_name': interface_name,
                    'interface_id': info['id'],
                    'db_info': str(info['db_info']),
                    'table_info': str(info['table_info']),
                    'column_count': info['column_count']
                }
                
                # 컬럼 매핑 정보 추가
                for idx, mapping in enumerate(info['column_mappings']):
                    row = base_row.copy()
                    row.update({
                        'mapping_index': idx + 1,
                        'source_column': mapping['source'],
                        'target_column': mapping['target'],
                        'data_type': mapping['type']
                    })
                    rows.append(row)
                
                # 컬럼 매핑이 없는 경우에도 기본 정보는 저장
                if not info['column_mappings']:
                    base_row.update({
                        'mapping_index': 0,
                        'source_column': '',
                        'target_column': '',
                        'data_type': ''
                    })
                    rows.append(base_row)
            
            # DataFrame 생성 및 저장
            df = pd.DataFrame(rows)
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            
            print(f"인터페이스 정보가 CSV로 저장되었습니다: {output_path}")
            print(f"총 {len(rows)}개의 행이 저장되었습니다.")
            
            return True
            
        except Exception as e:
            print(f"CSV 내보내기 중 오류 발생: {str(e)}")
            return False


class BWProcessFileParser:
    """
    TIBCO BW .process 파일에서 수신용 INSERT 쿼리를 추출하고 파라미터 매핑을 처리하는 클래스
    """
    
    def __init__(self):
        """BWProcessFileParser 초기화"""
        self.namespace = {
            'pd': 'http://xmlns.tibco.com/bw/process/2003',
            'xsd': 'http://www.w3.org/2001/XMLSchema'
        }
        self.parsed_files = {}
        self.error_files = []
    
    def parse_process_file(self, file_path: str) -> Dict[str, Any]:
        """
        BW .process 파일을 파싱하여 INSERT 쿼리와 파라미터 정보를 추출
        
        Args:
            file_path: .process 파일 경로
            
        Returns:
            파싱된 정보 딕셔너리
        """
        try:
            print(f"BW 프로세스 파일 파싱 시작: {file_path}")
            
            if not os.path.exists(file_path):
                print(f"오류: 프로세스 파일을 찾을 수 없습니다 - {file_path}")
                return {}
            
            # XML 파일 파싱
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            result = {
                'file_path': file_path,
                'file_name': os.path.basename(file_path),
                'insert_queries': [],
                'parameters': [],
                'activities': [],
                'parse_time': datetime.datetime.now().isoformat()
            }
            
            # INSERT 쿼리 추출
            insert_activities = root.findall('.//pd:activity[@name]', self.namespace)
            for activity in insert_activities:
                activity_name = activity.get('name', '')
                
                # SQL 관련 활동 찾기
                if 'insert' in activity_name.lower() or 'sql' in activity_name.lower():
                    # SQL 쿼리 추출
                    sql_elements = activity.findall('.//sql', self.namespace)
                    for sql_elem in sql_elements:
                        if sql_elem.text and 'INSERT' in sql_elem.text.upper():
                            result['insert_queries'].append({
                                'activity_name': activity_name,
                                'sql': sql_elem.text.strip(),
                                'line_count': len(sql_elem.text.strip().split('\n'))
                            })
                
                # 활동 정보 저장
                result['activities'].append({
                    'name': activity_name,
                    'type': activity.get('type', 'unknown')
                })
            
            # 파라미터 정보 추출
            param_elements = root.findall('.//pd:parameter', self.namespace)
            for param in param_elements:
                param_name = param.get('name', '')
                param_type = param.get('type', '')
                if param_name:
                    result['parameters'].append({
                        'name': param_name,
                        'type': param_type,
                        'xpath': self._get_element_xpath(param)
                    })
            
            # 결과 저장
            self.parsed_files[file_path] = result
            
            print(f"파싱 완료: INSERT 쿼리 {len(result['insert_queries'])}개, 파라미터 {len(result['parameters'])}개")
            
            return result
            
        except Exception as e:
            print(f"BW 프로세스 파일 파싱 중 오류 발생: {str(e)}")
            self.error_files.append(file_path)
            return {}
    
    def _get_element_xpath(self, element) -> str:
        """XML 요소의 XPath를 생성"""
        try:
            path_parts = []
            current = element
            while current is not None:
                tag = current.tag
                if '}' in tag:
                    tag = tag.split('}')[1]  # 네임스페이스 제거
                path_parts.insert(0, tag)
                current = current.getparent()
            return '/' + '/'.join(path_parts)
        except:
            return 'unknown'
    
    def parse_multiple_files(self, file_paths: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        여러 BW 프로세스 파일을 일괄 파싱
        
        Args:
            file_paths: 파싱할 파일 경로 목록
            
        Returns:
            파싱된 결과 딕셔너리
        """
        results = {}
        
        print(f"총 {len(file_paths)}개의 파일을 파싱합니다.")
        
        for i, file_path in enumerate(file_paths, 1):
            print(f"\n[{i}/{len(file_paths)}] 파싱 중: {os.path.basename(file_path)}")
            result = self.parse_process_file(file_path)
            if result:
                results[file_path] = result
        
        print(f"\n파싱 완료: 성공 {len(results)}개, 실패 {len(self.error_files)}개")
        
        return results
    
    def export_parsing_results(self, output_path: str) -> bool:
        """
        파싱 결과를 CSV 파일로 내보내기
        
        Args:
            output_path: 출력 CSV 파일 경로
            
        Returns:
            내보내기 성공 여부
        """
        try:
            if not self.parsed_files:
                print("내보낼 파싱 결과가 없습니다.")
                return False
            
            rows = []
            for file_path, info in self.parsed_files.items():
                base_row = {
                    'file_path': file_path,
                    'file_name': info['file_name'],
                    'parse_time': info['parse_time'],
                    'insert_query_count': len(info['insert_queries']),
                    'parameter_count': len(info['parameters']),
                    'activity_count': len(info['activities'])
                }
                
                # INSERT 쿼리 정보 추가
                for idx, query in enumerate(info['insert_queries']):
                    row = base_row.copy()
                    row.update({
                        'query_index': idx + 1,
                        'activity_name': query['activity_name'],
                        'sql_query': query['sql'],
                        'sql_line_count': query['line_count']
                    })
                    rows.append(row)
                
                # 쿼리가 없는 경우에도 기본 정보는 저장
                if not info['insert_queries']:
                    base_row.update({
                        'query_index': 0,
                        'activity_name': '',
                        'sql_query': '',
                        'sql_line_count': 0
                    })
                    rows.append(base_row)
            
            # DataFrame 생성 및 저장
            df = pd.DataFrame(rows)
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            
            print(f"파싱 결과가 CSV로 저장되었습니다: {output_path}")
            print(f"총 {len(rows)}개의 행이 저장되었습니다.")
            
            return True
            
        except Exception as e:
            print(f"파싱 결과 내보내기 중 오류 발생: {str(e)}")
            return False


def parse_bw_receive_file(file_path: str) -> Dict[str, Any]:
    """
    BW 수신파일 파싱을 위한 편의 함수
    
    Args:
        file_path: BW .process 파일 경로
        
    Returns:
        파싱된 정보
    """
    parser = BWProcessFileParser()
    return parser.parse_process_file(file_path)


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("인터페이스 정보 읽기 도구")
    print("=" * 60)
    
    excel_reader = InterfaceExcelReader()
    bw_parser = BWProcessFileParser()
    
    while True:
        print("\n메뉴:")
        print("1. Excel에서 인터페이스 정보 읽기")
        print("2. BW 프로세스 파일 파싱")
        print("3. 인터페이스 요약 정보 보기")
        print("4. 결과를 CSV로 내보내기")
        print("0. 종료")
        
        choice = input("\n선택하세요: ").strip()
        
        if choice == "1":
            excel_path = input("Excel 파일 경로를 입력하세요: ").strip()
            if excel_path:
                interfaces = excel_reader.read_excel(excel_path)
                if interfaces:
                    print(f"\n{len(interfaces)}개의 인터페이스가 로드되었습니다.")
                    for name in list(interfaces.keys())[:5]:  # 처음 5개만 표시
                        print(f"  - {name}")
                    if len(interfaces) > 5:
                        print(f"  ... 외 {len(interfaces) - 5}개")
            else:
                print("파일 경로를 입력해야 합니다.")
                
        elif choice == "2":
            process_path = input("BW 프로세스 파일 경로를 입력하세요: ").strip()
            if process_path:
                result = bw_parser.parse_process_file(process_path)
                if result:
                    print(f"\n파싱 완료:")
                    print(f"  - INSERT 쿼리: {len(result['insert_queries'])}개")
                    print(f"  - 파라미터: {len(result['parameters'])}개")
                    print(f"  - 활동: {len(result['activities'])}개")
            else:
                print("파일 경로를 입력해야 합니다.")
                
        elif choice == "3":
            summary = excel_reader.get_interface_summary()
            print(f"\n인터페이스 요약:")
            print(f"  - 총 인터페이스 수: {summary['total_interfaces']}")
            print(f"  - 총 컬럼 수: {summary['total_columns']}")
            print(f"  - 처리 성공: {summary['processed_count']}")
            print(f"  - 오류 발생: {summary['error_count']}")
            
        elif choice == "4":
            sub_choice = input("1. 인터페이스 정보 내보내기, 2. BW 파싱 결과 내보내기: ").strip()
            
            if sub_choice == "1":
                output_path = input("출력 CSV 파일 경로 (Enter: 기본값): ").strip()
                if not output_path:
                    output_path = "rft_interface_export.csv"
                excel_reader.export_to_csv(output_path)
                
            elif sub_choice == "2":
                output_path = input("출력 CSV 파일 경로 (Enter: 기본값): ").strip()
                if not output_path:
                    output_path = "rft_bw_parsing_export.csv"
                bw_parser.export_parsing_results(output_path)
                
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
            
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")


if __name__ == "__main__":
    main()