"""
YAML 처리 모듈

Excel 파일을 읽어 YAML 파일을 생성하고, 생성된 YAML을 기반으로 파일 복사 및 치환 작업을 수행합니다.
"""

import os
import re
import datetime
import pandas as pd
import yaml
import shutil
from typing import Dict, List, Optional, Any


class YAMLProcessor:
    """YAML 파일 생성 및 치환 작업을 처리하는 클래스"""
    
    def __init__(self, debug_mode: bool = True):
        """
        YAMLProcessor 초기화
        
        Args:
            debug_mode: 디버그 모드 활성화 여부
        """
        self.debug_mode = debug_mode
        
    def debug_print(self, *args, **kwargs):
        """디버그 모드일 때만 메시지를 출력"""
        if self.debug_mode:
            print("[DEBUG]", *args, **kwargs)
    
    def generate_yaml_from_excel(self, excel_path: str, yaml_path: str) -> bool:
        """
        Excel 파일을 읽어 YAML 파일을 생성
        
        Args:
            excel_path: 입력 Excel 파일 경로
            yaml_path: 출력 YAML 파일 경로
            
        Returns:
            생성 성공 여부
        """
        try:
            print(f"Excel 파일 읽기 시작: {excel_path}")
            
            # 파일 존재 확인
            if not os.path.exists(excel_path):
                print(f"오류: Excel 파일을 찾을 수 없습니다 - {excel_path}")
                return False
            
            # Excel 파일 읽기
            if excel_path.endswith('.csv'):
                df = pd.read_csv(excel_path, encoding='utf-8-sig')
            else:
                df = pd.read_excel(excel_path, engine='openpyxl')
            
            print(f"Excel 데이터 로드 완료: {len(df)}행 x {len(df.columns)}열")
            
            # 전체 YAML 구조를 저장할 딕셔너리
            full_yaml_structure = {}
            
            # 2행씩 처리 (일반행, 매칭행)
            for i in range(0, len(df), 2):
                if i + 1 >= len(df):
                    break
                
                normal_row = df.iloc[i]
                match_row = df.iloc[i+1]
                
                print(f"\n=== {i//2 + 1}번째 행 쌍 처리 ===")
                
                yaml_structure = {
                    f"{i//2 + 1}번째 행": {}
                }
                
                # 파일 경로 처리 함수들
                def modify_path(path):
                    """파일 경로를 수정하는 함수"""
                    if isinstance(path, str) and path.startswith("C:\\BwProject"):
                        return path.replace("C:\\BwProject", "C:\\TBwProject")
                    return path
                
                def extract_filename(path):
                    """파일 경로에서 파일명만 추출"""
                    if not isinstance(path, str):
                        return ""
                    return os.path.basename(path)
                
                def process_schema_path(schema_path):
                    """스키마 파일 경로를 처리하여 namespace와 schemaLocation 생성"""
                    if not isinstance(schema_path, str):
                        return None, None
                    
                    normalized_path = schema_path.replace('\\', '/')
                    shared_idx = normalized_path.find('/SharedResources')
                    if shared_idx == -1:
                        return None, None
                    
                    bb_start = normalized_path.rfind('/', 0, shared_idx)
                    if bb_start == -1:
                        return None, None
                    
                    relative_path = normalized_path[bb_start:]
                    schema_location = relative_path[relative_path.find('/SharedResources'):]
                    namespace = f"http://www.tibco.com/schemas{relative_path}"
                    
                    return namespace, schema_location
                
                def create_schema_replacements(filename, schema_path):
                    """스키마 파일 치환 목록 생성"""
                    if not filename.endswith('.xsd'):
                        return []
                    
                    namespace, schema_location = process_schema_path(schema_path)
                    if not namespace or not schema_location:
                        return []
                    
                    base_name = os.path.splitext(filename)[0]
                    return [{
                        "설명": "스키마 namespace 치환",
                        "찾기": {
                            "정규식": f'namespace\\s*=\\s*"[^"]*{base_name}[^"]*"'
                        },
                        "교체": {
                            "값": f'namespace="{namespace}"'
                        }
                    },
                    {
                        "설명": "스키마 schemaLocation 치환",
                        "찾기": {
                            "정규식": f'schemaLocation\\s*=\\s*"[^"]*{base_name}[^"]*"'
                        },
                        "교체": {
                            "값": f'schemaLocation="{schema_location}"'
                        }
                    }]
                
                def extract_process_path(file_path):
                    """프로세스 파일 경로에서 'Processes' 이후의 경로를 추출"""
                    if not isinstance(file_path, str):
                        return ""
                    
                    normalized_path = file_path.replace('\\', '/')
                    processes_idx = normalized_path.find('Processes/')
                    if processes_idx == -1:
                        return ""
                    
                    return normalized_path[processes_idx + len('Processes/'):]
                
                def create_process_replacements(source_path, target_path, match_row, normal_row):
                    """프로세스 파일의 치환 목록 생성"""
                    if not isinstance(source_path, str) or not isinstance(target_path, str):
                        return []
                    
                    source_filename = extract_filename(source_path)
                    target_process_path = extract_process_path(target_path)
                    
                    if not source_filename or not target_process_path:
                        return []
                    
                    replacements = [{
                        "설명": "프로세스 이름 치환",
                        "찾기": {
                            "정규식": f'<pd:name>Processes/[^<]*</pd:name>'
                        },
                        "교체": {
                            "값": f'<pd:name>Processes/{target_process_path}</pd:name>'
                        }
                    }]
                    
                    # 고정 문자열 치환 규칙들
                    fixed_replacements = [
                        {
                            "설명": "LHMES_MGR 치환",
                            "찾기": {"정규식": "LHMES_MGR"},
                            "교체": {"값": "LYMES_MGR"}
                        },
                        {
                            "설명": "VOMES_MGR 치환",
                            "찾기": {"정규식": "VOMES_MGR"},
                            "교체": {"값": "LZMES_MGR"}
                        },
                        {
                            "설명": "LH 문자열 치환",
                            "찾기": {"정규식": "'LH'"},
                            "교체": {"값": "'LY'"}
                        },
                        {
                            "설명": "VO 문자열 치환",
                            "찾기": {"정규식": "'VO'"},
                            "교체": {"값": "'LZ'"}
                        }
                    ]
                    replacements.extend(fixed_replacements)
                    
                    # 동적 치환 규칙들
                    try:
                        if 'Group ID' in match_row.index and 'Event_ID' in match_row.index:
                            origin_ifid = f"{match_row['Group ID']}.{match_row['Event_ID']}"
                            dest_ifid = f"{normal_row['Group ID']}.{match_row['Event_ID']}"
                            
                            if origin_ifid != dest_ifid:
                                replacements.append({
                                    "설명": "IFID 치환",
                                    "찾기": {"정규식": origin_ifid.replace(".", "\\.")},
                                    "교체": {"값": dest_ifid}
                                })
                    except:
                        pass
                    
                    return replacements
                
                # 1. 송신파일경로 처리
                if (pd.notna(normal_row.get('송신파일생성여부')) and 
                    str(normal_row.get('송신파일생성여부')).strip() == '1'):
                    
                    yaml_structure[f"{i//2 + 1}번째 행"]["송신파일경로"] = {
                        "원본파일": match_row.get('송신파일경로', ''),
                        "복사파일": modify_path(normal_row.get('송신파일경로', '')),
                        "치환목록": create_schema_replacements(
                            extract_filename(normal_row.get('송신스키마파일명', '')),
                            normal_row.get('송신스키마파일명', '')
                        ) + create_process_replacements(
                            match_row.get('송신파일경로', ''),
                            normal_row.get('송신파일경로', ''),
                            match_row,
                            normal_row
                        )
                    }
                    print(f"  송신파일경로 생성: {match_row.get('송신파일경로', '')} -> {modify_path(normal_row.get('송신파일경로', ''))}")
                
                # 2. 수신파일경로 처리
                if (pd.notna(normal_row.get('수신파일생성여부')) and 
                    str(normal_row.get('수신파일생성여부')).strip() == '1'):
                    
                    yaml_structure[f"{i//2 + 1}번째 행"]["수신파일경로"] = {
                        "원본파일": match_row.get('수신파일경로', ''),
                        "복사파일": modify_path(normal_row.get('수신파일경로', '')),
                        "치환목록": create_schema_replacements(
                            extract_filename(normal_row.get('수신스키마파일명', '')),
                            normal_row.get('수신스키마파일명', '')
                        ) + create_process_replacements(
                            match_row.get('수신파일경로', ''),
                            normal_row.get('수신파일경로', ''),
                            match_row,
                            normal_row
                        )
                    }
                    print(f"  수신파일경로 생성: {match_row.get('수신파일경로', '')} -> {modify_path(normal_row.get('수신파일경로', ''))}")
                
                # 3. 송신스키마파일명 처리
                if (pd.notna(normal_row.get('송신스키마파일생성여부')) and 
                    str(normal_row.get('송신스키마파일생성여부')).strip() == '1'):
                    
                    base_name = os.path.splitext(os.path.basename(normal_row.get('송신스키마파일명', '')))[0]
                    namespace, _ = process_schema_path(normal_row.get('송신스키마파일명', ''))
                    
                    if namespace:
                        yaml_structure[f"{i//2 + 1}번째 행"]["송신스키마파일명"] = {
                            "원본파일": match_row.get('송신스키마파일명', ''),
                            "복사파일": modify_path(normal_row.get('송신스키마파일명', '')),
                            "치환목록": [{
                                "설명": "xs:schema xmlns 치환",
                                "찾기": {"정규식": f'xmlns\\s*=\\s*"[^"]*{base_name}[^"]*"'},
                                "교체": {"값": f'xmlns="{namespace}"'}
                            },
                            {
                                "설명": "xs:schema targetNamespace 치환",
                                "찾기": {"정규식": f'targetNamespace\\s*=\\s*"[^"]*{base_name}[^"]*"'},
                                "교체": {"값": f'targetNamespace="{namespace}"'}
                            }]
                        }
                        print(f"  송신스키마파일명 생성: {match_row.get('송신스키마파일명', '')} -> {modify_path(normal_row.get('송신스키마파일명', ''))}")
                
                # 4. 수신스키마파일명 처리
                if (pd.notna(normal_row.get('수신스키마파일생성여부')) and 
                    str(normal_row.get('수신스키마파일생성여부')).strip() == '1'):
                    
                    base_name = os.path.splitext(os.path.basename(normal_row.get('수신스키마파일명', '')))[0]
                    namespace, _ = process_schema_path(normal_row.get('수신스키마파일명', ''))
                    
                    if namespace:
                        yaml_structure[f"{i//2 + 1}번째 행"]["수신스키마파일명"] = {
                            "원본파일": match_row.get('수신스키마파일명', ''),
                            "복사파일": modify_path(normal_row.get('수신스키마파일명', '')),
                            "치환목록": [{
                                "설명": "xs:schema xmlns 치환",
                                "찾기": {"정규식": f'xmlns\\s*=\\s*"[^"]*{base_name}[^"]*"'},
                                "교체": {"값": f'xmlns="{namespace}"'}
                            },
                            {
                                "설명": "xs:schema targetNamespace 치환",
                                "찾기": {"정규식": f'targetNamespace\\s*=\\s*"[^"]*{base_name}[^"]*"'},
                                "교체": {"값": f'targetNamespace="{namespace}"'}
                            }]
                        }
                        print(f"  수신스키마파일명 생성: {match_row.get('수신스키마파일명', '')} -> {modify_path(normal_row.get('수신스키마파일명', ''))}")
                
                # YAML 구조 출력 (디버그 모드일 때)
                if self.debug_mode and yaml_structure[f"{i//2 + 1}번째 행"]:
                    self.debug_print("\nYAML 구조:")
                    self.debug_print(yaml.dump(yaml_structure, allow_unicode=True, sort_keys=False))
                
                # 전체 YAML 구조에 추가
                full_yaml_structure.update(yaml_structure)
            
            # YAML 파일 생성
            try:
                with open(yaml_path, 'w', encoding='utf-8') as yf:
                    yaml.dump(full_yaml_structure, yf, allow_unicode=True, sort_keys=False)
                print(f"\nYAML 파일이 생성되었습니다: {yaml_path}")
                print(f"총 {len(full_yaml_structure)}개의 작업이 생성되었습니다.")
                return True
            except Exception as e:
                print(f"\nYAML 파일 생성 중 오류 발생: {str(e)}")
                return False
                
        except Exception as e:
            print(f"Excel to YAML 변환 중 오류 발생: {str(e)}")
            return False
    
    def apply_schema_replacements(self, file_path: str, replacements: List[Dict]) -> bool:
        """
        파일에 치환 목록을 적용
        
        Args:
            file_path: 대상 파일 경로
            replacements: 치환 규칙 목록
            
        Returns:
            치환 성공 여부
        """
        try:
            self.debug_print(f"\n=== 파일 치환 시작: {file_path} ===")
            
            # 파일 읽기
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            modified = False
            
            # 각 치환 규칙 적용
            for idx, repl in enumerate(replacements, 1):
                self.debug_print(f"\n--- 치환 규칙 {idx}/{len(replacements)} 적용 시도 ---")
                self.debug_print(f"설명: {repl.get('설명', '설명 없음')}")
                
                pattern = repl['찾기']['정규식']
                replacement = repl['교체']['값']
                
                self.debug_print(f"정규식 패턴: {pattern}")
                self.debug_print(f"교체할 값: {replacement}")
                
                # 패턴 매칭 확인
                matches = list(re.finditer(pattern, content))
                if not matches:
                    self.debug_print("패턴이 파일에서 발견되지 않음")
                    continue
                
                self.debug_print(f"패턴 매칭 수: {len(matches)}")
                
                try:
                    # 정규식 치환 수행
                    new_content = re.sub(pattern, replacement, content)
                    if new_content != content:
                        content = new_content
                        modified = True
                        self.debug_print("치환 성공")
                    else:
                        self.debug_print("치환 후 변경사항 없음")
                except Exception as e:
                    self.debug_print(f"치환 중 오류 발생: {str(e)}")
                    continue
            
            # 변경된 경우에만 파일 저장
            if modified:
                self.debug_print("\n파일 저장 시작")
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                self.debug_print("파일 저장 완료")
                return True
            else:
                self.debug_print("\n변경사항이 없어 파일을 저장하지 않음")
                return False
                
        except Exception as e:
            self.debug_print(f"치환 작업 중 예외 발생: {str(e)}")
            return False
    
    def copy_file_with_check(self, source: str, dest: str) -> bool:
        """
        파일을 복사하되, 대상 파일이 이미 존재하면 경고 출력
        
        Args:
            source: 원본 파일 경로
            dest: 대상 파일 경로
            
        Returns:
            복사 성공 여부
        """
        try:
            # 대상 디렉토리가 없으면 생성
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            
            # 대상 파일이 이미 존재하는지 확인
            if os.path.exists(dest):
                print(f"경고: 파일이 이미 존재합니다 - {dest}")
                return False
            
            # 파일 복사
            shutil.copy2(source, dest)
            print(f"파일 복사 완료: {source} -> {dest}")
            return True
        except Exception as e:
            print(f"파일 복사 중 오류 발생: {str(e)}")
            return False
    
    def execute_replacements(self, yaml_path: str, log_path: Optional[str] = None, result_excel: Optional[str] = None) -> bool:
        """
        YAML에 정의된 복사 및 치환 작업을 실행
        
        Args:
            yaml_path: YAML 파일 경로
            log_path: 로그 파일 경로 (None이면 자동 생성)
            result_excel: 결과 Excel 파일 경로 (None이면 자동 생성)
            
        Returns:
            실행 성공 여부
        """
        try:
            self.debug_print(f"YAML 파일 읽기 시작: {yaml_path}")
            
            # YAML 파일 읽기
            with open(yaml_path, 'r', encoding='utf-8') as yf:
                data = yaml.safe_load(yf)
            
            if not data:
                print("실행할 작업이 없습니다.")
                return False
            
            # 로그 파일 설정
            if log_path is None:
                log_path = f"rft_execution_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            
            # 로그 파일 초기화
            with open(log_path, 'w', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] 작업 시작\n")
            
            summary_data = []
            total_copies = 0
            total_replacements = 0
            
            # 각 행 처리
            for row_key, row_data in data.items():
                self.debug_print(f"\n=== {row_key} 처리 시작 ===")
                
                # 각 파일 타입 처리
                for file_type, file_info in row_data.items():
                    self.debug_print(f"\n--- {file_type} 처리 ---")
                    
                    source = file_info.get('원본파일')
                    dest = file_info.get('복사파일')
                    replacements = file_info.get('치환목록', [])
                    
                    self.debug_print(f"원본파일: {source}")
                    self.debug_print(f"복사파일: {dest}")
                    self.debug_print(f"치환규칙 수: {len(replacements)}")
                    
                    if not source or not dest:
                        self.debug_print("원본 또는 대상 파일 경로가 없음, 건너뜀")
                        continue
                    
                    # 1. 파일 복사
                    self.debug_print(f"\n파일 복사 시도: {source} -> {dest}")
                    if self.copy_file_with_check(source, dest):
                        total_copies += 1
                        self.debug_print("파일 복사 성공")
                        
                        # 2. 치환 목록이 있는 경우 치환 수행
                        if replacements:
                            self.debug_print(f"치환 작업 시작: {dest}")
                            if self.apply_schema_replacements(dest, replacements):
                                total_replacements += 1
                                self.debug_print("치환 작업 성공")
                            else:
                                self.debug_print("치환 작업 실패 또는 변경사항 없음")
                        else:
                            self.debug_print("치환 규칙 없음, 건너뜀")
                    
                    # 작업 결과 기록
                    summary = f"{file_type}: {source} -> {dest}"
                    if replacements:
                        summary += f" (치환: {len(replacements)}개 규칙)"
                    summary_data.append(summary)
                    
                    # 로그 기록
                    with open(log_path, 'a', encoding='utf-8') as lf:
                        lf.write(f"[{datetime.datetime.now()}] {summary}\n")
            
            # 결과 파일 생성
            if result_excel is None:
                result_excel = f"rft_execution_result_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            
            result_df = pd.DataFrame({
                '작업내용': summary_data,
                '실행시간': [datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * len(summary_data)
            })
            result_df.to_csv(result_excel, index=False, encoding='utf-8-sig')
            
            print(f"\n작업이 완료되었습니다.")
            print(f"총 복사 파일 수: {total_copies}")
            print(f"총 치환 파일 수: {total_replacements}")
            print(f"로그 파일: {log_path}")
            print(f"결과 파일: {result_excel}")
            
            return True
            
        except Exception as e:
            print(f"치환 실행 중 오류 발생: {str(e)}")
            return False


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("YAML 처리 도구")
    print("=" * 60)
    
    processor = YAMLProcessor()
    
    while True:
        print("\n메뉴:")
        print("1. Excel/CSV에서 YAML 생성")
        print("2. YAML 기반 파일 복사 및 치환 실행")
        print("0. 종료")
        
        choice = input("\n선택하세요: ").strip()
        
        if choice == "1":
            excel_path = input("Excel/CSV 파일 경로를 입력하세요: ").strip()
            yaml_path = input("생성할 YAML 파일 경로를 입력하세요: ").strip()
            
            if excel_path and yaml_path:
                processor.generate_yaml_from_excel(excel_path, yaml_path)
            else:
                print("파일 경로를 모두 입력해야 합니다.")
                
        elif choice == "2":
            yaml_path = input("YAML 파일 경로를 입력하세요: ").strip()
            
            if yaml_path:
                log_path = input("로그 파일 경로 (Enter: 자동생성): ").strip()
                result_path = input("결과 파일 경로 (Enter: 자동생성): ").strip()
                
                if not log_path:
                    log_path = None
                if not result_path:
                    result_path = None
                
                processor.execute_replacements(yaml_path, log_path, result_path)
            else:
                print("YAML 파일 경로를 입력해야 합니다.")
                
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
            
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")


if __name__ == "__main__":
    main()