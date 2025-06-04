"""
BW Tools YAML Processor
Excel 파일에서 YAML 파일을 생성하고, YAML 규칙에 따라 파일 복사 및 치환을 수행합니다.
기존 string_replacer.py의 역할을 수행합니다.
"""

import os
import yaml
import shutil
import re
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from bwtools_config import (
    COLUMN_NAMES, ADDITIONAL_COLUMNS, REPLACEMENT_RULES, TEST_CONFIG
)

class YAMLProcessor:
    def __init__(self):
        """YAMLProcessor 초기화"""
        self.yaml_data = {}
        self.log_entries = []
        self.copied_files = []
        
    def generate_yaml_from_excel(self, excel_path: str, yaml_path: Optional[str] = None) -> bool:
        """
        Excel 파일에서 YAML 파일을 생성합니다. (모드 1)
        
        Args:
            excel_path: 입력 Excel/CSV 파일 경로
            yaml_path: 출력 YAML 파일 경로 (기본값: 자동 생성)
            
        Returns:
            성공 여부
        """
        try:
            # Excel/CSV 읽기
            df = self._read_input_file(excel_path)
            
            # YAML 구조 생성
            self.yaml_data = self._create_yaml_structure(df)
            
            # YAML 파일 경로 설정
            if not yaml_path:
                base_name = os.path.splitext(os.path.basename(excel_path))[0]
                yaml_path = f"{base_name}_replacement_rules.yaml"
            
            # YAML 파일 저장
            with open(yaml_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.yaml_data, f, allow_unicode=True, default_flow_style=False)
            
            print(f"YAML 파일 생성 완료: {yaml_path}")
            print(f"총 {len(self.yaml_data)}개 행의 치환 규칙 생성")
            return True
            
        except Exception as e:
            print(f"YAML 생성 중 오류 발생: {str(e)}")
            return False
    
    def execute_replacements(self, yaml_path: str, 
                           log_path: Optional[str] = None,
                           result_excel_path: Optional[str] = None) -> bool:
        """
        YAML 파일의 규칙에 따라 파일 복사 및 치환을 수행합니다. (모드 3)
        
        Args:
            yaml_path: 입력 YAML 파일 경로
            log_path: 로그 파일 경로 (기본값: 자동 생성)
            result_excel_path: 결과 Excel 파일 경로 (기본값: 자동 생성)
            
        Returns:
            성공 여부
        """
        try:
            # YAML 파일 읽기
            with open(yaml_path, 'r', encoding='utf-8') as f:
                self.yaml_data = yaml.safe_load(f)
            
            if not self.yaml_data:
                print("실행할 작업이 없습니다.")
                return False
            
            # 로그 파일 경로 설정
            if not log_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                log_path = f"bwtools_replacement_log_{timestamp}.txt"
            
            # 치환 작업 실행
            self._execute_all_replacements()
            
            # 로그 파일 저장
            self._save_log_file(log_path)
            
            # 결과 Excel 생성
            if not result_excel_path:
                result_excel_path = "bwtools_replacement_result.xlsx"
            self._save_result_excel(result_excel_path)
            
            # 삭제 배치 파일 생성
            batch_path = os.path.splitext(result_excel_path)[0] + "_delete.bat"
            self._generate_delete_batch(batch_path)
            
            print(f"\n치환 작업 완료:")
            print(f"- 로그 파일: {log_path}")
            print(f"- 결과 Excel: {result_excel_path}")
            print(f"- 삭제 배치: {batch_path}")
            print(f"- 총 {len(self.copied_files)}개 파일 처리")
            
            return True
            
        except Exception as e:
            print(f"치환 실행 중 오류 발생: {str(e)}")
            return False
    
    def _read_input_file(self, file_path: str) -> pd.DataFrame:
        """입력 파일을 읽습니다."""
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext in ['.xlsx', '.xls']:
            return pd.read_excel(file_path)
        elif ext == '.csv':
            # 인코딩 자동 감지
            for encoding in ['utf-8', 'cp949', 'euc-kr']:
                try:
                    return pd.read_csv(file_path, encoding=encoding)
                except UnicodeDecodeError:
                    continue
            raise ValueError(f"CSV 파일 인코딩을 감지할 수 없습니다: {file_path}")
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {ext}")
    
    def _create_yaml_structure(self, df: pd.DataFrame) -> Dict:
        """DataFrame에서 YAML 구조를 생성합니다."""
        yaml_data = {}
        
        # 2줄씩 처리 (기준행, 매칭행)
        for i in range(0, len(df), 2):
            if i + 1 >= len(df):
                break
            
            base_row = df.iloc[i]
            matched_row = df.iloc[i + 1]
            
            row_key = f"row_{i//2 + 1}"
            row_data = {}
            
            # 파일 타입별 처리
            file_types = [
                ('send_file', ADDITIONAL_COLUMNS['send_file_path']),
                ('recv_file', ADDITIONAL_COLUMNS['recv_file_path']),
                ('send_schema', ADDITIONAL_COLUMNS['send_schema_file']),
                ('recv_schema', ADDITIONAL_COLUMNS['recv_schema_file'])
            ]
            
            for file_type, path_col in file_types:
                if path_col in base_row and path_col in matched_row:
                    source_path = str(matched_row[path_col])
                    dest_path = str(base_row[path_col])
                    
                    if pd.notna(source_path) and pd.notna(dest_path):
                        # 치환 규칙 생성
                        replacements = self._generate_replacement_rules(
                            base_row, matched_row, file_type
                        )
                        
                        row_data[file_type] = {
                            '원본파일': source_path,
                            '복사파일': dest_path,
                            '치환목록': replacements
                        }
            
            if row_data:
                yaml_data[row_key] = row_data
        
        return yaml_data
    
    def _generate_replacement_rules(self, base_row: pd.Series, 
                                  matched_row: pd.Series, 
                                  file_type: str) -> List[Dict]:
        """치환 규칙을 생성합니다."""
        rules = []
        
        # 기본 시스템 치환 규칙
        for old, new in REPLACEMENT_RULES['system'].items():
            rules.append({
                '설명': f'{old} → {new} 치환',
                '조건': {'파일타입': file_type},
                '찾기': {'정규식': old},
                '교체': {'값': new}
            })
        
        # 동적 치환 규칙 (컬럼 값 기반)
        dynamic_mappings = [
            (COLUMN_NAMES['send_system'], '송신시스템'),
            (COLUMN_NAMES['recv_system'], '수신시스템'),
            (COLUMN_NAMES['send_corp'], '송신법인'),
            (COLUMN_NAMES['recv_corp'], '수신법인'),
            (COLUMN_NAMES['ems_name'], 'EMS명'),
            (COLUMN_NAMES['group_id'], 'Group ID'),
            (COLUMN_NAMES['event_id'], 'Event_ID')
        ]
        
        for col_name, desc in dynamic_mappings:
            if col_name in base_row and col_name in matched_row:
                old_val = str(matched_row[col_name])
                new_val = str(base_row[col_name])
                
                if pd.notna(old_val) and pd.notna(new_val) and old_val != new_val:
                    rules.append({
                        '설명': f'{desc} 치환: {old_val} → {new_val}',
                        '조건': {'파일타입': file_type},
                        '찾기': {'정규식': re.escape(old_val)},
                        '교체': {'값': new_val}
                    })
        
        return rules
    
    def _execute_all_replacements(self):
        """모든 치환 작업을 실행합니다."""
        for row_key, row_data in self.yaml_data.items():
            for file_type, file_info in row_data.items():
                source = file_info.get('원본파일')
                dest = file_info.get('복사파일')
                replacements = file_info.get('치환목록', [])
                
                if not source or not dest:
                    continue
                
                # 파일 복사 및 치환
                success = self._copy_and_replace(source, dest, replacements)
                
                # 로그 기록
                self.log_entries.append({
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'row': row_key,
                    'file_type': file_type,
                    'source': source,
                    'dest': dest,
                    'success': success,
                    'replacements': len(replacements)
                })
                
                if success:
                    self.copied_files.append(dest)
    
    def _copy_and_replace(self, source: str, dest: str, replacements: List[Dict]) -> bool:
        """파일을 복사하고 치환을 수행합니다."""
        try:
            # 원본 파일 존재 확인
            if not os.path.exists(source):
                print(f"원본 파일이 존재하지 않음: {source}")
                return False
            
            # 대상 디렉토리 생성
            dest_dir = os.path.dirname(dest)
            if dest_dir and not os.path.exists(dest_dir):
                os.makedirs(dest_dir, exist_ok=True)
            
            # 파일 복사
            shutil.copy2(source, dest)
            
            # 치환 수행
            if replacements:
                self._apply_replacements(dest, replacements)
            
            return True
            
        except Exception as e:
            print(f"파일 처리 중 오류: {source} → {dest}")
            print(f"오류 내용: {str(e)}")
            return False
    
    def _apply_replacements(self, file_path: str, replacements: List[Dict]):
        """파일에 치환 규칙을 적용합니다."""
        # 파일 읽기
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 치환 수행
        modified = False
        for rule in replacements:
            pattern = rule['찾기']['정규식']
            replacement = rule['교체']['값']
            
            new_content = re.sub(pattern, replacement, content)
            if new_content != content:
                content = new_content
                modified = True
        
        # 변경사항이 있으면 저장
        if modified:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
    
    def _save_log_file(self, log_path: str):
        """로그 파일을 저장합니다."""
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"BW Tools 치환 작업 로그\n")
            f.write(f"생성 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'='*80}\n\n")
            
            for entry in self.log_entries:
                f.write(f"[{entry['timestamp']}] {entry['row']} - {entry['file_type']}\n")
                f.write(f"  원본: {entry['source']}\n")
                f.write(f"  대상: {entry['dest']}\n")
                f.write(f"  상태: {'성공' if entry['success'] else '실패'}\n")
                f.write(f"  치환 규칙 수: {entry['replacements']}\n")
                f.write("\n")
            
            f.write(f"\n{'='*80}\n")
            f.write(f"총 {len(self.log_entries)}개 작업 수행\n")
            f.write(f"성공: {sum(1 for e in self.log_entries if e['success'])}개\n")
            f.write(f"실패: {sum(1 for e in self.log_entries if not e['success'])}개\n")
    
    def _save_result_excel(self, excel_path: str):
        """결과를 Excel 파일로 저장합니다."""
        # 로그 데이터를 DataFrame으로 변환
        result_data = []
        for entry in self.log_entries:
            result_data.append({
                '시간': entry['timestamp'],
                '행': entry['row'],
                '파일타입': entry['file_type'],
                '원본파일': entry['source'],
                '복사파일': entry['dest'],
                '상태': '성공' if entry['success'] else '실패',
                '치환규칙수': entry['replacements']
            })
        
        df = pd.DataFrame(result_data)
        
        # Excel로 저장
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='치환결과', index=False)
            
            # 서식 적용
            workbook = writer.book
            worksheet = writer.sheets['치환결과']
            
            # 성공/실패에 따른 색상
            success_format = workbook.add_format({'bg_color': '#90EE90'})
            fail_format = workbook.add_format({'bg_color': '#FFB6C1'})
            
            for idx, status in enumerate(df['상태'].values):
                if status == '성공':
                    worksheet.set_row(idx + 1, None, success_format)
                else:
                    worksheet.set_row(idx + 1, None, fail_format)
    
    def _generate_delete_batch(self, batch_path: str):
        """삭제 배치 파일을 생성합니다."""
        with open(batch_path, 'w', encoding='utf-8') as f:
            f.write('@echo off\n')
            f.write('chcp 65001\n')  # UTF-8 코드 페이지
            f.write('cls\n')
            f.write('echo BW Tools 생성 파일 삭제 배치\n')
            f.write(f'echo 생성 시간: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            f.write(f'echo 총 {len(self.copied_files)}개 파일 삭제 예정\n')
            f.write('echo.\n')
            f.write('echo 주의: Everything 등 파일 인덱싱 도구를 종료한 후 실행하세요.\n')
            f.write('echo.\n')
            f.write('pause\n')
            f.write('echo.\n\n')
            
            # 파일 속성 제거 및 삭제
            for file_path in self.copied_files:
                f.write(f'echo 삭제 중: {file_path}\n')
                f.write(f'attrib -r -h -s "{file_path}" 2>nul\n')
                f.write(f'del /f /q "{file_path}"\n')
                f.write('if errorlevel 1 (\n')
                f.write(f'    echo [실패] {file_path}\n')
                f.write(') else (\n')
                f.write(f'    echo [성공] {file_path}\n')
                f.write(')\n')
                f.write('echo.\n')
            
            f.write('\necho 모든 작업이 완료되었습니다.\n')
            f.write('echo Windows 탐색기를 새로고침하려면 F5를 누르세요.\n')
            f.write('pause\n')


def main():
    """메인 실행 함수"""
    processor = YAMLProcessor()
    
    # 테스트: Excel → YAML 생성
    test_excel = "bwtools_excel_output_test.csv"
    print(f"YAML 생성 테스트 (입력: {test_excel})")
    
    if os.path.exists(test_excel):
        processor.generate_yaml_from_excel(test_excel)
    else:
        print(f"테스트 파일이 없습니다: {test_excel}")
        print("먼저 bwtools_excel_generator.py를 실행하여 테스트 파일을 생성하세요.")


if __name__ == "__main__":
    main()