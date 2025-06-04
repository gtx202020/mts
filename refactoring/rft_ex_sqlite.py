"""
Excel을 SQLite 데이터베이스로 변환하는 모듈

이 모듈은 Excel 파일을 읽어서 SQLite 데이터베이스로 변환합니다.
CSV 파일도 지원하며, 전체 데이터를 SQLite DB로 변환합니다.
"""

import os
import sqlite3
import pandas as pd
from typing import Optional, Union, Dict


class ExcelToSQLiteConverter:
    """Excel 파일을 SQLite 데이터베이스로 변환하는 클래스"""
    
    def __init__(self, db_filename: str = "iflist.sqlite"):
        """
        ExcelToSQLiteConverter 초기화
        
        Args:
            db_filename: 생성할 SQLite 데이터베이스 파일명
        """
        self.db_filename = db_filename
        self.table_name = "iflist"
    
    def convert_to_sqlite(self, file_path: str, table_name: Optional[str] = None, 
                         if_exists: str = 'replace') -> bool:
        """
        Excel 또는 CSV 파일을 SQLite 데이터베이스로 변환
        
        Args:
            file_path: Excel 또는 CSV 파일 경로
            table_name: 테이블명 (기본값: self.table_name)
            if_exists: 테이블이 존재할 경우 처리 방법 ('replace', 'append', 'fail')
            
        Returns:
            변환 성공 여부
        """
        try:
            print(f"파일 읽기 시작: {file_path}")
            
            # 파일이 존재하는지 확인
            if not os.path.exists(file_path):
                print(f"오류: 파일을 찾을 수 없습니다 - {file_path}")
                return False
            
            # 파일 읽기
            df = self._load_file(file_path)
            print(f"데이터 로드 완료: {len(df)}행 x {len(df.columns)}열")
            
            # 모든 행 출력 (디버깅용)
            print(f"전체 {len(df)}개 행을 처리합니다.")
            
            # 테이블명 설정
            table_name = table_name or self.table_name
            
            # SQLite 데이터베이스 연결
            conn = sqlite3.connect(self.db_filename)
            
            # DataFrame의 모든 행을 SQLite 테이블로 저장
            df.to_sql(table_name, conn, index=False, if_exists=if_exists)
            
            # 데이터 검증
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            row_count = cursor.fetchone()[0]
            
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns = cursor.fetchall()
            
            conn.close()
            
            print(f"SQLite 데이터베이스 생성 완료: {self.db_filename}")
            print(f"테이블명: {table_name}")
            print(f"저장된 행 수: {row_count} (전체 데이터)")
            print(f"컬럼 수: {len(columns)}")
            print(f"컬럼 목록: {[col[1] for col in columns]}")
            
            return True
            
        except Exception as e:
            print(f"파일 to SQLite 변환 중 오류 발생: {str(e)}")
            return False
    
    def convert_excel_to_sqlite(self, excel_path: str) -> bool:
        """
        Excel 파일을 SQLite 데이터베이스로 변환 (하위 호환성 유지)
        
        Args:
            excel_path: Excel 파일 경로
            
        Returns:
            변환 성공 여부
        """
        return self.convert_to_sqlite(excel_path)
    
    def _load_file(self, file_path: str) -> pd.DataFrame:
        """
        파일로부터 데이터를 로드합니다.
        
        Args:
            file_path: Excel 또는 CSV 파일 경로
            
        Returns:
            로드된 DataFrame
        """
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext in ['.xlsx', '.xls']:
            # Excel 파일 읽기
            return pd.read_excel(file_path, engine='openpyxl')
        elif file_ext == '.csv':
            # CSV 파일 읽기 (인코딩 자동 감지)
            encodings = ['utf-8', 'cp949', 'euc-kr', 'latin1']
            for encoding in encodings:
                try:
                    return pd.read_csv(file_path, encoding=encoding)
                except UnicodeDecodeError:
                    continue
            raise ValueError(f"CSV 파일 인코딩을 감지할 수 없습니다: {file_path}")
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {file_ext}")
    
    def create_test_database(self, num_rows: int = 100) -> bool:
        """
        테스트용 SQLite 데이터베이스 생성 (더 많은 데이터 생성)
        
        Args:
            num_rows: 생성할 행 수 (기본값: 100)
            
        Returns:
            생성 성공 여부
        """
        try:
            print(f"테스트 데이터베이스 생성 시작 ({num_rows}개 행)")
            
            # 테스트 데이터 생성
            test_data = self._generate_test_data(num_rows)
            df = pd.DataFrame(test_data)
            
            # 기존 데이터베이스 파일 삭제 (있는 경우)
            if os.path.exists(self.db_filename):
                os.remove(self.db_filename)
                print(f"기존 데이터베이스 파일 삭제: {self.db_filename}")
            
            # SQLite 데이터베이스 연결
            conn = sqlite3.connect(self.db_filename)
            
            # DataFrame을 SQLite 테이블로 저장 (전체 데이터)
            df.to_sql(self.table_name, conn, index=False, if_exists='replace')
            
            # 데이터 검증
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {self.table_name}")
            row_count = cursor.fetchone()[0]
            
            cursor.execute(f"PRAGMA table_info({self.table_name})")
            columns = cursor.fetchall()
            
            conn.close()
            
            print(f"테스트 SQLite 데이터베이스 생성 완료: {self.db_filename}")
            print(f"테이블명: {self.table_name}")
            print(f"저장된 행 수: {row_count} (전체 데이터)")
            print(f"컬럼 수: {len(columns)}")
            
            return True
            
        except Exception as e:
            print(f"테스트 데이터베이스 생성 중 오류 발생: {str(e)}")
            return False
    
    def _generate_test_data(self, num_rows: int) -> Dict:
        """
        테스트용 데이터를 생성합니다.
        
        Args:
            num_rows: 생성할 행 수
            
        Returns:
            테스트 데이터 딕셔너리
        """
        import random
        
        systems = ['LY', 'LZ', 'LH', 'VO']
        corps = ['KR', 'NJ', 'VH', 'JP', 'CN']
        dev_types = ['신규', '수정', '삭제', '변경']
        cycles = ['일배치', '실시간', '시간배치', '월배치']
        
        data = {
            '송신시스템': [],
            '수신시스템': [],
            'I/F명': [],
            '송신\n법인': [],
            '수신\n법인': [],
            '송신패키지': [],
            '수신패키지': [],
            '송신\n업무명': [],
            '수신\n업무명': [],
            'EMS명': [],
            'Group ID': [],
            'Event_ID': [],
            '개발구분': [],
            'Source Table': [],
            'Destination Table': [],
            'Routing': [],
            '스케쥴': [],
            '주기구분': [],
            '주기': [],
            '송신\nDB Name': [],
            '송신 \nSchema': []
        }
        
        for i in range(num_rows):
            send_sys = random.choice(systems)
            recv_sys = random.choice([s for s in systems if s != send_sys])
            send_corp = random.choice(corps)
            recv_corp = random.choice(corps)
            
            data['송신시스템'].append(f'{send_sys}_SYS{i+1:03d}')
            data['수신시스템'].append(f'{recv_sys}_REC{i+1:03d}')
            data['I/F명'].append(f'IF_{send_sys}_{recv_sys}_{i+1:04d}')
            data['송신\n법인'].append(send_corp)
            data['수신\n법인'].append(recv_corp)
            data['송신패키지'].append(f'PKG_{send_sys}_{i+1:03d}')
            data['수신패키지'].append(f'PKG_{recv_sys}_{i+1:03d}')
            data['송신\n업무명'].append(f'TASK_{send_sys}')
            data['수신\n업무명'].append(f'TASK_{recv_sys}')
            data['EMS명'].append(f'EMS{(i % 10) + 1:02d}')
            data['Group ID'].append(f'GRP{i+1:04d}')
            data['Event_ID'].append(f'EVT{i+1:05d}')
            data['개발구분'].append(random.choice(dev_types))
            data['Source Table'].append(f'{send_sys}.TB_SRC_{i+1:04d}')
            data['Destination Table'].append(f'{recv_sys}.TB_DEST_{i+1:04d}')
            data['Routing'].append(f'RT_{send_sys}_{recv_sys}_{i+1:02d}')
            data['스케쥴'].append(f'매일 {random.randint(0,23):02d}:{random.randint(0,59):02d}')
            data['주기구분'].append(random.choice(cycles))
            data['주기'].append('Daily' if '일' in data['주기구분'][-1] else 'Hourly')
            data['송신\nDB Name'].append(f'{send_sys}DB')
            data['송신 \nSchema'].append(f'{send_sys}SCHEMA')
        
        return data
    
    def verify_database(self) -> Dict:
        """
        생성된 데이터베이스를 검증합니다.
        
        Returns:
            검증 결과 딕셔너리
        """
        try:
            with sqlite3.connect(self.db_filename) as conn:
                # 테이블 존재 확인
                cursor = conn.cursor()
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{self.table_name}'")
                if not cursor.fetchone():
                    return {'success': False, 'error': f"테이블 '{self.table_name}'이 존재하지 않습니다."}
                
                # 컬럼 정보 확인
                cursor.execute(f"PRAGMA table_info({self.table_name})")
                columns = [col[1] for col in cursor.fetchall()]
                
                # 행 수 확인
                cursor.execute(f"SELECT COUNT(*) FROM {self.table_name}")
                row_count = cursor.fetchone()[0]
                
                # 샘플 데이터 확인 (처음 5개 행)
                cursor.execute(f"SELECT * FROM {self.table_name} LIMIT 5")
                sample_data = cursor.fetchall()
                
                return {
                    'success': True,
                    'table_name': self.table_name,
                    'columns': columns,
                    'column_count': len(columns),
                    'row_count': row_count,
                    'sample_data': sample_data
                }
                
        except Exception as e:
            return {'success': False, 'error': str(e)}


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("Excel/CSV to SQLite 변환 도구 (전체 데이터 처리)")
    print("=" * 60)
    
    converter = ExcelToSQLiteConverter()
    
    while True:
        print("\n메뉴:")
        print("1. Excel/CSV 파일을 SQLite로 변환 (전체 행)")
        print("2. 테스트 데이터베이스 생성")
        print("3. 데이터베이스 검증")
        print("0. 종료")
        
        choice = input("\n선택하세요: ").strip()
        
        if choice == "1":
            file_path = input("Excel 또는 CSV 파일 경로를 입력하세요: ").strip()
            if file_path:
                success = converter.convert_to_sqlite(file_path)
                if success:
                    # 변환 후 검증
                    result = converter.verify_database()
                    if result['success']:
                        print(f"\n검증 결과: 전체 {result['row_count']}개 행이 성공적으로 저장되었습니다.")
            else:
                print("파일 경로를 입력해야 합니다.")
                
        elif choice == "2":
            try:
                num_rows = input("생성할 행 수를 입력하세요 (기본값: 100): ").strip()
                num_rows = int(num_rows) if num_rows else 100
                converter.create_test_database(num_rows)
            except ValueError:
                print("올바른 숫자를 입력하세요.")
            
        elif choice == "3":
            result = converter.verify_database()
            if result['success']:
                print(f"\n데이터베이스 검증 결과:")
                print(f"- 테이블: {result['table_name']}")
                print(f"- 컬럼 수: {result['column_count']}")
                print(f"- 전체 행 수: {result['row_count']}")
                print(f"- 컬럼: {', '.join(result['columns'][:5])}...")
            else:
                print(f"검증 실패: {result['error']}")
            
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
            
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")


if __name__ == "__main__":
    main()