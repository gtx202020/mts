"""
BW Tools DB Creator
Excel 파일 또는 CSV 파일을 SQLite 데이터베이스로 변환합니다.
기존 ex_sqlite.py의 역할을 수행합니다.
"""

import sqlite3
import pandas as pd
import os
from typing import Optional, Union
from bwtools_config import DB_FILENAME, TABLE_NAME, COLUMN_NAMES, TEST_CONFIG

class DBCreator:
    def __init__(self, db_path: Optional[str] = None):
        """
        DBCreator 초기화
        
        Args:
            db_path: SQLite 데이터베이스 경로 (기본값: config의 DB_FILENAME)
        """
        self.db_path = db_path or DB_FILENAME
        self.table_name = TABLE_NAME
        
    def create_database(self, data_source: Union[str, pd.DataFrame], 
                       table_name: Optional[str] = None,
                       if_exists: str = 'replace') -> bool:
        """
        데이터소스로부터 SQLite 데이터베이스를 생성합니다.
        
        Args:
            data_source: Excel/CSV 파일 경로 또는 DataFrame
            table_name: 테이블명 (기본값: config의 TABLE_NAME)
            if_exists: 테이블이 존재할 경우 처리 방법 ('replace', 'append', 'fail')
            
        Returns:
            성공 여부
        """
        try:
            # 데이터 로드
            if isinstance(data_source, str):
                df = self._load_data(data_source)
            elif isinstance(data_source, pd.DataFrame):
                df = data_source
            else:
                raise ValueError("data_source는 파일 경로 또는 DataFrame이어야 합니다.")
            
            # 테이블명 설정
            table_name = table_name or self.table_name
            
            # SQLite에 저장
            with sqlite3.connect(self.db_path) as conn:
                df.to_sql(table_name, conn, if_exists=if_exists, index=False)
                
                # 저장된 행 수 확인
                cursor = conn.cursor()
                cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                row_count = cursor.fetchone()[0]
                print(f"데이터베이스 생성 완료: {self.db_path}")
                print(f"테이블 '{table_name}'에 {row_count}개 행 저장됨")
                
            return True
            
        except Exception as e:
            print(f"데이터베이스 생성 중 오류 발생: {str(e)}")
            return False
    
    def _load_data(self, file_path: str) -> pd.DataFrame:
        """
        파일로부터 데이터를 로드합니다.
        
        Args:
            file_path: Excel 또는 CSV 파일 경로
            
        Returns:
            로드된 DataFrame
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext in ['.xlsx', '.xls']:
            return pd.read_excel(file_path)
        elif file_ext == '.csv':
            # CSV 인코딩 자동 감지
            encodings = ['utf-8', 'cp949', 'euc-kr', 'latin1']
            for encoding in encodings:
                try:
                    return pd.read_csv(file_path, encoding=encoding)
                except UnicodeDecodeError:
                    continue
            raise ValueError(f"CSV 파일 인코딩을 감지할 수 없습니다: {file_path}")
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {file_ext}")
    
    def create_test_database(self) -> bool:
        """
        테스트용 데이터베이스를 생성합니다.
        
        Returns:
            성공 여부
        """
        # 테스트 데이터 생성
        test_data = self._generate_test_data()
        
        # 데이터베이스 생성
        return self.create_database(test_data)
    
    def _generate_test_data(self) -> pd.DataFrame:
        """
        테스트용 데이터를 생성합니다.
        
        Returns:
            테스트 DataFrame
        """
        import random
        
        config = TEST_CONFIG['sample_rows']
        data = []
        
        # LY/LZ 시스템 데이터 생성
        for i in range(config['ly_lz_systems']):
            system = random.choice(['LY', 'LZ'])
            opposite = 'LZ' if system == 'LY' else 'LY'
            
            row = {
                COLUMN_NAMES['send_system']: f'{system}MES',
                COLUMN_NAMES['recv_system']: f'{opposite}WMS',
                COLUMN_NAMES['if_name']: f'IF_{system}_{i:03d}',
                COLUMN_NAMES['send_corp']: f'{system}CORP',
                COLUMN_NAMES['recv_corp']: f'{opposite}CORP',
                COLUMN_NAMES['send_pkg']: f'PKG_{system}_SEND',
                COLUMN_NAMES['recv_pkg']: f'PKG_{opposite}_RECV',
                COLUMN_NAMES['send_task']: f'TASK_{system}_01',
                COLUMN_NAMES['recv_task']: f'TASK_{opposite}_01',
                COLUMN_NAMES['ems_name']: f'EMS_{system}_{opposite}',
                COLUMN_NAMES['group_id']: f'{i+1:03d}',
                COLUMN_NAMES['event_id']: f'EVT_{i+1:04d}',
                COLUMN_NAMES['send_db_name']: f'{system}DB',
                COLUMN_NAMES['send_schema']: f'{system}SCHEMA',
                COLUMN_NAMES['source_table']: f'TB_{system}_SOURCE_{i+1}',
                COLUMN_NAMES['dest_table']: f'TB_{opposite}_DEST_{i+1}',
                COLUMN_NAMES['dev_type']: '신규',  # iflist03a.py에서 확인한 값
                COLUMN_NAMES['routing']: 'DIRECT',
                COLUMN_NAMES['cycle_type']: '실시간',
                COLUMN_NAMES['cycle']: '1분',
                COLUMN_NAMES['schedule']: '매일 00:00-23:59'
            }
            data.append(row)
        
        # LH/VO 시스템 데이터 생성 (매칭되는 데이터)
        for i in range(config['lh_vo_systems']):
            system = random.choice(['LH', 'VO'])
            opposite = 'VO' if system == 'LH' else 'LH'
            
            row = {
                COLUMN_NAMES['send_system']: f'{system}MES',
                COLUMN_NAMES['recv_system']: f'{opposite}WMS',
                COLUMN_NAMES['if_name']: f'IF_{system[1]}{["Y","Z"][system=="VO"]}_{i:03d}',  # LY/LZ와 매칭되도록
                COLUMN_NAMES['send_corp']: f'{system}CORP',
                COLUMN_NAMES['recv_corp']: f'{opposite}CORP',
                COLUMN_NAMES['send_pkg']: f'PKG_{system}_SEND',
                COLUMN_NAMES['recv_pkg']: f'PKG_{opposite}_RECV',
                COLUMN_NAMES['send_task']: f'TASK_{system}_01',
                COLUMN_NAMES['recv_task']: f'TASK_{opposite}_01',
                COLUMN_NAMES['ems_name']: f'EMS_{system}_{opposite}',
                COLUMN_NAMES['group_id']: f'{i+1:03d}',
                COLUMN_NAMES['event_id']: f'EVT_{i+1:04d}',
                COLUMN_NAMES['send_db_name']: f'{system}DB',
                COLUMN_NAMES['send_schema']: f'{system}SCHEMA',
                COLUMN_NAMES['source_table']: f'TB_{system}_SOURCE_{i+1}',
                COLUMN_NAMES['dest_table']: f'TB_{opposite}_DEST_{i+1}',
                COLUMN_NAMES['dev_type']: '신규',  # iflist03a.py에서 확인한 값
                COLUMN_NAMES['routing']: 'DIRECT',
                COLUMN_NAMES['cycle_type']: '실시간',
                COLUMN_NAMES['cycle']: '1분',
                COLUMN_NAMES['schedule']: '매일 00:00-23:59'
            }
            data.append(row)
        
        # 매칭되지 않는 데이터 생성
        for i in range(config['unmatched']):
            row = {
                COLUMN_NAMES['send_system']: f'XMES',
                COLUMN_NAMES['recv_system']: f'YWMS',
                COLUMN_NAMES['if_name']: f'IF_UNMATCHED_{i:03d}',
                COLUMN_NAMES['send_corp']: 'XCORP',
                COLUMN_NAMES['recv_corp']: 'YCORP',
                COLUMN_NAMES['send_pkg']: 'PKG_X_SEND',
                COLUMN_NAMES['recv_pkg']: 'PKG_Y_RECV',
                COLUMN_NAMES['send_task']: 'TASK_X_01',
                COLUMN_NAMES['recv_task']: 'TASK_Y_01',
                COLUMN_NAMES['ems_name']: 'EMS_X_Y',
                COLUMN_NAMES['group_id']: f'{100+i:03d}',
                COLUMN_NAMES['event_id']: f'EVT_{100+i:04d}',
                COLUMN_NAMES['send_db_name']: 'XDB',
                COLUMN_NAMES['send_schema']: 'XSCHEMA',
                COLUMN_NAMES['source_table']: f'TB_X_SOURCE_{i+1}',
                COLUMN_NAMES['dest_table']: f'TB_Y_DEST_{i+1}',
                COLUMN_NAMES['dev_type']: '신규',
                COLUMN_NAMES['routing']: 'DIRECT',
                COLUMN_NAMES['cycle_type']: '배치',
                COLUMN_NAMES['cycle']: '1시간',
                COLUMN_NAMES['schedule']: '매시 정각'
            }
            data.append(row)
        
        return pd.DataFrame(data)
    
    def verify_database(self) -> dict:
        """
        생성된 데이터베이스를 검증합니다.
        
        Returns:
            검증 결과 딕셔너리
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
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
                
                return {
                    'success': True,
                    'table_name': self.table_name,
                    'columns': columns,
                    'column_count': len(columns),
                    'row_count': row_count
                }
                
        except Exception as e:
            return {'success': False, 'error': str(e)}


def main():
    """메인 실행 함수"""
    creator = DBCreator()
    
    # 테스트 데이터베이스 생성
    print("테스트 데이터베이스 생성 중...")
    if creator.create_test_database():
        # 검증
        result = creator.verify_database()
        if result['success']:
            print(f"\n데이터베이스 검증 완료:")
            print(f"- 테이블: {result['table_name']}")
            print(f"- 컬럼 수: {result['column_count']}")
            print(f"- 행 수: {result['row_count']}")
            print(f"- 컬럼: {', '.join(result['columns'][:5])}...")
        else:
            print(f"데이터베이스 검증 실패: {result['error']}")
    else:
        print("데이터베이스 생성 실패")


if __name__ == "__main__":
    main()