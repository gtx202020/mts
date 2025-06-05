"""
Excel을 SQLite 데이터베이스로 변환하는 모듈

이 모듈은 Excel 파일을 읽어서 SQLite 데이터베이스로 변환합니다.
"""

import os
import sqlite3
import pandas as pd
from typing import Optional


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
    
    def convert_excel_to_sqlite(self, excel_path: str) -> bool:
        """
        Excel 파일을 SQLite 데이터베이스로 변환
        
        Args:
            excel_path: Excel 파일 경로
            
        Returns:
            변환 성공 여부
        """
        try:
            print(f"Excel 파일 읽기 시작: {excel_path}")
            
            # Excel 파일이 존재하는지 확인
            if not os.path.exists(excel_path):
                print(f"오류: Excel 파일을 찾을 수 없습니다 - {excel_path}")
                return False
            
            # Excel 파일 읽기
            df = pd.read_excel(excel_path, engine='openpyxl')
            print(f"Excel 데이터 로드 완료: {len(df)}행 x {len(df.columns)}열")
            
            # 기존 데이터베이스 파일 삭제 (있는 경우)
            if os.path.exists(self.db_filename):
                os.remove(self.db_filename)
                print(f"기존 데이터베이스 파일 삭제: {self.db_filename}")
            
            # SQLite 데이터베이스 연결
            conn = sqlite3.connect(self.db_filename)
            
            # DataFrame을 SQLite 테이블로 저장
            df.to_sql(self.table_name, conn, index=False, if_exists='replace')
            
            # 데이터 검증
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {self.table_name}")
            row_count = cursor.fetchone()[0]
            
            cursor.execute(f"PRAGMA table_info({self.table_name})")
            columns = cursor.fetchall()
            
            conn.close()
            
            print(f"SQLite 데이터베이스 생성 완료: {self.db_filename}")
            print(f"테이블명: {self.table_name}")
            print(f"저장된 행 수: {row_count}")
            print(f"컬럼 수: {len(columns)}")
            
            return True
            
        except Exception as e:
            print(f"Excel to SQLite 변환 중 오류 발생: {str(e)}")
            return False
    
    def create_test_database(self) -> bool:
        """
        테스트용 SQLite 데이터베이스 생성
        
        Returns:
            생성 성공 여부
        """
        try:
            print("테스트 데이터베이스 생성 시작")
            
            # 테스트 데이터 생성
            test_data = {
                '송신시스템': ['LY_SYS1', 'LZ_SYS2', 'LY_SYS3'],
                '수신시스템': ['LH_REC1', 'VO_REC2', 'LH_REC3'],
                'I/F명': ['TEST_IF_001', 'TEST_IF_002', 'TEST_IF_003'],
                '송신\n법인': ['KR', 'NJ', 'VH'],
                '수신\n법인': ['KR', 'NJ', 'VH'],
                '송신패키지': ['PKG_LY_001', 'PKG_LZ_002', 'PKG_LY_003'],
                '수신패키지': ['PKG_LH_001', 'PKG_VO_002', 'PKG_LH_003'],
                '송신\n업무명': ['PNL_LY', 'MOD_LZ', 'PNL_LY'],
                '수신\n업무명': ['MES_LH', 'MES_VO', 'MES_LH'],
                'EMS명': ['MES01', 'MES02', 'MES01'],
                'Group ID': ['GRP01', 'GRP02', 'GRP03'],
                'Event_ID': ['EVT001', 'EVT002', 'EVT003'],
                '개발구분': ['신규', '수정', '신규'],
                'Source Table': ['LY.TB_TEST01', 'LZ.TB_TEST02', 'LY.TB_TEST03'],
                'Destination Table': ['LH.TB_DEST01', 'VO.TB_DEST02', 'LH.TB_DEST03'],
                'Routing': ['RT_LY_01', 'RT_LZ_02', 'RT_LY_03'],
                '스케쥴': ['매일 09:00', '매일 18:00', '매일 12:00'],
                '주기구분': ['일배치', '일배치', '일배치'],
                '주기': ['Daily', 'Daily', 'Daily'],
                '송신\nDB Name': ['LYDB', 'LZDB', 'LYDB'],
                '송신 \nSchema': ['LYSCH', 'LZSCH', 'LYSCH']
            }
            
            df = pd.DataFrame(test_data)
            
            # 기존 데이터베이스 파일 삭제 (있는 경우)
            if os.path.exists(self.db_filename):
                os.remove(self.db_filename)
                print(f"기존 데이터베이스 파일 삭제: {self.db_filename}")
            
            # SQLite 데이터베이스 연결
            conn = sqlite3.connect(self.db_filename)
            
            # DataFrame을 SQLite 테이블로 저장
            df.to_sql(self.table_name, conn, index=False, if_exists='replace')
            
            # 데이터 검증
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {self.table_name}")
            row_count = cursor.fetchone()[0]
            
            conn.close()
            
            print(f"테스트 SQLite 데이터베이스 생성 완료: {self.db_filename}")
            print(f"테이블명: {self.table_name}")
            print(f"저장된 행 수: {row_count}")
            
            return True
            
        except Exception as e:
            print(f"테스트 데이터베이스 생성 중 오류 발생: {str(e)}")
            return False


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("Excel to SQLite 변환 도구")
    print("=" * 60)
    
    converter = ExcelToSQLiteConverter()
    
    while True:
        print("\n메뉴:")
        print("1. Excel 파일을 SQLite로 변환")
        print("2. 테스트 데이터베이스 생성")
        print("0. 종료")
        
        choice = input("\n선택하세요: ").strip()
        
        if choice == "1":
            excel_path = input("Excel 파일 경로를 입력하세요: ").strip()
            if excel_path:
                converter.convert_excel_to_sqlite(excel_path)
            else:
                print("파일 경로를 입력해야 합니다.")
                
        elif choice == "2":
            converter.create_test_database()
            
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
            
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")


if __name__ == "__main__":
    main()