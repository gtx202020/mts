"""
Excel을 SQLite 데이터베이스로 변환하는 모듈

'작업용 EAI-BW.xlsx' 파일의 'IF현황' 시트를 읽어서 iflist.sqlite로 변환합니다.
"""

import pandas as pd
import sqlite3
import os
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
        self.default_excel_file = "작업용 EAI-BW.xlsx"
        self.default_sheet_name = "IF현황"
    
    def convert_excel_to_sqlite(self, excel_path: Optional[str] = None, sheet_name: Optional[str] = None) -> bool:
        """
        Excel 파일을 SQLite 데이터베이스로 변환
        
        Args:
            excel_path: Excel 파일 경로 (기본값: '작업용 EAI-BW.xlsx')
            sheet_name: 시트명 (기본값: 'IF현황')
            
        Returns:
            변환 성공 여부
        """
        try:
            # 기본값 설정
            excel_path = excel_path or self.default_excel_file
            sheet_name = sheet_name or self.default_sheet_name
            
            print(f"Excel 파일 읽기 시작: {excel_path}, 시트: {sheet_name}")
            
            # Excel 파일이 존재하는지 확인
            if not os.path.exists(excel_path):
                print(f"오류: Excel 파일을 찾을 수 없습니다 - {excel_path}")
                return False
            
            # Excel 파일 읽기
            df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
            print(f"Excel 데이터 로드 완료: {len(df)}행 x {len(df.columns)}열")
            
            # SQLite 데이터베이스 연결
            conn = sqlite3.connect(self.db_filename)
            
            try:
                # DataFrame을 SQLite 테이블로 저장
                df.to_sql(self.table_name, conn, index=False, if_exists='replace')
                
                # 데이터 검증
                cursor = conn.cursor()
                cursor.execute(f"SELECT COUNT(*) FROM {self.table_name}")
                row_count = cursor.fetchone()[0]
                
                cursor.execute(f"PRAGMA table_info({self.table_name})")
                columns = cursor.fetchall()
                
                print(f"SQLite 데이터베이스 생성 완료: {self.db_filename}")
                print(f"테이블명: {self.table_name}")
                print(f"저장된 행 수: {row_count}")
                print(f"컬럼 수: {len(columns)}")
                
                return True
                
            except Exception as e:
                print(f"데이터베이스 생성 중 오류 발생: {str(e)}")
                return False
            finally:
                conn.close()
            
        except Exception as e:
            print(f"Excel to SQLite 변환 중 오류 발생: {str(e)}")
            return False
    
    def create_test_database(self) -> bool:
        """
        기본 Excel 파일을 사용하여 데이터베이스 생성
        
        Returns:
            생성 성공 여부
        """
        return self.convert_excel_to_sqlite()


def convert_default_excel():
    """기본 Excel 파일을 SQLite로 변환하는 함수 (스크립트 실행용)"""
    df = pd.read_excel('작업용 EAI-BW.xlsx', sheet_name='IF현황', dtype=str)
    db_filename = 'iflist.sqlite'
    conn = sqlite3.connect(db_filename)

    try:
        df.to_sql('iflist', conn, index=False, if_exists='replace')
        print(f"데이터베이스 생성 완료: {db_filename}")
        print(f"행 수: {len(df)}")
    except Exception as e:
        print(f"데이터베이스 생성 중 오류 발생: {str(e)}")
    finally:
        conn.close()


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("Excel to SQLite 변환 도구")
    print("=" * 60)
    
    converter = ExcelToSQLiteConverter()
    
    while True:
        print("\n메뉴:")
        print("1. 기본 Excel 파일을 SQLite로 변환 (작업용 EAI-BW.xlsx)")
        print("2. 사용자 지정 Excel 파일을 SQLite로 변환")
        print("0. 종료")
        
        choice = input("\n선택하세요: ").strip()
        
        if choice == "1":
            success = converter.convert_excel_to_sqlite()
            if success:
                print("✓ 변환 완료")
            else:
                print("✗ 변환 실패")
                
        elif choice == "2":
            excel_path = input("Excel 파일 경로를 입력하세요: ").strip()
            sheet_name = input("시트명을 입력하세요 (Enter: IF현황): ").strip()
            if not sheet_name:
                sheet_name = "IF현황"
                
            if excel_path:
                success = converter.convert_excel_to_sqlite(excel_path, sheet_name)
                if success:
                    print("✓ 변환 완료")
                else:
                    print("✗ 변환 실패")
            else:
                print("파일 경로를 입력해야 합니다.")
                
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
            
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")


if __name__ == "__main__":
    # 스크립트로 직접 실행되면 기본 변환 수행
    convert_default_excel()



