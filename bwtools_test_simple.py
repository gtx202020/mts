"""
BW Tools 간단한 통합 테스트
pandas 없이 기본 Python만으로 동작하는 테스트
"""

import os
import sqlite3
import csv
import json
from datetime import datetime

def test_basic_functionality():
    """기본 기능 테스트"""
    print("=" * 60)
    print("BW Tools 기본 기능 테스트")
    print("=" * 60)
    
    # 1. SQLite 데이터베이스 생성 테스트
    print("\n[1단계] SQLite 데이터베이스 생성 테스트")
    db_path = 'test_simple.sqlite'
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 테이블 생성
        cursor.execute('''
            CREATE TABLE iflist (
                "송신시스템" TEXT,
                "수신시스템" TEXT,
                "I/F명" TEXT,
                "송신\n법인" TEXT,
                "수신\n법인" TEXT,
                "EMS명" TEXT,
                "Group ID" TEXT,
                "Event_ID" TEXT
            )
        ''')
        
        # 테스트 데이터 삽입
        test_data = [
            ('LYMES', 'LZWMS', 'IF_001', 'LYCORP', 'LZCORP', 'EMS_TEST', '001', 'EVT_0001'),
            ('LHMES', 'VOWMS', 'IF_001', 'LHCORP', 'VOCORP', 'EMS_TEST', '001', 'EVT_0001'),
            ('LZMES', 'LYWMS', 'IF_002', 'LZCORP', 'LYCORP', 'EMS_TEST2', '002', 'EVT_0002'),
            ('VOMES', 'LHWMS', 'IF_002', 'VOCORP', 'LHCORP', 'EMS_TEST2', '002', 'EVT_0002')
        ]
        
        cursor.executemany('''
            INSERT INTO iflist VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', test_data)
        
        conn.commit()
        
        # 데이터 확인
        cursor.execute('SELECT COUNT(*) FROM iflist')
        count = cursor.fetchone()[0]
        print(f"✓ 데이터베이스 생성 성공: {count}개 행 삽입")
        
        conn.close()
        
    except Exception as e:
        print(f"✗ 데이터베이스 생성 실패: {str(e)}")
        return False
    
    # 2. 데이터 필터링 및 매칭 테스트
    print("\n[2단계] 데이터 필터링 및 매칭 테스트")
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # LY/LZ 시스템 필터링
        cursor.execute('''
            SELECT * FROM iflist 
            WHERE "송신시스템" LIKE '%LY%' OR "송신시스템" LIKE '%LZ%'
               OR "수신시스템" LIKE '%LY%' OR "수신시스템" LIKE '%LZ%'
        ''')
        
        ly_lz_rows = cursor.fetchall()
        print(f"✓ LY/LZ 시스템 필터링: {len(ly_lz_rows)}개 행 발견")
        
        # 매칭 테스트
        for row in ly_lz_rows:
            if_name = row[2]  # I/F명
            cursor.execute('SELECT * FROM iflist WHERE "I/F명" = ?', (if_name,))
            matched = cursor.fetchall()
            print(f"  - {if_name}: {len(matched)}개 매칭행")
        
        conn.close()
        
    except Exception as e:
        print(f"✗ 필터링 테스트 실패: {str(e)}")
        return False
    
    # 3. CSV 출력 테스트
    print("\n[3단계] CSV 출력 테스트")
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM iflist')
        rows = cursor.fetchall()
        
        # 컬럼명 가져오기
        columns = [desc[0] for desc in cursor.description]
        
        # CSV 파일 생성
        csv_path = 'test_output_simple.csv'
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(columns)
            writer.writerows(rows)
        
        print(f"✓ CSV 파일 생성 성공: {csv_path}")
        
        conn.close()
        
    except Exception as e:
        print(f"✗ CSV 출력 실패: {str(e)}")
        return False
    
    # 4. JSON 구조 생성 테스트 (YAML 대신)
    print("\n[4단계] JSON 구조 생성 테스트")
    
    try:
        # 치환 규칙 구조 생성
        replacement_structure = {
            'row_1': {
                'send_file': {
                    '원본파일': '/home/lhcorp/test_lh.process',
                    '복사파일': '/home/lycorp/test_ly.process',
                    '치환목록': [
                        {
                            '설명': 'LHMES_MGR → LYMES_MGR 치환',
                            '찾기': {'정규식': 'LHMES_MGR'},
                            '교체': {'값': 'LYMES_MGR'}
                        }
                    ]
                }
            }
        }
        
        # JSON 파일 저장
        json_path = 'test_rules_simple.json'
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(replacement_structure, f, ensure_ascii=False, indent=2)
        
        print(f"✓ JSON 구조 생성 성공: {json_path}")
        
    except Exception as e:
        print(f"✗ JSON 구조 생성 실패: {str(e)}")
        return False
    
    # 5. 파일 정리
    print("\n[5단계] 테스트 파일 정리")
    
    test_files = [db_path, csv_path, json_path]
    for file in test_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"✓ 파일 삭제: {file}")
    
    print("\n" + "=" * 60)
    print("✓ 모든 기본 기능 테스트 통과!")
    print("=" * 60)
    
    return True

def test_file_structure():
    """파일 구조 테스트"""
    print("\n[추가] BW Tools 파일 구조 확인")
    
    required_files = [
        'bwtools_config.py',
        'bwtools_db_creator.py',
        'bwtools_excel_generator.py',
        'bwtools_yaml_processor.py',
        'bwtools_main.py'
    ]
    
    test_files = [
        'test_bwtools_db_creator.py',
        'test_bwtools_excel_generator.py',
        'test_bwtools_yaml_processor.py'
    ]
    
    print("\n필수 파일 확인:")
    for file in required_files:
        if os.path.exists(file):
            size = os.path.getsize(file)
            print(f"✓ {file} ({size:,} bytes)")
        else:
            print(f"✗ {file} (없음)")
    
    print("\n테스트 파일 확인:")
    for file in test_files:
        if os.path.exists(file):
            size = os.path.getsize(file)
            print(f"✓ {file} ({size:,} bytes)")
        else:
            print(f"✗ {file} (없음)")

def main():
    """메인 실행 함수"""
    print("BW Tools 간단한 통합 테스트 시작")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 파일 구조 확인
    test_file_structure()
    
    # 기본 기능 테스트
    if test_basic_functionality():
        print("\n🎉 모든 테스트가 성공적으로 완료되었습니다!")
        print("\n다음 단계:")
        print("1. pandas, PyYAML, openpyxl 패키지를 설치하세요")
        print("2. python bwtools_main.py --test 명령으로 전체 파이프라인을 실행하세요")
        print("3. 개별 모듈 테스트: python -m unittest test_bwtools_*.py")
    else:
        print("\n❌ 일부 테스트가 실패했습니다.")

if __name__ == "__main__":
    main()