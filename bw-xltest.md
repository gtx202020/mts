# BW-XLTEST 프로젝트 문서

## 1. 프로젝트 개요
BW-XLTEST는 데이터 이관이나 인터페이스 개발 시 송신/수신 테이블 간의 컬럼 호환성을 자동으로 검증하는 도구입니다. 기존 xltest.py에서 핵심 기능만 추출하여 더 견고하고 유지보수가 용이한 구조로 재설계했습니다.

### 주요 특징
- Oracle 데이터베이스의 실제 테이블 메타데이터를 조회하여 정확한 검증
- DATE-VARCHAR 타입 변환 등 특수한 경우에 대한 경고 제공
- 시각적으로 구분된 검증 결과 (정상=녹색, 경고=노란색, 오류=빨간색)
- 각 인터페이스별 독립적인 처리로 안정성 확보

## 2. 주요 기능
1. **Excel 파일에서 인터페이스 정보 읽기**
   - DB 연결 정보, 테이블 정보, 컬럼 매핑 정보 파싱

2. **Oracle DB 연결 및 테이블 메타데이터 조회**
   - all_tab_columns 시스템 뷰를 통한 실제 컬럼 정보 수집

3. **송신/수신 테이블 간 컬럼 호환성 검증**
   - 컬럼 존재 여부, 타입 호환성, 크기 비교, NULL 허용 여부 검사
   - DATE-VARCHAR 변환 특별 감지

4. **검증 결과를 Excel 파일로 출력**
   - 인터페이스별 시트 생성
   - 상세한 비교 결과와 시각적 표시

## 3. 제외 기능 (기존 xltest.py 대비)
- SQL 문 생성 기능
- XML 필드 정보 생성 기능
- 1024 바이트 초과 검사 기능

## 4. 상세 요구사항

### 4.1 Excel 입력 파일 구조
- **1행**: 인터페이스명
- **2행**: 인터페이스 ID
- **3행**: DB 연결 정보 (딕셔너리 형태)
  ```python
  {'sid': 'DB_SID', 'username': 'USER', 'password': 'PASS'}
  ```
- **4행**: 테이블 정보 (딕셔너리 형태)
  ```python
  {'owner': 'SCHEMA', 'table_name': 'TABLE'}
  ```
- **5행~**: 컬럼 매핑 정보
  - 송신 컬럼 (첫 번째 열)
  - 수신 컬럼 (두 번째 열)

### 4.2 컬럼 검증 규칙

#### 4.2.1 존재 여부 검사
- 송신 컬럼이 송신 테이블에 존재하는지 확인
- 수신 컬럼이 수신 테이블에 존재하는지 확인

#### 4.2.2 데이터 타입 비교
- 기본: 송신/수신 타입이 동일해야 함
- 예외: VARCHAR, VARCHAR2, CHAR는 서로 호환
- **추가 검사**: DATE ↔ VARCHAR 변환 감지
  - 송신 DATE → 수신 VARCHAR/VARCHAR2/CHAR: 경고
  - 송신 VARCHAR/VARCHAR2/CHAR → 수신 DATE: 경고

#### 4.2.3 크기 비교
- VARCHAR 계열 타입만 비교
- 송신 크기 > 수신 크기인 경우 경고
- DATE 타입은 크기 비교 제외

#### 4.2.4 NULL 허용 여부
- 송신 NULL 허용 + 수신 NOT NULL = 경고 (제약조건 위반 가능성)

### 4.3 출력 Excel 파일 구조
각 인터페이스별 시트 생성:
1. **인터페이스 기본 정보**
   - 인터페이스명, ID
   - 송신/수신 테이블 정보

2. **컬럼 비교 결과**
   - 송신 컬럼명, 타입, 크기, NULL 여부
   - 수신 컬럼명, 타입, 크기, NULL 여부
   - 비교 결과 및 상태

### 4.4 시각적 표시
- **녹색**: 정상
- **빨간색**: 오류 (컬럼 미존재, 심각한 타입 불일치 등)
- **노란색**: 경고 (DATE-VARCHAR 변환, 크기 차이 등)

## 5. 파일 구조 설계

### 5.1 파일 분리 전략
단일화된 구조를 유지하면서도 관리성을 위해 2-3개의 파일로 분리:

1. **bw_xltest_core.py** - 핵심 비즈니스 로직
   - ColumnValidator 클래스: 컬럼 검증 규칙
   - DatabaseHandler 클래스: DB 연결 및 메타데이터 조회

2. **bw_xltest_io.py** - 입출력 처리
   - ExcelReader 클래스: Excel 입력 파일 읽기
   - ExcelWriter 클래스: 결과 Excel 파일 생성

3. **bw_xltest.py** - 메인 실행 파일
   - 전체 프로세스 조정
   - 에러 처리 및 로깅
   - 사용자 인터페이스

### 5.2 클래스 구조
```
bw_xltest_core.py:
├── ColumnValidator      # 컬럼 검증 로직
└── DatabaseHandler      # DB 연결 및 메타데이터 조회

bw_xltest_io.py:
├── ExcelReader         # Excel 파일 읽기
└── ExcelWriter         # 결과 Excel 생성

bw_xltest.py:
└── main()              # 메인 실행 함수
```

### 5.3 에러 처리 전략
- 각 인터페이스별 독립적 처리 (한 인터페이스 오류가 전체 중단시키지 않음)
- 상세한 에러 로깅
- 사용자 친화적 에러 메시지

## 6. 작업 리스트

### 6.1 Phase 1: bw_xltest_core.py 구현
- [ ] 기본 구조 설정
  - [ ] 필요한 import 문
  - [ ] 상수 정의 (Oracle Client 경로 등)
  
- [ ] DatabaseHandler 클래스
  - [ ] __init__ 메서드 (Oracle Client 초기화)
  - [ ] connect_db 메서드 (DB 연결)
  - [ ] get_column_info 메서드 (테이블 메타데이터 조회)
  - [ ] close_connections 메서드 (연결 종료)
  
- [ ] ColumnValidator 클래스
  - [ ] check_column_exists 메서드 (존재 여부 검사)
  - [ ] check_type_compatibility 메서드 (타입 호환성 검사)
  - [ ] check_date_varchar_conversion 메서드 (DATE-VARCHAR 변환 감지)
  - [ ] check_size_compatibility 메서드 (크기 비교)
  - [ ] check_nullable_compatibility 메서드 (NULL 허용 여부 비교)
  - [ ] validate_columns 메서드 (전체 검증 수행)

### 6.2 Phase 2: bw_xltest_io.py 구현
- [ ] ExcelReader 클래스
  - [ ] read_interface_info 메서드 (인터페이스 정보 읽기)
  - [ ] parse_db_info 메서드 (DB 연결 정보 파싱)
  - [ ] parse_table_info 메서드 (테이블 정보 파싱)
  - [ ] read_column_mappings 메서드 (컬럼 매핑 읽기)
  - [ ] validate_input_format 메서드 (입력 형식 검증)
  
- [ ] ExcelWriter 클래스
  - [ ] create_workbook 메서드 (워크북 생성)
  - [ ] write_interface_info 메서드 (인터페이스 정보 작성)
  - [ ] write_comparison_results 메서드 (비교 결과 작성)
  - [ ] apply_conditional_formatting 메서드 (조건부 서식 적용)
  - [ ] apply_styles 메서드 (스타일 설정)
  - [ ] save_workbook 메서드 (파일 저장)

### 6.3 Phase 3: bw_xltest.py 구현
- [ ] 설정 상수 정의
  - [ ] 입력/출력 파일명
  - [ ] 로깅 설정
  
- [ ] process_interface 함수
  - [ ] 단일 인터페이스 처리 로직
  - [ ] 에러 처리
  
- [ ] main 함수
  - [ ] 전체 프로세스 조정
  - [ ] 진행 상황 출력
  - [ ] 결과 요약 출력
  
- [ ] 명령행 인터페이스
  - [ ] 인자 파싱
  - [ ] 도움말 제공

### 6.4 Phase 4: 통합 및 테스트
- [ ] 모듈 간 통합 테스트
- [ ] 다양한 시나리오 테스트
  - [ ] 정상 케이스
  - [ ] DATE-VARCHAR 변환 케이스
  - [ ] 컬럼 누락 케이스
  - [ ] DB 연결 실패 케이스
  
- [ ] 성능 최적화
  - [ ] DB 연결 재사용
  - [ ] 대용량 데이터 처리

### 6.5 Phase 5: 문서화
- [ ] 코드 주석 추가
- [ ] 사용 예제 작성
- [ ] README 파일 작성

## 7. 기술 스택
- Python 3.x
- openpyxl (Excel 처리)
- oracledb (Oracle DB 연결)
- ast (딕셔너리 문자열 파싱)
- logging (로깅)

## 8. 예상 파일 구조
```
cursor_edu/
├── bw_xltest.py          # 메인 실행 파일
├── bw_xltest_core.py     # 핵심 비즈니스 로직
├── bw_xltest_io.py       # Excel 입출력 처리
├── input.xlsx            # 입력 파일
└── output_bw.xlsx        # 출력 파일
```

## 9. 개선사항
1. **단순화된 구조**: 2-3개 파일로 관리하기 쉬운 구조
2. **DATE-VARCHAR 변환 감지**: 타입 변환 문제를 명시적으로 검출
3. **견고한 에러 처리**: 인터페이스별 독립 처리로 안정성 향상
4. **명확한 책임 분리**: 핵심 로직, 입출력, 실행 부분 분리
5. **확장 가능성**: 새로운 검증 규칙 추가 용이

## 10. 주의사항
- Oracle Instant Client 설치 필요
- DB 접속 권한 필요 (all_tab_columns 조회 권한)
- Excel 파일 형식 준수 필요
- Python 3.x 이상 필요

## 11. 구현 완료 내역

### 11.1 생성된 파일
1. **bw_xltest_core.py** (383줄)
   - DatabaseHandler 클래스: DB 연결 및 메타데이터 조회
   - ColumnValidator 클래스: 컬럼 검증 로직

2. **bw_xltest_io.py** (354줄)
   - ExcelReader 클래스: Excel 입력 파일 처리
   - ExcelWriter 클래스: 결과 Excel 파일 생성

3. **bw_xltest.py** (241줄)
   - 메인 실행 파일
   - 전체 프로세스 조정 및 에러 처리

### 11.2 주요 구현 기능
#### DatabaseHandler (bw_xltest_core.py)
- `connect_db()`: Oracle DB 연결
- `get_column_info()`: 테이블 메타데이터 조회
- `close_connections()`: 연결 종료

#### ColumnValidator (bw_xltest_core.py)
- `check_column_exists()`: 컬럼 존재 여부 검사
- `check_type_compatibility()`: 타입 호환성 검사
- `check_date_varchar_conversion()`: DATE-VARCHAR 변환 감지
- `check_size_compatibility()`: 크기 호환성 검사
- `check_nullable_compatibility()`: NULL 허용 여부 검사
- `validate_columns()`: 전체 검증 수행

#### ExcelReader (bw_xltest_io.py)
- `read_interface_info()`: 인터페이스 정보 읽기
- `parse_db_info()`: DB 연결 정보 파싱
- `parse_table_info()`: 테이블 정보 파싱
- `read_column_mappings()`: 컬럼 매핑 읽기

#### ExcelWriter (bw_xltest_io.py)
- `write_interface_result()`: 검증 결과 작성
- `_write_interface_info()`: 인터페이스 기본 정보 작성
- `_write_comparison_results()`: 컬럼 비교 결과 작성
- 조건부 서식 적용 (정상=녹색, 경고=노란색, 오류=빨간색)

### 11.3 사용 방법
```bash
# 기본 실행 (input.xlsx → output_bw.xlsx)
python bw_xltest.py

# 파일명 지정
python bw_xltest.py myinput.xlsx myoutput.xlsx

# 도움말
python bw_xltest.py -h
```

### 11.4 로깅 및 에러 처리
- 상세한 로깅으로 실행 과정 추적 가능
- 각 인터페이스별 독립적 처리로 부분 실패 시에도 전체 중단 방지
- 최종 결과 요약 제공 (성공/실패 건수, 처리 시간 등)

### 11.5 검증 규칙 상세
1. **컬럼 존재 여부**: 매핑된 컬럼이 실제 테이블에 존재하는지 확인
2. **타입 호환성**: 
   - 동일 타입 허용
   - VARCHAR, VARCHAR2, CHAR는 서로 호환
3. **DATE-VARCHAR 변환**: 
   - DATE → VARCHAR: "날짜 형식 확인 필요" 경고
   - VARCHAR → DATE: "날짜 형식 검증 필요" 경고
4. **크기 비교**: VARCHAR 계열에서 송신 > 수신인 경우 경고
5. **NULL 허용**: 송신 NULL 허용 + 수신 NOT NULL = 제약조건 위반 경고

### 11.6 출력 Excel 파일 구조
각 인터페이스별로 독립된 시트 생성:
- **헤더 섹션**: 인터페이스명, ID, 송신/수신 테이블 정보
- **비교 결과 테이블**: 
  - 송신 컬럼 정보 (컬럼명, 타입, 크기, NULL 여부)
  - 수신 컬럼 정보 (컬럼명, 타입, 크기, NULL 여부)
  - 비교 결과 메시지
  - 상태 (색상으로 구분)

## 12. 향후 개선 가능 사항
- 다양한 데이터베이스 지원 (PostgreSQL, MySQL 등)
- 웹 인터페이스 제공
- 검증 규칙 커스터마이징 기능
- 배치 처리 성능 최적화
- 검증 결과 리포트 다양화 (PDF, HTML 등)