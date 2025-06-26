# iflist_to.py 문서

## 개요
`iflist_to.py`는 엑셀 파일의 TEST 환경 경로를 PROD 환경 경로로 변환하고, 파일 존재 여부 및 디렉토리 정보를 확인하여 새로운 엑셀과 YAML 파일을 생성하는 도구입니다.

## 2024년 12월 업데이트 사항
1. **행 구조 변경**: 송신과 수신 데이터를 별도의 행으로 분리
2. **구분 컬럼 추가**: 첫 번째 컬럼에 "구분" 추가하여 송신/수신 표시
3. **컬럼명 정규화**: 모든 컬럼명에서 "송신"/"수신" 제거
4. **생성여부 표시 개선**: 1이 아닌 모든 값을 0으로 표시 (필터링 용이)
5. **PROD_SOURCE 필터링**: 입력 데이터에 PROD_SOURCE가 포함된 행은 자동으로 제외

## 주요 기능

### 1. 경로 변환 및 검증
- 엑셀의 4개 컬럼 대상 처리:
  - 송신파일경로
  - 수신파일경로
  - 송신스키마파일명
  - 수신스키마파일명
- 각 컬럼별 조건부 처리 (생성여부가 "1"인 경우만)
- `XX_TEST_SOURCE` → `XX_PROD_SOURCE` 자동 변환
- 변환된 파일 존재 여부 확인
- 디렉토리 내 전체 파일 개수 카운트

### 2. 출력 파일 생성

#### 2.1 엑셀 출력 (`iflist_to.xlsx`)
- **구조 변경 (2024.12)**:
  - 첫 번째 컬럼: "구분" (송신/수신)
  - 송신과 수신 데이터가 별도의 행으로 분리
  - 컬럼명에서 "송신"/"수신" 제거 (예: "송신파일경로" → "파일경로")
- 원본 컬럼 + 3개 파생 컬럼 생성
  - `[컬럼명]PROD`: TEST→PROD 변환된 경로
  - `[생성여부]PROD`: 파일 존재 시 "1", 없으면 "0"
  - `DFPROD` / `스키마DFPROD`: 디렉토리 파일 개수 ("X": 디렉토리 없음, 숫자: 파일 개수)
- 생성여부: 1이 아닌 모든 값은 0으로 표시

#### 2.2 YAML 출력 (`iflist_to.yaml`)
```yaml
files:
  - source: "C:\\BwProject\\LH_TEST_SOURCE\\AAA\\BBB\\CCC.process"
    destination: "C:\\BwProject\\LH_PROD_SOURCE\\AAA\\BBB\\CCC.process"
  - source: "C:\\BwProject\\LH_TEST_SOURCE\\AAA\\BBB\\DDD.xsd"
    destination: "C:\\BwProject\\LH_PROD_SOURCE\\AAA\\BBB\\DDD.xsd"
```

### 3. 파일 복사 기능
- YAML 파일 기반 일괄 복사
- 기존 파일 존재 시 skip (overwrite 방지)
- 상세 로그 생성
  - 성공: `복사 성공: [원본] → [대상]`
  - 에러: `[ERROR]` 접두어 사용
    - `[ERROR] 파일이 이미 존재: [경로]`
    - `[ERROR] 원본 파일 없음: [경로]`

## 사용 방법

### 1. 프로그램 실행
```bash
python iflist_to.py
```

### 2. 메뉴 옵션
```
=== 파일 경로 변환 도구 ===
1. 엑셀 분석 및 YAML 생성
2. 파일 복사 실행
0. 종료
```

### 3. 작업 흐름
1. **1번 메뉴**: 입력 엑셀 파일 지정 → `iflist_to.xlsx`와 `iflist_to.yaml` 생성
2. **2번 메뉴**: YAML 파일 기반으로 파일 복사 수행 → 타임스탬프가 포함된 로그 파일 생성

## 파일 구조

### 입력 엑셀 파일 필수 컬럼
- 송신파일경로, 송신파일생성여부
- 수신파일경로, 수신파일생성여부
- 송신스키마파일명, 송신스키마파일생성여부
- 수신스키마파일명, 수신스키마파일생성여부

### 출력 엑셀 파일 컬럼 구조 (2024.12 업데이트)
- 헤더: 구분, 파일경로, 파일경로PROD, 파일생성여부, 파일생성여부PROD, DFPROD, 스키마파일명, 스키마파일명PROD, 스키마파일생성여부, 스키마파일생성여부PROD, 스키마DFPROD
- 데이터 행:
  - 2행: 첫 번째 데이터의 송신 정보 (구분="송신")
  - 3행: 첫 번째 데이터의 수신 정보 (구분="수신")
  - 4행: 두 번째 데이터의 송신 정보 (구분="송신")
  - 5행: 두 번째 데이터의 수신 정보 (구분="수신")
  - 이후 반복

## 주요 함수

### `process_file_path(file_path, check_flag)`
- 파일 경로를 TEST에서 PROD로 변환
- 파일 존재 여부 및 디렉토리 정보 수집
- 반환: (PROD 경로, 파일 존재 여부, 디렉토리 파일 개수)

### `generate_excel_and_yaml(input_excel_path, output_excel_path, output_yaml_path)`
- 입력 엑셀 파일 읽기 및 처리
- TEST→PROD 변환 수행
- 결과를 엑셀과 YAML로 저장

### `execute_file_copy(yaml_path, log_path)`
- YAML 파일의 매핑 정보로 파일 복사
- overwrite 방지 (기존 파일 skip)
- 상세 로그 생성

### `save_excel_with_style(df, excel_path)`
- 엑셀 파일에 스타일 적용
- 헤더 배경색 (연한 파란색)
- 폰트 크기 10pt
- 컬럼 너비 자동 조절

## 로그 예시

```
[2024-01-15 10:00:00] 파일 복사 시작
[2024-01-15 10:00:01] 복사 성공: C:\BwProject\LH_TEST_SOURCE\A\a.process → C:\BwProject\LH_PROD_SOURCE\A\a.process
[2024-01-15 10:00:02] [ERROR] 파일이 이미 존재: C:\BwProject\LH_PROD_SOURCE\B\b.process
[2024-01-15 10:00:03] [ERROR] 원본 파일 없음: C:\BwProject\LH_TEST_SOURCE\C\c.process

[2024-01-15 10:00:04] 파일 복사 완료
성공: 1개, 건너뜀: 1개, 오류: 1개
```

## 특이사항 처리

1. **경로 변환 규칙**
   - 두 번째 디렉토리만 변환 (예: `C:\BwProject\LH_TEST_SOURCE\...` → `C:\BwProject\LH_PROD_SOURCE\...`)
   - Windows 경로 형식 사용 (`\`)

2. **파일 복사 정책**
   - 기존 파일이 있으면 덮어쓰지 않음 (skip)
   - 원본 파일이 없으면 에러로 처리
   - 대상 디렉토리는 자동 생성

3. **YAML 사용 목적**
   - 파일 이름 변경이 필요한 특수한 경우 대비
   - 생성 목록에서 특정 파일 제외 가능
   - 복사 작업의 유연성 확보
   - **독립적 운영**: YAML 파일을 직접 편집하여 복사 경로 변경 가능 (엑셀과 무관하게 작동)

4. **PROD_SOURCE 필터링 (2024.12 추가)**
   - 입력 엑셀에서 파일경로나 스키마파일명에 'PROD_SOURCE'가 포함된 행은 자동 제외
   - 모든 출력(엑셀, YAML)에서 제외됨
   - 디버그 모드에서 건너뛴 행 정보 출력

## 관련 파일
- `string_replacer.py`: 복잡한 파일 내용 치환 도구 (정규식 기반)
- `iflist05.py`: 단순 문자열 치환 도구

## 버전 정보
- 작성일: 2024년 (현재)
- Python 3.x 필요
- 필요 라이브러리: pandas, openpyxl, pyyaml