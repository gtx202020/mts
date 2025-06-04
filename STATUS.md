# 프로젝트 상태 보고서

## 프로젝트 개요
TIBCO BW 인터페이스 마이그레이션 자동화 도구 프로젝트로, 'LY'/'LZ' 환경에서 'LH'/'VO' 환경으로의 전환을 지원합니다.

## 주요 구성 요소

### 1. ex_sqlite.py (미구현)
- **목적**: Excel 파일의 내용을 SQLite 데이터베이스(iflist.sqlite)로 변환
- **상태**: 파일이 존재하지 않음 - 구현 필요

### 2. iflist03a.py (구현 완료)
- **기능**: 
  - SQLite DB에서 인터페이스 정보 추출
  - 복사행(기준행)과 원본행 기준으로 2줄 단위 Excel 생성
  - "Unnamed: XX" 컬럼 제외 처리
  - 송신/수신 파일 및 스키마 파일 존재 여부 확인
  - 15가지 비교 검증 수행
- **출력**: 색상 코딩된 검증 결과 Excel 파일

### 3. string_replacer.py (부분 구현)
- **모드 1** (구현 완료): iflist03a.py의 출력 파일에서 YAML 파일 생성
- **모드 2** (미구현): 구현 필요
- **모드 3** (구현 중): 
  - 파일 복사 및 치환 실행
  - 디렉토리 임시 수정 기능 필요
  - 파일 덮어쓰기 수정 필요 (os.path.exists)
- **추가 기능**:
  - YAML 규칙에 따른 치환 작업
  - 작업 로그 2개 생성
  - iflist05.xlsx 결과 파일 생성 (원본파일/복사파일 정보)
  - iflist05_delete.bat 파일 생성 (파일 삭제용 배치)
  - ⚠️ 파일 삭제 배치는 Everything 등에서 인식 문제 있음 - 주의 필요

### 4. test_iflist.py (구현 완료)
- **입력**: 
  - iflist_in.xlsx (인터페이스 정의)
  - iflist03a.py의 출력 Excel
- **기능**: 
  - 매핑 정보와 process 파일 간 일치성 검증
  - 5가지 관점에서 컬럼 매핑 비교
- **출력**: 
  - test_iflist_result.xlsx (검증 결과)
  - 상세 로그 파일

## 데이터 처리 흐름

```
1. Excel → SQLite DB 변환 (ex_sqlite.py) [미구현]
   ↓
2. DB에서 인터페이스 정보 추출 및 검증 (iflist03a.py)
   ↓
3. YAML 생성 및 파일 치환 (string_replacer.py)
   ↓
4. 매핑 정보 최종 검증 (test_iflist.py)
```

## 파일 구조

### 입력 파일
- input.xlsx: 초기 인터페이스 정의
- sender.process: 송신 프로세스 정의 (TIBCO BW)
- receiver.process: 수신 프로세스 정의 (TIBCO BW)

### 출력 파일
- iflist03a_output_sample.csv: 샘플 출력
- iflist05.xlsx: 최종 치환 결과
- test_iflist_result.xlsx: 검증 결과
- iflist05_delete.bat: 생성 파일 삭제용 배치

## 개선 필요 사항

1. **ex_sqlite.py 구현**
   - Excel → SQLite 변환 기능 개발 필요

2. **string_replacer.py 완성**
   - 모드 2 구현
   - 모드 3의 디렉토리 처리 및 파일 덮어쓰기 로직 개선

3. **파일 삭제 배치 안정성**
   - Everything 등 파일 인덱싱 도구와의 충돌 문제 해결

## 치환 규칙 예시

- 'LHMES_MGR' → 'LYMES_MGR'
- 'VOMES_MGR' → 'LZMES_MGR'
- 'LH' → 'LY'
- 'VO' → 'LZ'
- namespace, schemaLocation 등 XML 속성 치환
- IFID, Event_ID, 업무명 등 동적 치환

## 검증 항목

1. Excel 송신 컬럼 vs Process SELECT 컬럼
2. Excel 수신 컬럼 vs Process INSERT 컬럼
3. 송신-수신 컬럼 연결 매핑
4. 송신 스키마 vs Process 송신 컬럼
5. 수신 스키마 vs Process 수신 컬럼

## 프로젝트 상태
- 전체 진행률: 약 70%
- 핵심 기능 대부분 구현 완료
- ex_sqlite.py 구현 및 string_replacer.py 일부 기능 보완 필요