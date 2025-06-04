# 리팩토링된 인터페이스 처리 도구 (RFT - Refactored Tools)

이 디렉토리는 기존의 개별 스크립트들을 리팩토링하여 통합된 도구로 재구성한 결과입니다.

## 📁 파일 구조

### 주요 모듈
- **`rft_ex_sqlite.py`** - Excel 파일을 SQLite 데이터베이스로 변환 (기존 ex_sqlite.py 기능)
- **`rft_interface_processor.py`** - 인터페이스 데이터 처리 및 매칭 (기존 iflist03a.py 기능)
- **`rft_yaml_processor.py`** - YAML 생성 및 파일 치환 처리 (기존 string_replacer.py 기능)
- **`rft_interface_reader.py`** - 인터페이스 정보 읽기 및 BW 파싱 (기존 test_iflist.py 기능)

### 테스트 및 통합
- **`test_rft_modules.py`** - 모든 모듈의 단위 테스트
- **`rft_main.py`** - 통합 실행 파일 (메뉴 방식)

### 문서
- **`README.md`** - 이 파일

## 🚀 빠른 시작

### 1. 통합 도구 실행
```bash
python rft_main.py
```

### 2. 전체 파이프라인 실행 (권장)
1. 통합 도구 실행 후 메뉴에서 `7. 전체 파이프라인 실행` 선택
2. Excel 파일 경로 입력
3. 자동으로 다음 단계들이 실행됩니다:
   - Excel → SQLite 변환
   - 인터페이스 데이터 처리 (LY/LZ ↔ LH/VO 매칭)
   - YAML 규칙 생성

### 3. 개별 모듈 실행
각 모듈을 개별적으로 실행할 수도 있습니다:
```bash
python rft_ex_sqlite.py           # Excel to SQLite 변환
python rft_interface_processor.py # 인터페이스 처리
python rft_yaml_processor.py      # YAML 생성/실행
python rft_interface_reader.py    # 인터페이스 읽기
```

## 📋 주요 기능

### 1. Excel to SQLite 변환 (`rft_ex_sqlite.py`)
- Excel 파일을 SQLite 데이터베이스로 변환
- 테스트용 데이터베이스 생성 기능
- 데이터 검증 및 오류 처리

### 2. 인터페이스 데이터 처리 (`rft_interface_processor.py`)
- SQLite에서 LY/LZ 시스템 필터링
- LH/VO 시스템과 자동 매칭
- 파일 경로 생성 및 존재 여부 확인
- CSV 형태로 결과 출력

### 3. YAML 처리 (`rft_yaml_processor.py`)
- Excel/CSV에서 YAML 규칙 생성
- 파일 복사 및 내용 치환 실행
- 정규식 기반 스마트 치환
- 실행 로그 및 결과 기록

### 4. 인터페이스 정보 읽기 (`rft_interface_reader.py`)
- 특별한 형식의 Excel에서 인터페이스 정보 추출
- TIBCO BW .process 파일 파싱
- INSERT 쿼리 및 파라미터 추출
- CSV 내보내기 지원

## 🧪 테스트

### 전체 테스트 실행
```bash
python test_rft_modules.py
```

### 통합 도구에서 테스트
```bash
python rft_main.py
# 메뉴에서 8. 테스트 실행 선택
```

### 테스트 항목
- Excel to SQLite 변환 테스트
- 테스트 데이터베이스 생성 테스트
- 인터페이스 처리 테스트
- YAML 생성 및 실행 테스트
- 인터페이스 읽기 테스트

## 📊 처리 흐름

```
Excel 파일
    ↓
SQLite DB (iflist.sqlite)
    ↓
인터페이스 처리 (LY/LZ ↔ LH/VO 매칭)
    ↓
CSV 출력 (rft_interface_processed.csv)
    ↓
YAML 규칙 생성 (rft_rules.yaml)
    ↓
파일 복사 및 치환 실행
```

## 🔧 설정 및 커스터마이징

### 기본 파일명 변경
`rft_main.py`에서 다음 변수들을 수정:
```python
self.default_db = "iflist.sqlite"
self.default_output_csv = "rft_interface_processed.csv"
self.default_yaml = "rft_rules.yaml"
```

### 디버그 모드
각 모듈에서 디버그 모드를 활성화할 수 있습니다:
```python
# 상세한 로그 출력
processor = YAMLProcessor(debug_mode=True)
```

## ⚠️ 주의사항

1. **파일 백업**: YAML 실행 시 실제 파일이 복사되고 수정됩니다. 중요한 파일은 미리 백업하세요.

2. **경로 확인**: 파일 경로가 올바른지 확인하고, 필요한 디렉토리가 존재하는지 확인하세요.

3. **테스트 우선**: 새로운 데이터로 작업하기 전에 테스트 기능을 활용하여 동작을 확인하세요.

4. **CSV vs Excel**: 테스트 편의성을 위해 CSV 형태로 출력하도록 기본 설정되어 있습니다.

## 🔍 문제 해결

### 모듈 import 오류
모든 파일이 같은 디렉토리에 있는지 확인하세요.

### 데이터베이스 파일 없음
먼저 "Excel to SQLite 변환" 기능을 실행하세요.

### 파일 경로 오류
절대 경로를 사용하거나, 현재 디렉토리에서 상대 경로가 올바른지 확인하세요.

### 테스트 실패
임시 디렉토리 권한이나 필요한 라이브러리 설치를 확인하세요.

## 📈 개선사항

기존 코드 대비 개선된 점:
- ✅ 모듈화된 구조
- ✅ 포괄적인 테스트 지원
- ✅ 통합된 사용자 인터페이스
- ✅ 오류 처리 강화
- ✅ 로깅 및 디버깅 지원
- ✅ 공통 접두사 (rft_)로 파일명 통일
- ✅ CSV 기반 테스트 지원

## 🤝 기여 방법

1. 새로운 기능이나 버그 수정 시 해당하는 테스트를 추가하세요.
2. 코드 변경 시 `test_rft_modules.py`로 전체 테스트를 실행하세요.
3. 새로운 모듈 추가 시 `rft_` 접두사를 사용하세요.

---

**버전**: 1.0  
**작성일**: 2024년  
**기존 코드**: iflist03a.py, string_replacer.py, test_iflist.py 등을 리팩토링