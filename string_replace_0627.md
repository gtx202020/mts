# string_replace_0627.py 작업 기록

## 작업 개요
string_replace_0627.py는 iflist03b.py에서 생성된 Excel 파일을 읽어 YAML 치환 규칙을 생성하는 프로그램입니다.

## 주요 작업 이력

### 1. 초기 분석 및 규칙 단순화
- **날짜**: 2025-06-27
- **작업**: 기존 string_replacer.py 기반으로 복잡한 치환 규칙들을 단순화
- **변경사항**: 
  - 고정 문자열 치환 규칙 대부분 제거
  - 4개 핵심 규칙만 유지:
    1. 스키마 namespace 치환
    2. 스키마 schemaLocation 치환  
    3. ProcessDefinition namespace 치환
    4. 프로세스 이름 치환

### 2. RTS_GM 관련 치환 규칙 추가
- **요구사항**: RTS_GM2/RTS_GM 시스템명 변환 처리
- **추가된 규칙**:
  ```python
  # 규칙 1: "Check RTS_GM2" → "Check RTS_GM 2" (프로세스 이름 보호)
  {
      "설명": "Check RTS_GM2 → Check RTS_GM 2 치환",
      "찾기": {
          "정규식": r'([">\s])([Cc]heck\s+)RTS_GM2(["<\s])'
      },
      "교체": {
          "값": r'\1\2RTS_GM 2\3'
      }
  }
  
  # 규칙 2: 전역적으로 RTS_GM → RTS_GM2 (시스템명 변경)
  {
      "설명": "RTS_GM → RTS_GM2 치환",
      "찾기": {
          "정규식": r'RTS_GM(?!2)'
      },
      "교체": {
          "값": "RTS_GM2"
      }
  }
  ```

### 3. 치환 순서 조정
- **문제**: 이중 치환으로 인한 RTS_GM22 현상
- **해결**: 치환 규칙 순서 조정
  - Check RTS_GM2 규칙을 먼저 적용
  - 그 다음 RTS_GM → RTS_GM2 규칙 적용

### 4. XML 태그 처리 개선
- **문제**: `<pd:to>Check RTS_GM2</pd:to>` 형태의 XML 태그 미처리
- **해결**: 정규식 패턴 개선
  ```python
  r'([">\s])([Cc]heck\s+)RTS_GM2(["<\s])'
  ```

### 5. xs:schema 치환 규칙 복원
- **요구사항**: 스키마 파일의 xmlns, targetNamespace 치환 기능 복원
- **복원된 규칙**:
  ```python
  {
      "설명": "xs:schema xmlns 치환",
      "찾기": {
          "정규식": f'xmlns\\s*=\\s*"[^"]*{base_name}[^"]*"'
      },
      "교체": {
          "값": f'xmlns="{namespace}"'
      }
  },
  {
      "설명": "xs:schema targetNamespace 치환",
      "찾기": {
          "정규식": f'targetNamespace\\s*=\\s*"[^"]*{base_name}[^"]*"'
      },
      "교체": {
          "값": f'targetNamespace="{namespace}"'
      }
  }
  ```

### 6. Namespace 조정 로직 구현
- **요구사항**: 파일경로에서 계산된 namespace를 스키마 파일에서 재사용
- **구현 방식**:
  1. 파일경로 처리 시 "스키마 namespace 치환" 규칙에서 namespace 추출
  2. 추출된 namespace를 변수에 저장 (`송신_namespace_from_path`, `수신_namespace_from_path`)
  3. 스키마 파일 처리 시 저장된 namespace 재사용

#### 구현 코드:
```python
# 파일경로에서 계산된 namespace 저장 변수
송신_namespace_from_path = None
수신_namespace_from_path = None

# namespace 추출 로직 (송신파일경로)
for 치환 in 송신_치환목록:
    if 치환.get("설명") == "스키마 namespace 치환":
        namespace_value = 치환["교체"]["값"]
        if 'namespace="' in namespace_value:
            송신_namespace_from_path = namespace_value.split('namespace="')[1].split('"')[0]
        break

# 스키마 파일에서 추출된 namespace 재사용
if 송신_namespace_from_path:
    schema_replacements.extend([
        {
            "설명": "xs:schema xmlns 치환",
            "찾기": {
                "정규식": f'xmlns\\s*=\\s*"[^"]*{base_name}[^"]*"'
            },
            "교체": {
                "값": f'xmlns="{송신_namespace_from_path}"'
            }
        }
    ])
```

## 현재 상태

### 작동 원리
1. Excel 파일을 2행씩 처리 (일반행, 매칭행)
2. 각 행 쌍에 대해 송신/수신 파일경로와 스키마파일명 처리
3. 파일경로에서 namespace 계산 및 저장
4. 스키마 파일에서 저장된 namespace 재사용
5. YAML 형태로 치환 규칙 출력

### 핵심 치환 규칙 (최종)
1. **스키마 namespace 치환**: namespace 속성 값 변경
2. **스키마 schemaLocation 치환**: schemaLocation 속성 값 변경
3. **ProcessDefinition namespace 치환**: xmlns:pfx3 속성 값 변경
4. **프로세스 이름 치환**: 프로세스 파일명 기반 치환
5. **Check RTS_GM2 → Check RTS_GM 2**: 프로세스 이름 보호
6. **RTS_GM → RTS_GM2**: 시스템명 전역 변경
7. **xs:schema xmlns/targetNamespace 치환**: 스키마 파일 내부 namespace 치환

### string_replacer.py와의 차이점
- **namespace 재사용**: 파일경로에서 계산된 namespace를 스키마 파일에서 재사용
- **RTS_GM 처리**: Check 패턴 보호 및 시스템명 변경 로직 추가
- **치환 규칙 단순화**: 불필요한 고정 문자열 치환 제거

## 검증 완료 사항
- ✅ namespace 계산 로직이 string_replacer.py와 동일함 확인
- ✅ 파일경로 "스키마 namespace 치환" 규칙의 교체 값이 올바름 확인  
- ✅ namespace 조정 로직이 의도대로 작동함 확인

## 향후 개선 방향
- 실제 YAML 생성 결과 검증
- 다양한 입력 데이터에 대한 테스트
- 오류 처리 로직 보완

## 참고 파일
- `iflist03b.py`: Excel 파일 생성 (데이터 소스)
- `string_replacer.py`: 원본 참조 구현
- `bwtools_yaml_processor.py`: 관련 YAML 처리 로직