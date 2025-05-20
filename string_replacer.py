import openpyxl
import yaml
import difflib
import os
import datetime
import pandas as pd
import shutil
import re

# 디버그 모드 설정
DEBUG_MODE = True  # 디버그 정보 출력 여부를 제어하는 플래그

def debug_print(*args, **kwargs):
    """디버그 모드일 때만 메시지를 출력하는 함수"""
    if DEBUG_MODE:
        print("[DEBUG]", *args, **kwargs)

# chardet 모듈 가져오기 시도 (설치되지 않았을 경우 대비)
try:
    import chardet
    HAS_CHARDET = True
except ImportError:
    print("경고: chardet 모듈을 찾을 수 없습니다. 기본 인코딩(utf-8)을 사용합니다.")
    print("pip install chardet 명령어로 설치할 수 있습니다.")
    HAS_CHARDET = False

def generate_yaml_from_excel(excel_path, yaml_path):
    """엑셀 파일을 읽어 YAML 파일을 생성한다."""
    # pandas로 엑셀 파일 읽기
    df = pd.read_excel(excel_path, engine='openpyxl')
    
    # 전체 YAML 구조를 저장할 딕셔너리
    full_yaml_structure = {}
    
    # 2행씩 처리 (일반행, 매칭행)
    for i in range(0, len(df), 2):
        if i + 1 >= len(df):  # 마지막 행이 홀수인 경우 처리
            break
            
        normal_row = df.iloc[i]  # 일반행
        match_row = df.iloc[i+1]  # 매칭행
        
        print(f"\n=== {i//2 + 1}번째 행 쌍 ===")
        
        # 파일 생성 여부 확인 및 YAML 구조 생성
        yaml_structure = {
            f"{i//2 + 1}번째 행": {}
        }
        
        def modify_path(path):
            """파일 경로를 수정하는 함수 (테스트용)"""
            if isinstance(path, str) and path.startswith("C:\\BwProject"):
                return path.replace("C:\\BwProject", "C:\\TBwProject")
            return path

        def extract_filename(path):
            """파일 경로에서 파일명만 추출"""
            if not isinstance(path, str):
                return ""
            return os.path.basename(path)

        def extract_process_filename(path):
            """프로세스 파일 경로에서 파일명만 추출"""
            if not isinstance(path, str):
                return ""
            return os.path.basename(path)

        def process_schema_path(schema_path):
            """스키마 파일 경로를 처리하여 namespace와 schemaLocation 생성"""
            if not isinstance(schema_path, str):
                return None, None
            
            # 1. 경로 구분자 변경
            normalized_path = schema_path.replace('\\', '/')
            
            # 2. '/SharedResources' 이후 부분 추출
            shared_idx = normalized_path.find('/SharedResources')
            if shared_idx == -1:
                return None, None
                
            # BB 부분을 포함한 경로 추출
            bb_start = normalized_path.rfind('/', 0, shared_idx)
            if bb_start == -1:
                return None, None
                
            relative_path = normalized_path[bb_start:]  # /BB/SharedResources/...
            schema_location = relative_path[relative_path.find('/SharedResources'):]  # /SharedResources/...
            namespace = f"http://www.tibco.com/schemas{relative_path}"
            
            return namespace, schema_location

        def create_schema_replacements(filename, schema_path):
            """스키마 파일 치환 목록 생성"""
            if not filename.endswith('.xsd'):
                return []
            
            namespace, schema_location = process_schema_path(schema_path)
            if not namespace or not schema_location:
                return []
                
            base_name = os.path.splitext(filename)[0]
            return [{
                "설명": "스키마 namespace 치환",
                "조건": {
                    "파일명패턴": filename,
                    "태그": "xsd:import",
                    "속성": ["namespace"]
                },
                "찾기": {
                    "정규식": f'(\\bnamespace\\s*=\\s*")[^"]*{base_name}[^"]*(")'
                },
                "교체": {
                    "값": namespace
                }
            },
            {
                "설명": "스키마 schemaLocation 치환",
                "조건": {
                    "파일명패턴": filename,
                    "태그": "xsd:import",
                    "속성": ["schemaLocation"]
                },
                "찾기": {
                    "정규식": f'(\\bschemaLocation\\s*=\\s*")[^"]*{base_name}[^"]*(")'
                },
                "교체": {
                    "값": schema_location
                }
            },
            {
                "설명": "ProcessDefinition namespace 치환",
                "조건": {
                    "파일명패턴": filename,
                    "태그": "pd:ProcessDefinition",
                    "속성": ["xmlns:pfx3"]
                },
                "찾기": {
                    "정규식": f'(\\bxmlns:pfx3\\s*=\\s*")[^"]*{base_name}[^"]*(")'
                },
                "교체": {
                    "값": namespace
                }
            }]
        
        def extract_process_path(file_path):
            """프로세스 파일 경로에서 'Processes' 이후의 경로를 추출하고 디렉토리 구분자를 변경"""
            if not isinstance(file_path, str):
                return ""
            
            # 디렉토리 구분자를 '/'로 통일
            normalized_path = file_path.replace('\\', '/')
            
            # 'Processes' 위치 찾기
            processes_idx = normalized_path.find('Processes/')
            if processes_idx == -1:
                return ""
            
            # 'Processes/' 이후의 경로 추출
            relative_path = normalized_path[processes_idx + len('Processes/'):]
            
            return relative_path

        def create_process_replacements(source_path, target_path, match_row, normal_row):
            """프로세스 파일의 치환 목록 생성
            source_path: 매칭행의 파일 경로 (패턴 매칭용)
            target_path: 기본행의 파일 경로 (교체용)
            match_row: 매칭행 데이터
            normal_row: 기본행 데이터
            """
            if not isinstance(source_path, str) or not isinstance(target_path, str):
                return []
            
            # 매칭행의 파일명으로 패턴 매칭 (찾을 패턴)
            source_filename = extract_process_filename(source_path)
            # 기본행의 경로에서 Processes 이후 경로 추출 (교체할 값)
            target_process_path = extract_process_path(target_path)
            
            if not source_filename or not target_process_path:
                return []
            
            replacements = [{
                "설명": "프로세스 이름 치환",
                "조건": {
                    "파일명패턴": source_filename,  # 매칭행의 파일명으로 패턴 매칭
                    "태그": "pd:ProcessDefinition/pd:name",
                    "속성": []
                },
                "찾기": {
                    "정규식": f'(<pd:name>Processes/)[^<]*(</pd:name>)'
                },
                "교체": {
                    "값": target_process_path  # Processes 이후의 전체 경로로 교체
                }
            }]

            # 고정 문자열 치환 규칙 추가
            fixed_replacements = [
                {
                    "설명": "LHMES_MGR 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": "LHMES_MGR"
                    },
                    "교체": {
                        "값": "LYMES_MGR"
                    }
                },
                {
                    "설명": "VOMES_MGR 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": "VOMES_MGR"
                    },
                    "교체": {
                        "값": "LZMES_MGR"
                    }
                },
                {
                    "설명": "LH 문자열 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": "'LH'"
                    },
                    "교체": {
                        "값": "'LY'"
                    }
                },
                {
                    "설명": "VO 문자열 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": "'VO'"
                    },
                    "교체": {
                        "값": "'LZ'"
                    }
                }
            ]
            replacements.extend(fixed_replacements)

            # IFID 치환 규칙 추가
            origin_ifid = f"{match_row['Group ID']}.{match_row['Event_ID']}"
            dest_ifid = f"{normal_row['Group ID']}.{match_row['Event_ID']}"
            
            # IFID가 다른 경우에만 치환 규칙 추가
            if origin_ifid != dest_ifid:
                replacements.append({
                    "설명": "IFID 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": origin_ifid.replace(".", "\\.")  # 점(.)을 이스케이프
                    },
                    "교체": {
                        "값": dest_ifid
                    }
                })
            
            # Event_ID 치환 규칙 추가
            if match_row['Event_ID'] != normal_row['Event_ID']:
                replacements.append({
                    "설명": "Event_ID 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": match_row['Event_ID']
                    },
                    "교체": {
                        "값": normal_row['Event_ID']
                    }
                })

            # Group ID &quot; 형식 치환 규칙 추가
            if match_row['Group ID'] != normal_row['Group ID']:
                replacements.append({
                    "설명": "Group ID &quot; 형식 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": f'&quot;{match_row["Group ID"]}&quot;'
                    },
                    "교체": {
                        "값": f'&quot;{normal_row["Group ID"]}&quot;'
                    }
                })

            # 송신업무명 &quot; 형식 치환 규칙 추가
            if match_row['송신\n업무명'] != normal_row['송신\n업무명']:
                replacements.append({
                    "설명": "송신업무명 &quot; 형식 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": f'&quot;{match_row["송신\n업무명"]}&quot;'
                    },
                    "교체": {
                        "값": f'&quot;{normal_row["송신\n업무명"]}&quot;'
                    }
                })

            # 수신업무명 &quot; 형식 치환 규칙 추가
            if match_row['수신\n업무명'] != normal_row['수신\n업무명']:
                replacements.append({
                    "설명": "수신업무명 &quot; 형식 치환",
                    "조건": {
                        "파일명패턴": source_filename
                    },
                    "찾기": {
                        "정규식": f'&quot;{match_row["수신\n업무명"]}&quot;'
                    },
                    "교체": {
                        "값": f'&quot;{normal_row["수신\n업무명"]}&quot;'
                    }
                })

            # 송신업무명 치환 규칙 추가 (pd:from/to Check 형식)
            if match_row['송신\n업무명'] != normal_row['송신\n업무명']:
                # pd:from 태그 치환
                replacements.append({
                    "설명": "송신업무명 from 태그 치환",
                    "조건": {
                        "파일명패턴": source_filename,
                        "태그": "pd:from"
                    },
                    "찾기": {
                        "정규식": f'(<pd:from>Check {match_row["송신\n업무명"]})'
                    },
                    "교체": {
                        "값": f'<pd:from>Check {normal_row["송신\n업무명"]}'
                    }
                })
                # pd:to 태그 치환
                replacements.append({
                    "설명": "송신업무명 to 태그 치환",
                    "조건": {
                        "파일명패턴": source_filename,
                        "태그": "pd:to"
                    },
                    "찾기": {
                        "정규식": f'(<pd:to>Check {match_row["송신\n업무명"]})'
                    },
                    "교체": {
                        "값": f'<pd:to>Check {normal_row["송신\n업무명"]}'
                    }
                })

            # 수신업무명 치환 규칙 추가 (pd:from/to Check 형식)
            if match_row['수신\n업무명'] != normal_row['수신\n업무명']:
                # pd:from 태그 치환
                replacements.append({
                    "설명": "수신업무명 from 태그 치환",
                    "조건": {
                        "파일명패턴": source_filename,
                        "태그": "pd:from"
                    },
                    "찾기": {
                        "정규식": f'(<pd:from>Check {match_row["수신\n업무명"]})'
                    },
                    "교체": {
                        "값": f'<pd:from>Check {normal_row["수신\n업무명"]}'
                    }
                })
                # pd:to 태그 치환
                replacements.append({
                    "설명": "수신업무명 to 태그 치환",
                    "조건": {
                        "파일명패턴": source_filename,
                        "태그": "pd:to"
                    },
                    "찾기": {
                        "정규식": f'(<pd:to>Check {match_row["수신\n업무명"]})'
                    },
                    "교체": {
                        "값": f'<pd:to>Check {normal_row["수신\n업무명"]}'
                    }
                })
            
            return replacements

        # 1. 송신파일경로 처리
        if pd.notna(normal_row.get('송신파일생성여부')) and float(normal_row['송신파일생성여부']) == 1.0:
            yaml_structure[f"{i//2 + 1}번째 행"]["송신파일경로"] = {
                "원본파일": match_row['송신파일경로'],
                "복사파일": modify_path(normal_row['송신파일경로']),  # 경로 수정
                "치환목록": create_schema_replacements(
                    extract_filename(normal_row['송신스키마파일명']),
                    normal_row['송신스키마파일명']
                ) + create_process_replacements(
                    match_row['송신파일경로'],    # 매칭행의 경로로 패턴 매칭
                    normal_row['송신파일경로'],    # 기본행의 경로로 교체
                    match_row,
                    normal_row
                )
            }
            print("\n[송신파일경로 생성]")
            print(f"  원본파일: {match_row['송신파일경로']}")
            print(f"  복사파일: {modify_path(normal_row['송신파일경로'])}")
        
        # 2. 수신파일경로 처리
        if pd.notna(normal_row.get('수신파일생성여부')) and float(normal_row['수신파일생성여부']) == 1.0:
            yaml_structure[f"{i//2 + 1}번째 행"]["수신파일경로"] = {
                "원본파일": match_row['수신파일경로'],
                "복사파일": modify_path(normal_row['수신파일경로']),  # 경로 수정
                "치환목록": create_schema_replacements(
                    extract_filename(normal_row['수신스키마파일명']),
                    normal_row['수신스키마파일명']
                ) + create_process_replacements(
                    match_row['수신파일경로'],    # 매칭행의 경로로 패턴 매칭
                    normal_row['수신파일경로'],    # 기본행의 경로로 교체
                    match_row,
                    normal_row
                )
            }
            print("\n[수신파일경로 생성]")
            print(f"  원본파일: {match_row['수신파일경로']}")
            print(f"  복사파일: {modify_path(normal_row['수신파일경로'])}")
        
        # 3. 송신스키마파일명 처리
        if pd.notna(normal_row.get('송신스키마파일생성여부')) and float(normal_row['송신스키마파일생성여부']) == 1.0:
            yaml_structure[f"{i//2 + 1}번째 행"]["송신스키마파일명"] = {
                "원본파일": match_row['송신스키마파일명'],
                "복사파일": modify_path(normal_row['송신스키마파일명'])  # 경로 수정
            }
            print("\n[송신스키마파일명 생성]")
            print(f"  원본파일: {match_row['송신스키마파일명']}")
            print(f"  복사파일: {modify_path(normal_row['송신스키마파일명'])}")
        
        # 4. 수신스키마파일명 처리
        if pd.notna(normal_row.get('수신스키마파일생성여부')) and float(normal_row['수신스키마파일생성여부']) == 1.0:
            yaml_structure[f"{i//2 + 1}번째 행"]["수신스키마파일명"] = {
                "원본파일": match_row['수신스키마파일명'],
                "복사파일": modify_path(normal_row['수신스키마파일명'])  # 경로 수정
            }
            print("\n[수신스키마파일명 생성]")
            print(f"  원본파일: {match_row['수신스키마파일명']}")
            print(f"  복사파일: {modify_path(normal_row['수신스키마파일명'])}")
        
        # YAML 구조 출력
        print("\n[YAML 구조]")
        print(yaml.dump(yaml_structure, allow_unicode=True, sort_keys=False))
        print("=" * 50)
        
        # 전체 YAML 구조에 현재 구조 추가
        full_yaml_structure.update(yaml_structure)
    
    # YAML 파일 생성
    try:
        with open(yaml_path, 'w', encoding='utf-8') as yf:
            yaml.dump(full_yaml_structure, yf, allow_unicode=True, sort_keys=False)
        print(f"\nYAML 파일이 생성되었습니다: {yaml_path}")
        return len(full_yaml_structure)  # 생성된 작업 수 반환
    except Exception as e:
        print(f"\nYAML 파일 생성 중 오류 발생: {str(e)}")
        return 0

def apply_replacements(text, replacements):
    """여러 치환 규칙을 순차적으로 적용하여 새로운 텍스트 반환."""
    new_text = text
    for repl in replacements:
        # 단순 문자열 치환 (정규식의 경우 필요시 re.sub로 변경 가능)
        from_str = repl.get("from", "")
        to_str = repl.get("to", "")
        if from_str:
            new_text = new_text.replace(from_str, to_str)
    return new_text

def compute_diff(original_text, modified_text, fromfile="[Before]", tofile="[After]"):
    """두 텍스트 버전에 대한 unified diff를 생성하여 리스트로 반환."""
    original_lines = original_text.splitlines(keepends=True)
    modified_lines = modified_text.splitlines(keepends=True)
    diff_lines = difflib.unified_diff(original_lines, modified_lines,
                                    fromfile=fromfile, tofile=tofile, lineterm='')
    return list(diff_lines)

def preview_diff(yaml_path):
    """YAML 파일에 정의된 각 작업에 대해 diff 미리보기를 콘솔에 출력."""
    try:
        with open(yaml_path, 'r', encoding='utf-8') as yf:
            data = yaml.safe_load(yf)
    except FileNotFoundError:
        print(f"YAML 파일을 찾을 수 없습니다: {yaml_path}")
        return

    jobs = data.get("jobs", []) if data else []
    for job in jobs:
        source = job.get("source")
        dest = job.get("destination", "(preview)")
        replacements = job.get("replacements", [])
        if not source or not replacements:
            continue
        try:
            with open(source, 'r', encoding='utf-8') as sf:
                original_text = sf.read()
        except FileNotFoundError:
            print(f"\n[오류] 원본 파일을 찾을 수 없습니다: {source}")
            continue

        # 치환 적용 (미리보기이므로 파일 저장 안 함)
        modified_text = apply_replacements(original_text, replacements)

        # diff 계산
        diff_lines = compute_diff(original_text, modified_text, fromfile=source, tofile=f"{dest} (preview)")

        # diff 출력
        print(f"\n*** {source} vs {dest} 미리보기 diff ***")
        for line in diff_lines:
            print(line, end='')

def detect_encoding(file_path):
    """파일의 인코딩을 감지합니다."""
    if HAS_CHARDET:
        with open(file_path, 'rb') as f:
            raw = f.read()
            result = chardet.detect(raw)
            return result['encoding']
    else:
        return 'utf-8'  # 기본값으로 utf-8 사용

def copy_file_with_check(source, dest):
    """파일을 복사하되, 대상 파일이 이미 존재하면 경고를 출력합니다."""
    try:
        # 대상 디렉토리가 없으면 생성
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        
        # 대상 파일이 이미 존재하는지 확인
        if os.path.exists(dest):
            print(f"경고: 파일이 이미 존재합니다 - {dest}")
            return False
            
        # 바이너리 모드로 파일 복사
        with open(source, 'rb') as src, open(dest, 'wb') as dst:
            dst.write(src.read())
        print(f"파일 복사 완료: {source} -> {dest}")
        return True
    except Exception as e:
        print(f"파일 복사 중 오류 발생: {str(e)}")
        return False

def apply_schema_replacements(file_path, replacements):
    """파일에 치환 목록을 적용합니다."""
    try:
        # 파일 인코딩 감지
        debug_print(f"\n=== 파일 치환 시작: {file_path} ===")
        encoding = detect_encoding(file_path)
        if not encoding:
            debug_print(f"경고: 파일의 인코딩을 감지할 수 없습니다 - {file_path}")
            encoding = 'utf-8'  # 기본값으로 utf-8 사용
        debug_print(f"감지된 인코딩: {encoding}")
        
        # 바이너리 모드로 파일 읽기
        debug_print("파일 읽기 시작 (바이너리 모드)")
        with open(file_path, 'rb') as f:
            content_bytes = f.read()
            debug_print(f"파일 크기: {len(content_bytes)} bytes")
            
        # 디코딩
        debug_print(f"파일 디코딩 시도 ({encoding})")
        try:
            content = content_bytes.decode(encoding)
            debug_print("디코딩 성공")
        except UnicodeDecodeError as e:
            debug_print(f"디코딩 오류 발생: {str(e)}")
            debug_print("utf-8로 재시도")
            try:
                content = content_bytes.decode('utf-8')
                encoding = 'utf-8'
                debug_print("utf-8 디코딩 성공")
            except UnicodeDecodeError:
                debug_print("utf-8 디코딩 실패, latin-1로 시도")
                content = content_bytes.decode('latin-1')
                encoding = 'latin-1'
                debug_print("latin-1 디코딩 성공")
        
        modified = False
        # 각 치환 규칙 적용
        for idx, repl in enumerate(replacements, 1):
            debug_print(f"\n--- 치환 규칙 {idx}/{len(replacements)} 적용 시도 ---")
            debug_print(f"설명: {repl.get('설명', '설명 없음')}")
            
            pattern = repl['찾기']['정규식']
            replacement = repl['교체']['값']
            
            debug_print(f"정규식 패턴: {pattern}")
            debug_print(f"교체할 값: {replacement}")
            
            # 파일명 패턴 조건 확인
            if '조건' in repl and '파일명패턴' in repl['조건']:
                filename = os.path.basename(file_path)
                if filename != repl['조건']['파일명패턴']:
                    debug_print(f"파일명 패턴 불일치. 건너뜀 (예상: {repl['조건']['파일명패턴']}, 실제: {filename})")
                    continue
                debug_print("파일명 패턴 일치")
            
            # 패턴이 파일에 존재하는지 확인
            matches = list(re.finditer(pattern, content))
            if not matches:
                debug_print("패턴이 파일에서 발견되지 않음")
                continue
            debug_print(f"패턴 매칭 수: {len(matches)}")
            
            # 첫 번째 매칭 내용 출력 (디버깅용)
            if matches and DEBUG_MODE:
                first_match = matches[0]
                debug_print("첫 번째 매칭:")
                debug_print(f"  위치: {first_match.start()}-{first_match.end()}")
                debug_print(f"  내용: {first_match.group()}")
            
            # 정규식에서 캡처 그룹을 사용하는 경우
            if '(' in pattern and ')' in pattern:
                debug_print("캡처 그룹이 있는 패턴 감지됨")
                try:
                    # 첫 번째와 마지막 캡처 그룹을 유지하면서 중간 값만 교체
                    new_content = re.sub(pattern, r'\1' + replacement + r'\2', content)
                    if new_content == content:
                        debug_print("캡처 그룹 방식 치환 후 변경사항 없음")
                    else:
                        debug_print("캡처 그룹 방식 치환 성공")
                except Exception as e:
                    debug_print(f"캡처 그룹 치환 중 오류 발생: {str(e)}")
                    continue
            else:
                debug_print("일반 패턴 감지됨")
                try:
                    # 일반 정규식 치환
                    new_content = re.sub(pattern, replacement, content)
                    if new_content == content:
                        debug_print("일반 방식 치환 후 변경사항 없음")
                    else:
                        debug_print("일반 방식 치환 성공")
                except Exception as e:
                    debug_print(f"일반 치환 중 오류 발생: {str(e)}")
                    continue
                
            if new_content != content:
                content = new_content
                modified = True
                debug_print(f"치환 규칙 {idx} 적용 완료")
            else:
                debug_print(f"치환 규칙 {idx}: 내용이 변경되지 않음")
        
        # 변경된 경우에만 파일 저장
        if modified:
            debug_print("\n파일 저장 시작")
            # 인코딩하여 바이너리 모드로 저장
            content_bytes = content.encode(encoding)
            with open(file_path, 'wb') as f:
                f.write(content_bytes)
            debug_print(f"파일 저장 완료 (크기: {len(content_bytes)} bytes)")
            return True
        else:
            debug_print("\n변경사항이 없어 파일을 저장하지 않음")
            return False
    except Exception as e:
        debug_print(f"치환 작업 중 예외 발생: {str(e)}")
        return False

def execute_replacements(yaml_path, log_path, summary_path):
    """YAML에 정의된 복사 및 치환 작업을 실행하고 로그를 생성합니다."""
    try:
        debug_print(f"YAML 파일 읽기 시작: {yaml_path}")
        with open(yaml_path, 'r', encoding='utf-8') as yf:
            data = yaml.safe_load(yf)
            debug_print(f"YAML 데이터 로드 완료: {len(data) if data else 0}개 항목")
    except FileNotFoundError:
        print(f"YAML 파일을 찾을 수 없습니다: {yaml_path}")
        return
    except Exception as e:
        print(f"YAML 파일 읽기 중 오류 발생: {str(e)}")
        return

    if not data:
        print("실행할 작업이 없습니다.")
        return

    # 로그 파일 초기화
    debug_print(f"로그 파일 초기화: {log_path}")
    with open(log_path, 'w', encoding='utf-8') as lf:
        lf.write(f"[{datetime.datetime.now()}] 작업 시작\n")

    summary_data = []
    total_copies = 0
    total_replacements = 0

    # 각 행 처리
    for row_key, row_data in data.items():
        debug_print(f"\n=== {row_key} 처리 시작 ===")
        debug_print(f"행 데이터: {row_data.keys()}")
        
        # 각 파일 타입 처리 (송신파일경로, 수신파일경로 등)
        for file_type, file_info in row_data.items():
            debug_print(f"\n--- {file_type} 처리 ---")
            source = file_info.get('원본파일')
            dest = file_info.get('복사파일')
            replacements = file_info.get('치환목록', [])
            
            debug_print(f"원본파일: {source}")
            debug_print(f"복사파일: {dest}")
            debug_print(f"치환규칙 수: {len(replacements)}")
            
            if not source or not dest:
                debug_print("원본 또는 대상 파일 경로가 없음, 건너뜀")
                continue
            
            if DEBUG_MODE and replacements:
                debug_print("\n치환 규칙 목록:")
                for idx, repl in enumerate(replacements, 1):
                    debug_print(f"  {idx}. {repl.get('설명', '설명 없음')}")
                    debug_print(f"     조건: {repl.get('조건', {})}")
                    debug_print(f"     찾기: {repl.get('찾기', {})}")
                    debug_print(f"     교체: {repl.get('교체', {})}")
                
            # 1. 파일 복사
            debug_print(f"\n파일 복사 시도: {source} -> {dest}")
            if copy_file_with_check(source, dest):
                total_copies += 1
                debug_print("파일 복사 성공")
                
                # 2. 치환 목록이 있는 경우 치환 수행
                if replacements:
                    debug_print(f"치환 작업 시작: {dest}")
                    if apply_schema_replacements(dest, replacements):
                        total_replacements += 1
                        debug_print("치환 작업 성공")
                    else:
                        debug_print("치환 작업 실패 또는 변경사항 없음")
                else:
                    debug_print("치환 규칙 없음, 건너뜀")
                        
            # 작업 결과 기록
            summary = f"{file_type}: {source} -> {dest}"
            if replacements:
                summary += f" (치환: {len(replacements)}개 규칙)"
            summary_data.append(summary)
            debug_print(f"작업 결과: {summary}")
            
            # 로그 기록
            with open(log_path, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] {summary}\n")

    # 요약 파일 생성
    debug_print(f"\n요약 파일 생성: {summary_path}")
    with open(summary_path, 'w', encoding='utf-8') as sf:
        sf.write("작업 요약:\n")
        for line in summary_data:
            sf.write(line + "\n")
        sf.write(f"\n총 복사 파일 수: {total_copies}")
        sf.write(f"\n총 치환 파일 수: {total_replacements}")

    debug_print("\n=== 전체 작업 완료 ===")
    debug_print(f"총 복사 파일 수: {total_copies}")
    debug_print(f"총 치환 파일 수: {total_replacements}")

    print(f"\n작업이 완료되었습니다.")
    print(f"총 복사 파일 수: {total_copies}")
    print(f"총 치환 파일 수: {total_replacements}")

def main():
    while True:
        print("\n=== 문자열 치환 도구 ===")
        print("1. YAML 생성 (엑셀 -> YAML)")
        print("2. 미리보기 (YAML 기반 diff 출력)")
        print("3. 실행 (파일 복사 및 치환)")
        print("0. 종료")
        
        choice = input("\n원하는 작업을 선택하세요: ").strip()
        
        if choice == "1":
            excel_path = input("엑셀 파일 경로를 입력하세요: ").strip()
            yaml_path = input("생성할 YAML 파일 경로를 입력하세요: ").strip()
            try:
                count = generate_yaml_from_excel(excel_path, yaml_path)
                print(f"YAML 파일이 생성되었습니다. (총 {count}개 작업)")
            except Exception as e:
                print(f"오류 발생: {str(e)}")
        
        elif choice == "2":
            yaml_path = input("YAML 파일 경로를 입력하세요: ").strip()
            preview_diff(yaml_path)
        
        elif choice == "3":
            yaml_path = input("YAML 파일 경로를 입력하세요: ").strip()
            log_path = input("로그 파일 경로를 입력하세요: ").strip()
            summary_path = input("요약 파일 경로를 입력하세요: ").strip()
            execute_replacements(yaml_path, log_path, summary_path)
        
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
        
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")

if __name__ == "__main__":
    main() 