import openpyxl
import yaml
import difflib
import os
import datetime
import pandas as pd

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
            }]
        
        # 1. 송신파일경로 처리
        if pd.notna(normal_row.get('송신파일생성여부')) and float(normal_row['송신파일생성여부']) == 1.0:
            yaml_structure[f"{i//2 + 1}번째 행"]["송신파일경로"] = {
                "원본파일": match_row['송신파일경로'],
                "복사파일": modify_path(normal_row['송신파일경로']),  # 경로 수정
                "치환목록": create_schema_replacements(
                    extract_filename(normal_row['송신파일경로']),
                    normal_row['송신스키마파일명']  # 송신스키마파일명 사용
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
                    extract_filename(normal_row['수신파일경로']),
                    normal_row['수신스키마파일명']  # 수신스키마파일명 사용
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

def execute_replacements(yaml_path, log_path, summary_path):
    """YAML에 정의된 치환 작업을 실행하고 로그를 생성한다."""
    try:
        with open(yaml_path, 'r', encoding='utf-8') as yf:
            data = yaml.safe_load(yf)
    except FileNotFoundError:
        print(f"YAML 파일을 찾을 수 없습니다: {yaml_path}")
        return

    jobs = data.get("jobs", []) if data else []
    if not jobs:
        print("실행할 작업이 없습니다.")
        return

    # 로그 파일 초기화
    with open(log_path, 'w', encoding='utf-8') as lf:
        lf.write(f"[{datetime.datetime.now()}] 치환 작업 시작 (총 {len(jobs)}개 작업)\n")

    total_replacements = 0
    summary_data = []

    for i, job in enumerate(jobs, 1):
        source = job.get("source")
        dest = job.get("destination")
        replacements = job.get("replacements", [])
        if not source or not dest or not replacements:
            continue

        try:
            with open(source, 'r', encoding='utf-8') as sf:
                original_text = sf.read()
        except FileNotFoundError:
            with open(log_path, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] [오류] 원본 파일을 찾을 수 없습니다: {source}\n")
            continue

        # 치환 적용
        modified_text = apply_replacements(original_text, replacements)

        # 치환 횟수 계산
        job_replacements = 0
        for repl in replacements:
            from_str = repl.get("from", "")
            if from_str:
                count = original_text.count(from_str)
                job_replacements += count
                with open(log_path, 'a', encoding='utf-8') as lf:
                    lf.write(f"[{datetime.datetime.now()}] {source} -> {dest}: '{from_str}' -> '{repl.get('to', '')}' {count}건 치환\n")

        # 대상 파일 저장
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        with open(dest, 'w', encoding='utf-8') as df:
            df.write(modified_text)

        total_replacements += job_replacements
        summary_data.append(f"{source} -> {dest}: {job_replacements}건 치환")

        with open(log_path, 'a', encoding='utf-8') as lf:
            lf.write(f"[{datetime.datetime.now()}] 작업 {i} 완료: {job_replacements}건 치환\n")

    # 요약 파일 생성
    with open(summary_path, 'w', encoding='utf-8') as sf:
        sf.write("치환 작업 요약:\n")
        for line in summary_data:
            sf.write(line + "\n")
        sf.write(f"\n총 파일 수: {len(jobs)}, 총 치환 횟수: {total_replacements}")

    print(f"치환 작업이 완료되었습니다. (총 {total_replacements}건 치환)")

def main():
    while True:
        print("\n=== 문자열 치환 도구 ===")
        print("1. YAML 생성 (엑셀 -> YAML)")
        print("2. 미리보기 (YAML 기반 diff 출력)")
        print("3. 실행 (치환 적용 및 로그 저장)")
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