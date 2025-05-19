import openpyxl
import yaml
import difflib
import os
import datetime

def generate_yaml_from_excel(excel_path, yaml_path):
    """엑셀 파일을 읽어 YAML 파일을 생성한다."""
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active  # 첫 번째 시트를 사용
    jobs_dict = {}  # (source, dest)를 키로 치환 리스트를 모음

    for row in sheet.iter_rows(min_row=2, values_only=True):  # 2행부터 데이터 읽기 (1행은 헤더)
        source_file, destination_file, from_str, to_str = row
        if source_file is None or destination_file is None or from_str is None or to_str is None:
            continue  # 빈 행은 건너뜀
        key = (str(source_file).strip(), str(destination_file).strip())
        if key not in jobs_dict:
            jobs_dict[key] = []
        jobs_dict[key].append({"from": str(from_str), "to": str(to_str)})

    # jobs_dict를 YAML용 구조로 변환
    jobs = []
    for (source, dest), replacements in jobs_dict.items():
        jobs.append({
            "source": source,
            "destination": dest,
            "replacements": replacements
        })

    # YAML 파일로 저장
    with open(yaml_path, 'w', encoding='utf-8') as yf:
        yaml.safe_dump({"jobs": jobs}, yf, allow_unicode=True)
    return len(jobs)

def apply_replacements(text, replacements):
    """여러 치환 규칙을 순차적으로 적용하여 새로운 텍스트 반환."""
    new_text = text
    for repl in replacements:
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
    """YAML 파일에 정의된 작업들을 실제로 실행하고 로그를 기록."""
    try:
        with open(yaml_path, 'r', encoding='utf-8') as yf:
            data = yaml.safe_load(yf)
    except FileNotFoundError:
        print(f"YAML 파일을 찾을 수 없습니다: {yaml_path}")
        return

    jobs = data.get("jobs", []) if data else []
    total_files = len(jobs)
    total_replacements = 0

    # 로그 파일 초기화
    with open(log_path, 'w', encoding='utf-8') as lf:
        lf.write(f"[{datetime.datetime.now()}] Started replacement jobs (total {total_files} jobs)\n")

    # 요약 파일 초기화
    with open(summary_path, 'w', encoding='utf-8') as sf:
        sf.write("치환 작업 요약:\n")

    for job_idx, job in enumerate(jobs, 1):
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

        # 대상 파일 저장
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        with open(dest, 'w', encoding='utf-8') as df:
            df.write(modified_text)

        # 치환 횟수 계산 및 로그 기록
        job_replacements = 0
        with open(log_path, 'a', encoding='utf-8') as lf:
            lf.write(f"[{datetime.datetime.now()}] Job{job_idx}: {source} -> {dest}\n")
            for repl in replacements:
                from_str = repl.get("from", "")
                to_str = repl.get("to", "")
                if from_str:
                    count = original_text.count(from_str)
                    job_replacements += count
                    lf.write(f"[{datetime.datetime.now()}] \"{from_str}\" -> \"{to_str}\": {count} occurrences replaced\n")
            lf.write(f"[{datetime.datetime.now()}] Job{job_idx} completed: {job_replacements} replacements in total.\n")

        # 요약 파일에 기록
        with open(summary_path, 'a', encoding='utf-8') as sf:
            sf.write(f"{source} -> {dest}: {job_replacements}건 치환 ({len(replacements)}종류 문자열)\n")

        total_replacements += job_replacements

    # 최종 요약 기록
    with open(log_path, 'a', encoding='utf-8') as lf:
        lf.write(f"[{datetime.datetime.now()}] All jobs completed. Total files: {total_files}, Total replacements: {total_replacements}.\n")
    with open(summary_path, 'a', encoding='utf-8') as sf:
        sf.write(f"\n총 파일 수: {total_files}, 총 치환 횟수: {total_replacements}\n")

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
                job_count = generate_yaml_from_excel(excel_path, yaml_path)
                print(f"YAML 파일 생성 완료 (총 {job_count}개 작업)")
            except Exception as e:
                print(f"오류 발생: {str(e)}")
        
        elif choice == "2":
            yaml_path = input("YAML 파일 경로를 입력하세요: ").strip()
            preview_diff(yaml_path)
        
        elif choice == "3":
            yaml_path = input("YAML 파일 경로를 입력하세요: ").strip()
            log_path = input("로그 파일 경로를 입력하세요: ").strip()
            summary_path = input("요약 파일 경로를 입력하세요: ").strip()
            try:
                execute_replacements(yaml_path, log_path, summary_path)
                print("치환 작업이 완료되었습니다.")
            except Exception as e:
                print(f"오류 발생: {str(e)}")
        
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
        
        else:
            print("잘못된 선택입니다. 다시 선택해주세요.")

if __name__ == "__main__":
    main() 