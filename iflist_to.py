import pandas as pd
import openpyxl
import yaml
import os
import datetime
import shutil
from openpyxl.styles import Font, PatternFill, Alignment

# 디버그 모드 설정
DEBUG_MODE = True

def debug_print(*args, **kwargs):
    """디버그 모드일 때만 메시지를 출력하는 함수"""
    if DEBUG_MODE:
        print("[DEBUG]", *args, **kwargs)

def process_file_path(file_path, check_flag):
    """
    파일 경로를 TEST에서 PROD로 변환하고 관련 정보를 수집
    
    Args:
        file_path: 원본 파일 경로
        check_flag: 생성여부 플래그 (1인 경우만 처리)
    
    Returns:
        tuple: (PROD 경로, 파일 존재 여부, 디렉토리 파일 개수)
    """
    # 생성여부가 1이 아니면 처리하지 않음
    if pd.isna(check_flag) or float(check_flag) != 1.0:
        return None, None, None
    
    if pd.isna(file_path) or not isinstance(file_path, str):
        return None, None, None
    
    # TEST → PROD 변환
    prod_path = file_path.replace('_TEST_SOURCE', '_PROD_SOURCE')
    
    # 파일 존재 여부 확인
    file_exists = "1" if os.path.exists(prod_path) else ""
    
    # 디렉토리 파일 개수 확인
    dir_path = os.path.dirname(prod_path)
    if os.path.exists(dir_path):
        try:
            file_count = str(len(os.listdir(dir_path)))
        except:
            file_count = "X"
    else:
        file_count = "X"
    
    return prod_path, file_exists, file_count

def generate_excel_and_yaml(input_excel_path, output_excel_path, output_yaml_path):
    """
    입력 엑셀 파일을 읽어 TEST→PROD 변환 후 결과를 엑셀과 YAML로 출력
    
    Args:
        input_excel_path: 입력 엑셀 파일 경로
        output_excel_path: 출력 엑셀 파일 경로
        output_yaml_path: 출력 YAML 파일 경로
    """
    debug_print(f"입력 엑셀 파일 읽기: {input_excel_path}")
    
    # 엑셀 파일 읽기
    df = pd.read_excel(input_excel_path, engine='openpyxl')
    debug_print(f"총 {len(df)}개 행 읽기 완료")
    
    # 처리할 컬럼 정의
    file_types = [
        ('송신파일경로', '송신파일생성여부'),
        ('수신파일경로', '수신파일생성여부'),
        ('송신스키마파일명', '송신스키마파일생성여부'),
        ('수신스키마파일명', '수신스키마파일생성여부')
    ]
    
    # YAML 데이터 준비
    yaml_data = {'files': []}
    
    # 새로운 데이터프레임 준비
    result_data = []
    
    # 각 행 처리
    for idx, row in df.iterrows():
        row_data = {}
        
        # 송신 관련 데이터 (같은 행)
        send_data = {}
        for file_col, check_col in file_types[:2]:  # 송신파일경로, 송신스키마파일명
            if file_col in df.columns and check_col in df.columns:
                file_path = row.get(file_col)
                check_flag = row.get(check_col)
                
                prod_path, file_exists, file_count = process_file_path(file_path, check_flag)
                
                # 원본 데이터 저장
                send_data[file_col] = file_path
                send_data[check_col] = check_flag
                
                # PROD 데이터 저장
                send_data[f"{file_col}PROD"] = prod_path
                send_data[f"{check_col}PROD"] = file_exists
                
                # DFPROD 데이터 저장
                if file_col == '송신파일경로':
                    send_data["송신DFPROD"] = file_count
                else:  # 송신스키마파일명
                    send_data["송신스키마DFPROD"] = file_count
                
                # YAML 데이터 추가
                if prod_path and file_path:
                    yaml_data['files'].append({
                        'source': file_path,
                        'destination': prod_path
                    })
        
        # 수신 관련 데이터 (같은 행)
        recv_data = {}
        for file_col, check_col in file_types[2:]:  # 수신파일경로, 수신스키마파일명
            if file_col in df.columns and check_col in df.columns:
                file_path = row.get(file_col)
                check_flag = row.get(check_col)
                
                prod_path, file_exists, file_count = process_file_path(file_path, check_flag)
                
                # 원본 데이터 저장
                recv_data[file_col] = file_path
                recv_data[check_col] = check_flag
                
                # PROD 데이터 저장
                recv_data[f"{file_col}PROD"] = prod_path
                recv_data[f"{check_col}PROD"] = file_exists
                
                # DFPROD 데이터 저장
                if file_col == '수신파일경로':
                    recv_data["수신DFPROD"] = file_count
                else:  # 수신스키마파일명
                    recv_data["수신스키마DFPROD"] = file_count
                
                # YAML 데이터 추가
                if prod_path and file_path:
                    yaml_data['files'].append({
                        'source': file_path,
                        'destination': prod_path
                    })
        
        # 행 데이터 병합
        row_data.update(send_data)
        row_data.update(recv_data)
        result_data.append(row_data)
    
    # 결과 데이터프레임 생성
    result_df = pd.DataFrame(result_data)
    
    # 컬럼 순서 정의
    columns_order = []
    # 송신 관련 컬럼
    columns_order.extend(['송신파일경로', '송신파일경로PROD', '송신파일생성여부', '송신파일생성여부PROD', '송신DFPROD'])
    columns_order.extend(['송신스키마파일명', '송신스키마파일명PROD', '송신스키마파일생성여부', '송신스키마파일생성여부PROD', '송신스키마DFPROD'])
    # 수신 관련 컬럼
    columns_order.extend(['수신파일경로', '수신파일경로PROD', '수신파일생성여부', '수신파일생성여부PROD', '수신DFPROD'])
    columns_order.extend(['수신스키마파일명', '수신스키마파일명PROD', '수신스키마파일생성여부', '수신스키마파일생성여부PROD', '수신스키마DFPROD'])
    
    # 존재하는 컬럼만 선택
    existing_columns = [col for col in columns_order if col in result_df.columns]
    result_df = result_df[existing_columns]
    
    # 엑셀 파일 저장
    save_excel_with_style(result_df, output_excel_path)
    
    # YAML 파일 저장
    with open(output_yaml_path, 'w', encoding='utf-8') as yf:
        yaml.dump(yaml_data, yf, allow_unicode=True, sort_keys=False)
    
    print(f"\n엑셀 파일 생성 완료: {output_excel_path}")
    print(f"YAML 파일 생성 완료: {output_yaml_path}")
    print(f"총 {len(yaml_data['files'])}개 파일 매핑 생성")
    
    return len(yaml_data['files'])

def save_excel_with_style(df, excel_path):
    """데이터프레임을 스타일이 적용된 엑셀 파일로 저장"""
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    df.to_excel(writer, index=False)
    
    # 워크시트 가져오기
    worksheet = writer.sheets['Sheet1']
    
    # 폰트 크기를 10으로 설정
    font_10 = Font(size=10)
    
    # 헤더(첫 행) 스타일 설정
    light_blue_fill = PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')
    header_font = Font(size=10, bold=True)
    
    # 모든 셀에 대해 폰트 크기 10 적용
    for row in worksheet.iter_rows():
        for cell in row:
            cell.font = font_10
    
    # 헤더 행에 스타일 적용
    for cell in worksheet[1]:
        cell.fill = light_blue_fill
        cell.font = header_font
    
    # 컬럼 너비 자동 조절
    for column in worksheet.columns:
        max_length = 0
        column_letter = openpyxl.utils.get_column_letter(column[0].column)
        
        # 각 셀의 길이를 확인하여 최대 길이 계산
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        # 컬럼 너비 설정 (최대 길이 + 여유 공간)
        adjusted_width = min(max_length + 2, 50)  # 최대 50으로 제한
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # 파일 저장
    writer.close()

def execute_file_copy(yaml_path, log_path):
    """
    YAML 파일에 정의된 파일 복사 작업을 실행
    
    Args:
        yaml_path: YAML 파일 경로
        log_path: 로그 파일 경로
    """
    try:
        with open(yaml_path, 'r', encoding='utf-8') as yf:
            data = yaml.safe_load(yf)
    except FileNotFoundError:
        print(f"YAML 파일을 찾을 수 없습니다: {yaml_path}")
        return
    
    files = data.get('files', [])
    if not files:
        print("복사할 파일이 없습니다.")
        return
    
    # 로그 파일 초기화
    with open(log_path, 'w', encoding='utf-8') as lf:
        lf.write(f"[{datetime.datetime.now()}] 파일 복사 시작\n")
    
    success_count = 0
    skip_count = 0
    error_count = 0
    
    for file_info in files:
        source = file_info.get('source')
        destination = file_info.get('destination')
        
        if not source or not destination:
            continue
        
        # 원본 파일 존재 확인
        if not os.path.exists(source):
            with open(log_path, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] [ERROR] 원본 파일 없음: {source}\n")
            error_count += 1
            continue
        
        # 대상 파일이 이미 존재하는지 확인
        if os.path.exists(destination):
            with open(log_path, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] [ERROR] 파일이 이미 존재: {destination}\n")
            skip_count += 1
            continue
        
        # 대상 디렉토리 생성
        dest_dir = os.path.dirname(destination)
        try:
            os.makedirs(dest_dir, exist_ok=True)
        except Exception as e:
            with open(log_path, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] [ERROR] 디렉토리 생성 실패: {dest_dir} - {str(e)}\n")
            error_count += 1
            continue
        
        # 파일 복사
        try:
            shutil.copy2(source, destination)
            with open(log_path, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] 복사 성공: {source} → {destination}\n")
            success_count += 1
        except Exception as e:
            with open(log_path, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now()}] [ERROR] 복사 실패: {source} → {destination} - {str(e)}\n")
            error_count += 1
    
    # 최종 결과 로그
    with open(log_path, 'a', encoding='utf-8') as lf:
        lf.write(f"\n[{datetime.datetime.now()}] 파일 복사 완료\n")
        lf.write(f"성공: {success_count}개, 건너뜀: {skip_count}개, 오류: {error_count}개\n")
    
    print(f"\n파일 복사 완료")
    print(f"성공: {success_count}개")
    print(f"건너뜀: {skip_count}개 (이미 존재)")
    print(f"오류: {error_count}개")
    print(f"로그 파일: {log_path}")

def main():
    """메인 함수 - 메뉴 시스템"""
    while True:
        print("\n=== 파일 경로 변환 도구 ===")
        print("1. 엑셀 분석 및 YAML 생성")
        print("2. 파일 복사 실행")
        print("0. 종료")
        
        choice = input("\n원하는 작업을 선택하세요: ").strip()
        
        if choice == "1":
            input_excel = input("입력 엑셀 파일 경로를 입력하세요: ").strip()
            if not input_excel:
                print("파일 경로를 입력해주세요.")
                continue
            
            # 출력 파일 경로 설정
            output_excel = "iflist_to.xlsx"
            output_yaml = "iflist_to.yaml"
            
            try:
                count = generate_excel_and_yaml(input_excel, output_excel, output_yaml)
                print(f"\n작업 완료: {count}개 파일 매핑 생성")
            except Exception as e:
                print(f"오류 발생: {str(e)}")
        
        elif choice == "2":
            yaml_path = input("YAML 파일 경로를 입력하세요 (기본값: iflist_to.yaml): ").strip()
            if not yaml_path:
                yaml_path = "iflist_to.yaml"
            
            # 로그 파일 이름 생성
            log_path = f"file_copy_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            
            try:
                execute_file_copy(yaml_path, log_path)
            except Exception as e:
                print(f"오류 발생: {str(e)}")
        
        elif choice == "0":
            print("프로그램을 종료합니다.")
            break
        
        else:
            print("잘못된 선택입니다. 다시 시도하세요.")

if __name__ == "__main__":
    main()