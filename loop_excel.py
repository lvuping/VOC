import openpyxl
import time

def process_excel_rows(start_row):
    # Excel 파일 열기
    workbook = openpyxl.load_workbook('파일명.xlsx')
    
    # 활성화된 시트 선택
    sheet = workbook.active

    current_row = start_row

    while True:
        # A, B, C, D 열의 값을 읽음
        a_value = sheet[f'A{current_row}'].value
        b_value = sheet[f'B{current_row}'].value
        c_value = sheet[f'C{current_row}'].value
        d_value = sheet[f'D{current_row}'].value

        # A열이 비어있으면 종료
        if a_value is None or a_value == "":
            break

        # 여기서 a_value, b_value, c_value, d_value를 사용하여 원하는 작업을 수행
        print(f"Row {current_row}: A={a_value}, B={b_value}, C={c_value}, D={d_value}")
        
        # 예: 이 값들을 사용하여 다른 작업을 수행할 수 있습니다.
        # process_data(a_value, b_value, c_value, d_value)

        current_row += 1

    print(f"{start_row}행부터 {current_row-1}행까지 처리되었습니다.")
    
    return current_row  # 다음 시작 행 번호 반환

# 시작
start_row = 1

while True:
    start_row = process_excel_rows(start_row)

    # 사용자에게 계속할지 묻기
    user_input = input("계속하시겠습니까? (y/n): ").lower()
    if user_input != 'y':
        break

    # 사용자가 데이터를 처리할 시간을 주기 위해 잠시 대기
    print("5초 후 다음 범위를 처리합니다...")
    time.sleep(5)

print("프로그램을 종료합니다.")
