import openpyxl
import pyperclip
import win32com.client
import os, time
from datetime import datetime
import pyautogui as gui
import shutil
from pathlib import Path
import subprocess
import os
from dotenv import load_dotenv

load_dotenv()

# 환경 변수에서 값 가져오기, 없으면 기본값 사용
SAP_PROGRAM_NAME = os.getenv('SAP_PROGRAM_NAME', '[CIC] ALL transaction report')
SAP_PATH = os.getenv('SAP_PATH', r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

servers = {
    'CAP': os.getenv('CAP_SERVER_NAME', '[CS PORTAL] 아주 운영'),
    'CEP': os.getenv('CEP_SERVER_NAME', '[CS PORTAL] 구주 운영'),
    'CLP': os.getenv('CLP_SERVER_NAME', '[CS PORTAL] 중남미 운영'),
    'CUP': os.getenv('CUP_SERVER_NAME', '[CS PORTAL] 미주 운영')
}

SAP_USERNAME = os.getenv('SAP_USERNAME', 'daiden.kim')
SAP_PASSWORD = os.getenv('SAP_PASSWORD')


def open_sap():
    subprocess.Popen(SAP_PATH)
    print("SAP 실행 명령 전송됨")
    
    start_time = time.time()
    max_wait_time = 10 
    
    while time.time() - start_time < max_wait_time:
        for win in gui.getAllWindows():
            if "SAP Logon 760" in win.title:
                win.activate()
                print("SAP 객체 활성화")
                return  # SAP가 열리면 함수 종료
        
        time.sleep(1)  # 1초 대기
    
    print("SAP를 열지 못했습니다. 최대 대기 시간 초과.")


def login_sap(session):
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "100"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USERNAME
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASSWORD
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
    session.findById("wnd[0]").sendVKey(0)

def is_logged_in(session):
    try:
        # 로그인 화면의 특정 요소를 찾아봅니다.
        session.findById("wnd[0]/usr/txtRSYST-MANDT")
        return False  # 로그인 화면 요소가 있다면 로그인되지 않은 상태
    except:
        try:
            # 로그인된 상태의 특정 요소를 찾아봅니다 (예: SAP Easy Access 화면의 요소)
            session.findById("wnd[0]/tbar[0]/okcd")
            return True  # SAP Easy Access 화면 요소가 있다면 로그인된 상태
        except:
            return False  # 둘 다 아니라면 상태를 확실히 알 수 없으므로 False 반환

def process_data(company_code, title, date_low, date_high,server_name):
    open_sap()
    sap_session = get_sap(server_name)
    
    if not is_logged_in(sap_session):
        login_sap(sap_session)
    else:
        print("이미 로그인된 상태입니다.")
    
    set_gui_initial_page(sap_session,company_code,date_low,date_high)
    execute_batch(sap_session)
    monitor_batch_status(sap_session)
    wait_generating_excel()
    wait_excel_ready()
    save_path = os.path.join(os.getcwd(), title)
    rename_sap_file(title, ".xls")
    # save_andclose_excel(save_path)


def set_title(company_code, date1, date2=None):
    result = company_code
    date1 = date1.replace(".", "").replace("/","")

    if date2:
        date2 = date2.replace(".","")
        result += f"_{date1}_{date2[4:8]}"
    else:
        result += f"_{date1}"
    result += ".xls"
    return result


def find_start_row(sheet):
    current_row = 2
    while sheet[f'A{current_row}'].value is not None:
        if sheet[f'D{current_row}'].value != 'X':
            return current_row
        current_row += 1
    return current_row

def process_excel_rows(server, server_name, start_row=None):
    fileName = server + '.xlsx'
    workbook = openpyxl.load_workbook(fileName)
    sheet = workbook.active

    if start_row is None:
        current_row = find_start_row(sheet)
    else:
        current_row = start_row

    while True:
        company_code = sheet[f'A{current_row}'].value
        date_low = sheet[f'B{current_row}'].value
        date_high = sheet[f'C{current_row}'].value

        if company_code is None or company_code == "":
            break

        print(f"Row {current_row}: CODE={company_code}, Date-low={date_low}, Date-high={date_high}")

        title = set_title(company_code, date_low, date_high)
        process_data(company_code, title, date_low, date_high, server_name)

        # Mark the previous row as done
        if current_row > 2:
            sheet[f'D{current_row - 1}'] = 'X'

        current_row += 1

    # Mark the last processed row as done
    if current_row > 2:
        sheet[f'D{current_row - 1}'] = 'X'

    # Save the workbook
    workbook.save(fileName)

    print(f"{start_row}행부터 {current_row - 1}행까지 처리되었습니다.")
    close_sap()
    return 

def copy_excel_ragne(start_row, end_row):

    wb = openpyxl.load_workbook('bp.xlsx')
    sheet = wb.active
    values = [str(sheet[f'A{i}'].value) for i in range(start_row, end_row) if sheet[f'A{i}'].value is not None]
    text = '\n'.join(values)
    pyperclip.copy(text)
    time.sleep(3)
    print("[Clipboard]")


def get_sap(server_name=None):
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    if not type(sap_gui_auto) == win32com.client.CDispatch:
        return

    application = sap_gui_auto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        sap_gui_auto = None
        return

    if server_name:
        connection = application.OpenConnection(server, True)
    else:
        connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        sap_gui_auto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        sap_gui_auto = None
        return

    print("[SAP Connection]")
    return session


def save_andclose_excel(save_path):
    max_attempts = 30
    for i in range(max_attempts):
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            break
        except Exception as e:
            print(f"Excel 객체 획득 오류: {i}번째 시도중.. {e}")
            time.sleep(2)
            if i == max_attempts:
                excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    excel.DisplayAlerts = False

    while True:
        try:
            wb = excel.Workbooks(1)
            if "TRAN" in wb.Name:
                wb.Activate()
                wb.SaveAs(save_path)
                print(f"파일이 {save_path}로 저장 되었습니다.")
                excel.Quit()
                print("Excel 종료")
                break

        except Exception as e:
            time.sleep(3)
            print(f"오류 발생: {e}")
            print(f"기존의 excel 객체 활성화 시도")
            excel = win32com.client.GetActiveObject("Excel.Application")
            pass

def rename_sap_file(new_filename, file_extension=".xls"):
    saved_file_name = ""
    for win in gui.getAllWindows():
        if "TRAN" in win.title:
            win.activate()
            saved_file_name = win.title
            gui.hotkey('ctrl', 'f4')
            break

    temp_folder = Path(os.path.expanduser("~")) / "AppData" / "Local" / "SAP" / "SAP GUI" / "tmp"
    newfilename_with_ext = new_filename

    for file in temp_folder.glob("*"):
        file_time = file.stat().st_mtime
        file_name, file_ext = os.path.splitext(file)

        if 'TRAN' in file_name:
            print(file)
            new_path = temp_folder / newfilename_with_ext
            try:
                file.rename(new_path)
                print(f"파일 이름이 성공적으로 변경되었습니다: {new_path}")
            except Exception as e:
                print(f"파일 이름 변경중 오류 발생{new_path}")

# rename_sap_file("test", ".xls")

def execute_batch(sap_session):

    time.sleep(1)
    gui.press('f8')
    time.sleep(2)
    gui.press('enter')
    time.sleep(2)
    gui.press('enter')



def set_gui_initial_page(sap_session,company_code,date_low,date_high):
    sap_session.findById("wnd[0]").maximize
    sap_session.findById("wnd[0]/tbar[0]/okcd").text = "/nZRIC001"
    sap_session.findById("wnd[0]").sendVKey(0)

    sap_session.findById("wnd[0]/usr/txtS_COUNT-LOW").text = ""
    sap_session.findById("wnd[0]/usr/txtS_AGENT-LOW").text = ""
    sap_session.findById("wnd[0]/usr/chkCHK_BACK").setFocus
    sap_session.findById("wnd[0]/usr/chkCHK_BACK").selected = True
    sap_session.findById("wnd[0]/usr/ctxtS_PRTYPE-LOW").text = "ZC11"
    sap_session.findById("wnd[0]/usr/ctxtP_BUKRS").text = company_code
    sap_session.findById("wnd[0]/usr/ctxtS_PDAT-LOW").text = date_low
    if date_high:
        sap_session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").text = date_high
    else:
        sap_session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").text = ""
    sap_session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").setFocus
    sap_session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").caretPosition = 0
    sap_session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    win = gui.getWindowsWithTitle(SAP_PROGRAM_NAME)[0]
    if win.isActive == False:
        win.activate()
    print("GUI - Initial page activated")


def monitor_batch_status(sap_session):
    while True:
        try:
            gui.hotkey('ctrl', 'f2')
            time.sleep(3)
            img_capture = gui.locateOnScreen("finished.png")
            gui.press('enter')
            print(f"Time: {datetime.now()} Found  finished.png!")
            break
        except:

            print(f"Time: {datetime.now()} :Cannot find finished.png")
            gui.press('enter')
            time.sleep(3)
            pass
    time.sleep(3)
    # Get Background data
    gui.hotkey('ctrl', 'f1')


def wait_generating_excel():
    while True:
        try:
            win = gui.getWindowsWithTitle("Microsoft Excel")[0]
            if win.isActive == False:
                win.activate()
                print(f"Time: {datetime.now()} Microsoft Excel Activated")
                time.sleep(3)
                pass
            else:
                gui.press('y')
                print(f"Time: {datetime.now()} TRAN file opened")
                break
        except:
            time.sleep(1)
            pass

def wait_excel_ready():
    while True:
        fore = gui.getActiveWindow()
        time.sleep(1)
        try:
            if "TRAN" in fore.title:
                print("File opened, ready to save")
                break
            else:
                print(fore.title)
        except:
            pass

def close_sap():
    sap_session = get_sap()
    sap_session.findById("wnd[0]").maximize
    sap_session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
    sap_session.findById("wnd[0]").sendVKey(0)
    time.sleep(5)






# save_path = os.path.join(os.getcwd(), title)
# for win in gui.getAllWindows():
#     print(win)


# Main execution
for server, server_name in servers.items():
    process_excel_rows(server, server_name)






# save_path = os.path.join(os.getcwd(), title)
# for win in gui.getAllWindows():
#     print(win)


# for server in servers:
#     if server == 'CAP':
#         server_name = '[CS PORTAL] 아주 운영'
#     elif server == 'CEP':
#         server_name = '[CS PORTAL] 구주 운영'
#     elif server == 'CLP':
#         server_name = '[CS PORTAL] 중남미 운영'
#     elif server == 'CUP':
#         server_name = '[CS PORTAL] 미주 운영'
#     process_excel_rows("CEP")


gui.

    # for win in gui.getAllWindows():
    #     if "SAP Logon 760" in win.title:
    #         win.activate()
    #         print(f"SAP 객체 활성화")
    #     else: