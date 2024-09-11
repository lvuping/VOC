import openpyxl
import pyperclip
import win32com.client
import os, time
from datetime import datetime
import pyautogui as gui

SAP_PROGRAM_NAME = "[CIC] ALL transaction report"


def process_data(company_code, title, date_low, date_high):
    sap_session = get_sap()
    set_gui_initial_page(sap_session, company_code, date_low, date_high)
    execute_batch()
    monitor_batch_status()
    wait_generating_excel()
    wait_excel_ready()
    save_path = os.path.join(os.getcwd(), title)
    save_andclose_excel(save_path)


def set_title(company_code, date1, date2=None):
    result = company_code
    date1 = date1.replace(".", "").replace("/", "")

    if date2:
        date2 = date2.replace(".", "")
        result += f"_{date1}_{date2[4:8]}"
    else:
        result += f"_{date1}"
    return result


def process_excel_rows(server, start_row=None):
    # Excel 파일 열기
    if start_row:
        current_row = start_row
    else:
        current_row = 2

    fileName = server + ".xlsx"
    workbook = openpyxl.load_workbook(fileName)

    # 활성화된 시트 선택
    sheet = workbook.active

    while True:
        # A, B, C, D 열의 값을 읽음
        company_code = sheet[f"A{current_row}"].value
        date_low = sheet[f"B{current_row}"].value
        date_high = sheet[f"C{current_row}"].value
        filename = sheet[f"D{current_row}"].value

        # A열이 비어있으면 종료
        if company_code is None or company_code == "":
            break

        # 여기서 a_value, b_value, c_value, d_value를 사용하여 원하는 작업을 수행
        print(
            f"Row {current_row}: CODE={company_code}, Date-low={date_low}, C={date_high}, D={filename}"
        )

        title = set_title(company_code, date_low, date_high)
        # 예: 이 값들을 사용하여 다른 작업을 수행할 수 있습니다.
        process_data(company_code, title, date_low, date_high)

        current_row += 1

    print(f"{start_row}행부터 {current_row - 1}행까지 처리되었습니다.")
    close_sap()
    return current_row  # 다음 시작 행 번호 반환


def copy_excel_ragne(start_row, end_row):

    wb = openpyxl.load_workbook("bp.xlsx")
    sheet = wb.active
    values = [
        str(sheet[f"A{i}"].value)
        for i in range(start_row, end_row)
        if sheet[f"A{i}"].value is not None
    ]
    text = "\n".join(values)
    pyperclip.copy(text)
    time.sleep(3)
    print("[Clipboard]")


def get_sap():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    if not type(sap_gui_auto) == win32com.client.CDispatch:
        return

    application = sap_gui_auto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        sap_gui_auto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        sap_gui_auto = None
        return

    sap_session = connection.Children(0)
    if not type(sap_session) == win32com.client.CDispatch:
        connection = None
        application = None
        sap_gui_auto = None
        return
    print("[SAP Connection]")
    return sap_session


def save_andclose_excel(save_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False

    try:
        for i in range(excel.Workbooks.Count):
            wb = excel.Workbooks(i + 1)
            if "TRAN" in wb.Name:
                wb.Activate()
                wb.SaveAs(save_path)
                print(f"파일이 {save_path}로 저장 되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

    finally:
        excel.Quit()

        print("Excel 종료")


# save_path = os.path.join(os.getcwd(), "test.xls")
# save_andclose_excel(save_path)


# sap_session = get_sap()


#
# img_capture = gui.locateOnScreen("finished.png", confidence=0.7)
# print(img_capture)
# gui.click(img_capture)


def execute_batch():
    # execute '
    # gui.hotkey('ctrl', 'f1')
    time.sleep(1)
    gui.press("f8")
    time.sleep(2)
    gui.press("enter")
    time.sleep(2)
    gui.press("enter")


def set_gui_initial_page(sap_session, company_code, date_low, date_high):
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


def monitor_batch_status():
    while True:
        try:
            # execute 'Batch job monitoring'
            gui.hotkey("ctrl", "f2")
            time.sleep(3)
            img_capture = gui.locateOnScreen("finished.png")
            gui.press("enter")
            print(f"Time: {datetime.now()} Found  finished.png!")
            break
        except:

            print(f"Time: {datetime.now()} :Cannot find finished.png")
            gui.press("enter")
            time.sleep(3)
            pass
    time.sleep(3)
    # Get Background data
    gui.hotkey("ctrl", "f1")


def wait_generating_excel():
    while True:
        try:
            win = gui.getWindowsWithTitle("Microsoft Excel")[0]
            if win.isActive == False:
                win.activate()
                time.sleep(3)
                pass
            else:
                break
        except:
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


# sap_session = get_sap()
# set_gui_initial_page(sap_session)
# execute_batch()
# monitor_batch_status()
# wait_generating_excel()
# wait_excel_ready()
# save_path = os.path.join(os.getcwd(), title)
# save_andclose_excel(save_path)

# for win in gui.getAllWindows():
#     print(win)

servers = ["CAP", "CEP", "CLP", "CUP"]

for server in servers:
    if server == "CAP":
        text = "C550"
    elif server == "CEP":
        text = "C780"
    elif server == "CLP":
        text = "C820"
    elif server == "CUP":
        text = "C310"

    pyperclip.copy(text)
    gui.hotkey("ctrl", "alt", "c")
    time.sleep(20)
    process_excel_rows(server)
