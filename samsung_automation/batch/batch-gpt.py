import openpyxl
import pyperclip
import win32com.client
import os
import time
from datetime import datetime
import pyautogui as gui
import json
import logging
from typing import Optional, Tuple

# 로깅 설정
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# 설정 파일 로드
with open("config.json", "r") as config_file:
    CONFIG = json.load(config_file)

SAP_PROGRAM_NAME = CONFIG["SAP_PROGRAM_NAME"]


def process_data(company_code: str, title: str, date_low: str, date_high: str) -> None:
    """SAP에서 데이터를 처리하고 Excel 파일로 저장하는 메인 함수"""
    sap_session = get_sap()
    set_gui_initial_page(sap_session, company_code, date_low, date_high)
    execute_batch()
    monitor_batch_status()
    wait_generating_excel()
    wait_excel_ready()
    save_path = os.path.join(os.getcwd(), title)
    save_andclose_excel(save_path)


def set_title(company_code: str, date1: str, date2: Optional[str] = None) -> str:
    """파일 제목 설정 함수"""
    result = company_code
    date1 = date1.replace(".", "").replace("/", "")

    if date2:
        date2 = date2.replace(".", "")
        result += f"_{date1}_{date2[4:8]}"
    else:
        result += f"_{date1}"
    return result


def process_excel_rows(server: str, start_row: Optional[int] = None) -> int:
    """Excel 파일의 행을 처리하는 함수"""
    current_row = start_row if start_row else 2
    fileName = f"{server}.xlsx"

    workbook = openpyxl.load_workbook(fileName)
    sheet = workbook.active

    while True:
        company_code = sheet[f"A{current_row}"].value
        date_low = sheet[f"B{current_row}"].value
        date_high = sheet[f"C{current_row}"].value
        filename = sheet[f"D{current_row}"].value

        if not company_code:
            break

        logging.info(
            f"Processing Row {current_row}: CODE={company_code}, Date-low={date_low}, Date-high={date_high}, Filename={filename}"
        )

        title = set_title(company_code, date_low, date_high)
        process_data(company_code, title, date_low, date_high)

        current_row += 1

    logging.info(f"Processed rows from {start_row} to {current_row - 1}")
    close_sap()
    return current_row


def copy_excel_ragne(start_row: int, end_row: int) -> None:
    """Excel 범위를 복사하는 함수"""
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
    logging.info("Copied data to clipboard")


def get_sap() -> win32com.client.CDispatch:
    """SAP 세션을 가져오는 함수"""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    sap_session = connection.Children(0)
    logging.info("SAP session established")
    return sap_session


def save_andclose_excel(save_path: str) -> None:
    """Excel 파일을 저장하고 닫는 함수"""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False

    try:
        for i in range(excel.Workbooks.Count):
            wb = excel.Workbooks(i + 1)
            if "TRAN" in wb.Name:
                wb.Activate()
                wb.SaveAs(save_path)
                logging.info(f"File saved as {save_path}")
    except Exception as e:
        logging.error(f"Error occurred: {e}")
    finally:
        excel.Quit()
        logging.info("Excel closed")


def execute_batch() -> None:
    """배치 작업을 실행하는 함수"""
    time.sleep(1)
    gui.press("f8")
    time.sleep(2)
    gui.press("enter")
    time.sleep(2)
    gui.press("enter")


def set_gui_initial_page(
    sap_session: win32com.client.CDispatch,
    company_code: str,
    date_low: str,
    date_high: str,
) -> None:
    """SAP GUI 초기 페이지를 설정하는 함수"""
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
    sap_session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").text = (
        date_high if date_high else ""
    )
    sap_session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").setFocus
    sap_session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").caretPosition = 0
    sap_session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    win = gui.getWindowsWithTitle(SAP_PROGRAM_NAME)[0]
    if not win.isActive:
        win.activate()
    logging.info("GUI - Initial page activated")


def monitor_batch_status() -> None:
    """배치 작업 상태를 모니터링하는 함수"""
    while True:
        try:
            gui.hotkey("ctrl", "f2")
            time.sleep(3)
            img_capture = gui.locateOnScreen("finished.png")
            gui.press("enter")
            logging.info(f"Batch job finished at {datetime.now()}")
            break
        except:
            logging.info(f"Batch job still running at {datetime.now()}")
            gui.press("enter")
            time.sleep(3)
    time.sleep(3)
    gui.hotkey("ctrl", "f1")


def wait_generating_excel() -> None:
    """Excel 생성을 기다리는 함수"""
    while True:
        try:
            win = gui.getWindowsWithTitle("Microsoft Excel")[0]
            if not win.isActive:
                win.activate()
                time.sleep(3)
            else:
                break
        except:
            pass


def wait_excel_ready() -> None:
    """Excel 파일이 준비될 때까지 기다리는 함수"""
    while True:
        fore = gui.getActiveWindow()
        time.sleep(1)
        try:
            if "TRAN" in fore.title:
                logging.info("File opened, ready to save")
                break
            else:
                logging.info(f"Current window: {fore.title}")
        except:
            pass


def close_sap() -> None:
    """SAP 세션을 종료하는 함수"""
    sap_session = get_sap()
    sap_session.findById("wnd[0]").maximize
    sap_session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
    sap_session.findById("wnd[0]").sendVKey(0)
    time.sleep(5)
    logging.info("SAP session closed")


def main() -> None:
    """메인 함수"""
    for server, text in CONFIG["SERVERS"].items():
        pyperclip.copy(text)
        gui.hotkey("ctrl", "alt", "c")
        time.sleep(20)
        process_excel_rows(server)


if __name__ == "__main__":
    main()
