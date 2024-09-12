import openpyxl
import pyperclip
import win32com.client
import os
import time
from datetime import datetime
import pyautogui as gui
import json
import logging
from typing import Optional, Dict

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class Config:
    def __init__(self, config_file: str = 'config.json'):
        with open(config_file, 'r') as file:
            self.config = json.load(file)
        self.SAP_PROGRAM_NAME = self.config['SAP_PROGRAM_NAME']
        self.SERVERS = self.config['SERVERS']

class SAPSession:
    def __init__(self):
        self.session = self.get_sap()

    def get_sap(self) -> win32com.client.CDispatch:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        connection = application.Children(0)
        sap_session = connection.Children(0)
        logging.info("SAP session established")
        return sap_session

    def set_gui_initial_page(self, company_code: str, date_low: str, date_high: str) -> None:
        self.session.findById("wnd[0]").maximize
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nZRIC001"
        self.session.findById("wnd[0]").sendVKey(0)

        self.session.findById("wnd[0]/usr/txtS_COUNT-LOW").text = ""
        self.session.findById("wnd[0]/usr/txtS_AGENT-LOW").text = ""
        self.session.findById("wnd[0]/usr/chkCHK_BACK").setFocus
        self.session.findById("wnd[0]/usr/chkCHK_BACK").selected = True
        self.session.findById("wnd[0]/usr/ctxtS_PRTYPE-LOW").text = "ZC11"
        self.session.findById("wnd[0]/usr/ctxtP_BUKRS").text = company_code
        self.session.findById("wnd[0]/usr/ctxtS_PDAT-LOW").text = date_low
        self.session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").text = date_high if date_high else ""
        self.session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").setFocus
        self.session.findById("wnd[0]/usr/ctxtS_PDAT-HIGH").caretPosition = 0
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        win = gui.getWindowsWithTitle(Config().SAP_PROGRAM_NAME)[0]
        if not win.isActive:
            win.activate()
        logging.info("GUI - Initial page activated")

    def close(self) -> None:
        self.session.findById("wnd[0]").maximize
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(5)
        logging.info("SAP session closed")

class ExcelProcessor:
    def __init__(self, filename: str):
        self.workbook = openpyxl.load_workbook(filename)
        self.sheet = self.workbook.active

    def process_rows(self, start_row: int = 2) -> int:
        current_row = start_row
        while True:
            company_code = self.sheet[f'A{current_row}'].value
            date_low = self.sheet[f'B{current_row}'].value
            date_high = self.sheet[f'C{current_row}'].value
            filename = self.sheet[f'D{current_row}'].value

            if not company_code:
                break

            logging.info(f"Processing Row {current_row}: CODE={company_code}, Date-low={date_low}, Date-high={date_high}, Filename={filename}")

            title = self.set_title(company_code, date_low, date_high)
            DataProcessor.process_data(company_code, title, date_low, date_high)

            current_row += 1

        logging.info(f"Processed rows from {start_row} to {current_row - 1}")
        return current_row

    @staticmethod
    def set_title(company_code: str, date1: str, date2: Optional[str] = None) -> str:
        result = company_code
        date1 = date1.replace(".", "").replace("/","")

        if date2:
            date2 = date2.replace(".","")
            result += f"_{date1}_{date2[4:8]}"
        else:
            result += f"_{date1}"
        return result

    @staticmethod
    def save_andclose_excel(save_path: str) -> None:
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

class GUIAutomation:
    @staticmethod
    def execute_batch() -> None:
        time.sleep(1)
        gui.press('f8')
        time.sleep(2)
        gui.press('enter')
        time.sleep(2)
        gui.press('enter')

    @staticmethod
    def monitor_batch_status() -> None:
        while True:
            try:
                gui.hotkey('ctrl', 'f2')
                time.sleep(3)
                gui.locateOnScreen("finished.png")
                gui.press('enter')
                logging.info(f"Batch job finished at {datetime.now()}")
                break
            except:
                logging.info(f"Batch job still running at {datetime.now()}")
                gui.press('enter')
                time.sleep(3)
        time.sleep(3)
        gui.hotkey('ctrl', 'f1')

    @staticmethod
    def wait_generating_excel() -> None:
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

    @staticmethod
    def wait_excel_ready() -> None:
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

class DataProcessor:
    @staticmethod
    def process_data(company_code: str, title: str, date_low: str, date_high: str) -> None:
        sap_session = SAPSession()
        sap_session.set_gui_initial_page(company_code, date_low, date_high)
        GUIAutomation.execute_batch()
        GUIAutomation.monitor_batch_status()
        GUIAutomation.wait_generating_excel()
        GUIAutomation.wait_excel_ready()
        save_path = os.path.join(os.getcwd(), title)
        ExcelProcessor.save_andclose_excel(save_path)
        sap_session.close()

class Main:
    def __init__(self):
        self.config = Config()

    def run(self) -> None:
        for server, text in self.config.SERVERS.items():
            pyperclip.copy(text)
            gui.hotkey('ctrl', 'alt', 'c')
            time.sleep(20)
            excel_processor = ExcelProcessor(f"{server}.xlsx")
            excel_processor.process_rows()

if __name__ == "__main__":
    Main().run()
