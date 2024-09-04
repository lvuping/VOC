import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtGui import QFont
import openpyxl
import win32com.client
import pyautogui
import time

class DataExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_file = ""

    def initUI(self):
        self.setWindowTitle('Data Extractor')
        self.setGeometry(100, 100, 350, 300)
        self.setFont(QFont('Malgun Gothic', 10))

        # GUI 요소들 생성
        self.table_label = QLabel('Table Name:', self)
        self.table_label.move(10, 50)
        self.table_edit = QLineEdit(self)
        self.table_edit.move(140, 50)
        self.table_edit.setText('BUT000')

        # 나머지 GUI 요소들도 비슷한 방식으로 생성...

        self.file_button = QPushButton('Excel File', self)
        self.file_button.move(10, 230)
        self.file_button.clicked.connect(self.select_file)

        self.start_button = QPushButton('Start', self)
        self.start_button.move(175, 230)
        self.start_button.clicked.connect(self.start_extraction)

        self.show()

    def select_file(self):
        self.selected_file, _ = QFileDialog.getOpenFileName(self, "Select an excel file", "", "XLSX Files (*.xlsx)")
        if self.selected_file:
            QMessageBox.information(self, "File Read Success", f"파일 읽기 성공\n{self.selected_file}")

    def start_extraction(self):
        if not self.selected_file:
            self.selected_file = r"C:\Users\lvupi\Downloads\bp_no.xlsx"
        self.get_data(self.selected_file)

    def get_data(self, xls_path):
        # SAP 세션 가져오기
        sap_session = self.get_sap()
        sap_session.findById("wnd[0]").maximize()

        # Excel 작업
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        workbook = excel.Workbooks.Open(xls_path)
        sheet = workbook.Sheets(1)

        start_num = 1
        end_num = 10000

        for i in range(int(self.repeat_edit.text())):
            if int(self.index_edit.text()) >= i + 1:
                start_num += 10000
                end_num += 10000
                continue

            # Excel에서 데이터 복사
            pyautogui.hotkey('ctrl', 'g')
            pyautogui.write(f"{self.column_edit.text()}{start_num}")
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'g')
            pyautogui.write(f"{self.column_edit.text()}{end_num}")
            time.sleep(1)
            pyautogui.hotkey('shift', 'enter')
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.1)

            start_num += 10000
            end_num += 10000

            # SAP 작업
            sap_session.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16"
            sap_session.findById("wnd[0]").sendVKey(0)
            sap_session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = self.table_edit.text()
            sap_session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 8
            sap_session.findById("wnd[0]").sendVKey(0)

            # 나머지 SAP 작업...

            # 파일 내보내기
            export_file = f"{self.result_file_edit.text()}_{i+1}.XLS"
            sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = export_file
            sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
            sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # 파일 저장 대기 및 닫기
            time.sleep(5)
            pyautogui.hotkey('alt', 'y')
            time.sleep(45)
            pyautogui.hotkey('alt', 'f4')

    def get_sap(self):
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

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sap_gui_auto = None
            return

        return session

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DataExtractor()
    sys.exit(app.exec_())