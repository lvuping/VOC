SetWorkingDir %A_ScriptDir%
SendMode Input
#SingleInstance Force
#Include search_image_from_monitor.ahk

; 전역 변수 선언을 한 곳에 모아두어 가독성 향상
global SelectedFile := ""
global excelTitle := ""

Gui, Add, Radio, x10 y10 w130 h20 vFirst Checked, 첫 번째 모니터
Gui, Add, Radio, x10 y40 w130 h20 vSecond, 두 번째 모니터
Gui, Add, Radio, x10 y70 w130 h20 vThird, 세 번째 모니터
Gui, Add, Button, x150 y10 w70 h50 gFile, Excel 파일
Gui, Add, Button, x10 y100 w200 h30 gBtn, 시작
Gui, Show, w230 h150, 자동화
return

Btn:
    Gui, Submit, NoHide
    Get_ZRIC001(SelectedFile)
return

File:
    FileSelectFile, SelectedFile, 3,, Select a CSV file, CSV Files (*.csv)
    if (SelectedFile != "") {
        MsgBox, 파일 읽기 성공`n%SelectedFile%
        Process, Close, Excel.exe
    }
return

GuiClose:
ExitApp

Esc::
ExitApp
return

GetSAP() {
    SetTitleMatchMode, 2
    WinActivate ahk_class SAP_FRONTEND_SESSION
    _oSAP := ComObjGet("SAPGUI").GetScriptingEngine
return _oSAP.ActiveSession
}

Get_ZRIC001(XLS_PATH) {
    ; Excel 객체 생성 및 파일 열기
    excel := ComObjCreate("Excel.Application")
    excel.Visible := false
    workbook := excel.Workbooks.Open(XLS_PATH)
    sheet := workbook.Sheets(1)

    ; 마지막 행 찾기
    lastRow := sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row

    ; SAP 세션 가져오기
    session := GetSAP()

    ; SAP 초기 설정 
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text := "/nZRIC001"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows := "0"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

    ; 데이터 처리 루프
    Loop, % lastRow - 1 ; 헤더를 제외한 데이터 행 수만큼 반복
    {
        cellValue := sheet.Cells(A_Index + 1, 2).Value ; 헤더를 건너뛰기 위해 A_Index + 1
        excelTitle := sheet.Cells(A_Index + 1, 3).Value

        ; SAP 데이터 입력 및 처리 (기존 코드와 유사)
        session.findById("wnd[0]/usr/ctxtS_PDAT-LOW").text := cellValue
        Sleep, 1000

        ; 실행 확인 및 대기
        Loop, 2 {
            SendInput, {Enter}
            Sleep, 1500
        }

        ; 작업 완료 대기 및 결과 처리
        While (true) {
            Click, 306, 132 ; Batch Job Monitoring
            Sleep, 1500
            if (Search_Image_from_monitor(3, "finished.png") == 0) {
                SendInput, {Enter}
                Sleep, 1500
                MouseMove, 0, 0
                Click, 156, 130 ; Get background data

                WinWait, ahk_exe EXCEL.EXE
                Sleep, 5000
                SendInput, Y
                Sleep, 15000

                ; 새로운 Excel 파일 저장 및 종료
                fileName := excelTitle . ".xls"
                newFile := ComObject("Excel.Application")
                try {
                    ; 새로 생성된 파일이 이미 열려있다고 가정
                    newFile.ActiveWorkbook.SaveAs(fileName)
                    newFile.ActiveWorkbook.Close(0)
                } catch e {
                    MsgBox, % "Error saving Excel file: " . e.Message
                } finally {
                    newFile.Quit()
                    ObjRelease(newFile)
                    newFile := ""
                }

                ; Excel 프로세스 종료 (필요한 경우)
                Process, Close, EXCEL.EXE
                break
            } else {
                Sleep, 1500
                SendInput, {Enter}
            }
        }
    }

    ; Excel 종료
    workbook.Close(false)
    excel.Quit()
}