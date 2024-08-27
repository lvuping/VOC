SetWorkingDir %A_ScriptDir%
SendMode Input
#SingleInstance Force
#Include search_image_from_monitor.ahk

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
        CloseAllExcelInstances()
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
    excel := ComObject("Excel.Application")
    excel.Visible := false

    try {
        workbook := excel.Workbooks.Open(XLS_PATH)
        sheet := workbook.Sheets(1)

        lastRow := sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row

        session := GetSAP()

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text := "/nZRIC001"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows := "0"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

        Loop, % lastRow - 1
        {
            cellValue := sheet.Cells(A_Index + 1, 2).Value
            excelTitle := sheet.Cells(A_Index + 1, 3).Value

            session.findById("wnd[0]/usr/ctxtS_PDAT-LOW").text := cellValue
            Sleep, 2000

            Loop, 2 {
                SendInput, {Enter}
                Sleep, 2000
            }

            While (true) {
                Click, 306, 132 ; Batch Job Monitoring
                Sleep, 2000
                if (Search_Image_from_monitor(3, "finished.png") == 0) {
                    SendInput, {Enter}
                    Sleep, 2000
                    MouseMove, 0, 0
                    Click, 156, 130 ; Get background data

                    WinWait, ahk_exe EXCEL.EXE
                    Sleep, 10000
                    SendInput, Y
                    Sleep, 20000

                    SaveAndCloseExcelFile(excel, excelTitle)
                    break
                } else {
                    Sleep, 2000
                    SendInput, {Enter}
                }
            }

            CloseAllExcelInstances()
        }

        MsgBox, 프로세스가 완료되었습니다.
    }
    catch e {
        MsgBox, 에러가 발생했습니다: %e%
    }
    finally {
        workbook.Close(false)
        excel.Quit()
        ObjRelease(excel)
    }
}

SaveAndCloseExcelFile(excel, excelTitle) {
    try {
        newWorkbook := excel.ActiveWorkbook
        fileName := A_ScriptDir . "\" . excelTitle . ".xlsx"
        newWorkbook.SaveAs(fileName)
        newWorkbook.Close(false)
    }
    catch e {
        MsgBox, Excel 파일 저장 중 오류 발생: %e%
    }
}

CloseAllExcelInstances() {
    Process, Close, EXCEL.EXE
    Sleep, 2000 ; Excel 프로세스가 완전히 종료될 때까지 대기
}