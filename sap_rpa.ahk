SetWorkingDir %A_ScriptDir%
SendMode Input
#SingleInstance Force
#Include search_image_from_monitor.ahk


global SelectedFile
global excelTitle



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
MsgBox, 파일 읽기 성공`n%SelectedFile%
Process,Close,Excel.exe
return

GuiClose:
ExitApp

Esc::
	ExitApp
	return



GetSAP(){
	SetTitleMatchMode, 2
    WinActivate ahk_class SAP_FRONTEND_SESSION
    _oSAP := ComObjGet("SAPGUI").GetScriptingEngine


    session := _oSAP.ActiveSession
    return session
}

Get_ZRIC001(XLS_PATH)
{
	excel := ComObjCreate("Excel.Application")
	excel.Visible := false ; Excel 창을 보이게 설정 (선택사항)
	workbook := excel.Workbooks.Open(XLS_PATH)
	sheet := workbook.Sheets(1) ; 첫 번째 시트 선택

    ; Get A's Last row
    lastRow := sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row


	session := GetSAP()
	session.findById("wnd[0]").maximize
	session.findById("wnd[0]/tbar[0]/okcd").text := "/nZRIC001"
	session.findById("wnd[0]").sendVKey(0)
	session.findById("wnd[0]/tbar[1]/btn[17]").press
	session.findById("wnd[1]/tbar[0]/btn[8]").press
	session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows := "0"
	session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
	session.findById("wnd[0]/usr/ctxtS_PDAT-LOW").text := "test"

    ; A열 검색 및 C열 값 추출
    ;~ found := false
	LineNum := 0
    Loop, %lastRow%
	{
		isRunning := true
		LineNum++
		if (LineNum == 1)  ; 헤더 줄 건너뛰기
			continue
        cellValue := sheet.Cells(A_Index, 2).Value
        excelTitle := sheet.Cells(A_Index, 3).Value
		session.findById("wnd[0]/usr/ctxtS_PDAT-LOW").text := cellValue
        ;~ if (cellValue = ClipboardContent)
		Sleep,1000

		;~ Sendinput, {F8}

		;START 실행하겠습니까? 혹은 이미 실행중인데, 다시 실행할거냐 메시지에 대한 Yes 2번
		Sendinput, {Enter}
		Sleep, 1500
		Sendinput, {Enter}
		Sleep, 1500
		Sendinput, {Enter}
		;~ END Yes 2번 종료

		isRunning := true
		While(isRunning)
		{
			;~ Send, {Ctrl Down}{F2}{Ctrl Up} ;Batch Job Monitoring
			Click, 306, 132
			Sleep,1500
			Result:=Search_Image_from_monitor(3, "finished.png")
			IF (Result == 0)
			{
				;~ Get Background Data
				isRunning := false
				Sendinput, {Enter}
				Sleep, 1500
				MouseMove, 0, 0
				Click,156, 130 ;click Get background data

				WinWait, ahk_exe EXCEL.EXE
				Sleep, 5000
				SendInput, Y
				Sleep, 15000
				;~ Click, 144, 130 ;download data
				fileName := %excelTitle% . ".xls"
				newFile := ComObjActive("excel.application")
				newFile.activeworkbook.saveas(fileName)
				newFile.Quit()
				ObjRelease(newFile)
			}

			else{
				Sleep, 1500
				Sendinput, {Enter}
			}
		}

    }

    ; Excel 종료
    workbook.Close(false) ; 변경사항 저장하지 않음
    excel.Quit()
	return
}
