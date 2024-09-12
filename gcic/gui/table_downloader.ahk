SetWorkingDir %A_ScriptDir%
SendMode Input
#SingleInstance Force
FileEncoding, UTF-8
Gui, Font, s10, Malgun Gothic

global SelectedFile := ""
Gui, Add, Text, x10 y50 w100 h20, Table Name:
Gui, Add, Edit, x140 y50 w170 h20 vTable, BUT000

Gui, Add, Text, x10 y80 w100 h20, Key Column
Gui, Add, Edit, x140 y80 w170 h20 vColumn, A

Gui, Add, Text, x10 y110 w100 h20, Query Amount per Time
Gui, Add, Edit, x140 y110 w170 h20 vCount, 10000

Gui, Add, Text, x10 y140 w100 h20, Repeat Count
Gui, Add, Edit, x140 y140 w170 h20 vRepeat, 50

Gui, Add, Text, x10 y170 w100 h20, Resume from Middle
Gui, Add, Edit, x140 y170 w170 h20 vindex,

Gui, Add, Text, x10 y200 w100 h20, File Name
Gui, Add, Edit, x140 y200 w170 h20 vResultFile, BUT000

Gui, Add, Button, x10 y230 w140 h50 gFile, Excel File
Gui, Add, Button, x175 y230 w140 h50 gBtn, Start
Gui, Show, w350 h300, Data Extractor
return

GuiClose:
ExitApp
return
Esc::ExitApp

File:
    FileSelectFile, SelectedFile, 3,, Select a excel file, XLSX Files (*.xlsx)
    if (SelectedFile != "") {
        MsgBox, 파일 읽기 성공`n%SelectedFile%
    }
return

Btn:
    Gui, Submit, NoHide
    if (SelectedFile = "") {
        SelectedFile := "C:\Users\lvupi\Downloads\bp_no.xlsx"
    }
    getData(SelectedFile)
return

getData(XLS_PATH) {
    Gui, Submit, NoHide
    global index
    global Table
    global Count
    global Column
    global Repeat
    global ResultFile

    startNum := 1
    endNum := 10000

    excel := ComObjCreate("Excel.Application")
    excel.Visible := true
    workbook := excel.Workbooks.Open(XLS_PATH)
    sheet := workbook.Sheets(1)
    session := GetSAP()
    session.findById("wnd[0]").maximize
    SplitPath, XLS_PATH, fileName
    Loop, %Repeat%
    {

        if (index >= A_Index)
        {
            startNum += 10000
            endNum += 10000
            continue
        }

        WinWait, %fileName%
        WinActivate, %fileName%

        Clipboard := ""
        SendInput, ^g
        SendInput, %Column%%startNum%
        Sleep,1000
        sendinput, {Enter}
        Sleep,1000
        SendInput, ^g
        SendInput, %Column%%endNum%
        Sleep,1000
        SendInput, {Shift Down}
        sendinput, {Enter}
        SendInput, {Shift Up}
        sleep,100
        SendInput, {Ctrl down}c{Ctrl up}
        ClipWait, , 5
        startNum += 10000
        endNum += 10000
        session.findById("wnd[0]/tbar[0]/okcd").text := "/nzse16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text := Table
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition := "8"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press
        sleep,1500
        session.findById("wnd[1]/tbar[0]/btn[24]").press
        sleep,1500
        session.findById("wnd[1]/tbar[0]/btn[8]").press

        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/tbar[1]/btn[43]").press
        session.findById("wnd[1]/usr/radRB_1").setFocus
        session.findById("wnd[1]/usr/radRB_1").select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        exportFile := ResultFile . "_" . A_Index . ".XLS"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := exportFile
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        ;~ MsgBox, wait %ResultFile%
        WinWait, Microsoft Excel

        sleep,5000
        ;~ MsgBox, activate %ResultFile%
        WinActivate, Microsoft Excel
        SendInput, {Alt Down}y{Alt Up}
        Sleep, 45000
        WinWait, BUT000_
        WinActivate, BUT000_
        SendInput, {Alt Down}{F4}{Alt Up}
        ;~ newWorkbook := excel.ActiveWorkbook
        ;~ newWorkbook.Close(false)

    }
return
}

GetSAP() {
    SetTitleMatchMode, 2
    WinActivate ahk_class SAP_FRONTEND_SESSION
    _oSAP := ComObjGet("SAPGUI").GetScriptingEngine
return _oSAP.ActiveSession
}