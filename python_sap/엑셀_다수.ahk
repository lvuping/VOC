#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

global SelectedFile := ""

Gui, Add, Button, x10 y10 w200 h30 gSelectFile, Excel 파일 선택
Gui, Add, Button, x10 y50 w200 h30 gRunTest, 테스트 실행
Gui, Show, w220 h90, Excel 자동화 테스트
return

SelectFile:
    FileSelectFile, SelectedFile, 3,, Select an Excel file, Excel Files (*.xlsx; *.xls)
    if (SelectedFile != "") {
        MsgBox, 파일 선택 성공`n%SelectedFile%
    }
return

RunTest:
    if (SelectedFile == "") {
        MsgBox, 먼저 Excel 파일을 선택해주세요.
        return
    }
    TestExcelAutomation(SelectedFile)
return

GuiClose:
ExitApp

Esc::
ExitApp
return

TestExcelAutomation(filePath) {
    ; Excel 애플리케이션 생성
    excel := ComObject("Excel.Application")
    excel.Visible := false

    try {
        ; 원본 파일 열기
        workbook := excel.Workbooks.Open(filePath)
        sheet := workbook.Sheets(1)

        ; 마지막 행 찾기
        lastRow := sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row

        ; 데이터 읽기 및 처리
        Loop, % lastRow - 1 ; 헤더를 제외한 데이터 행 수만큼 반복
        {
            cellValue := sheet.Cells(A_Index + 1, 2).Value ; B열 값 읽기
            excelTitle := sheet.Cells(A_Index + 1, 3).Value ; C열 값 읽기

            ; 새 워크북 생성 (SAP에서 생성되는 파일을 시뮬레이션)
            newWorkbook := excel.Workbooks.Add()
            newSheet := newWorkbook.Sheets(1)

            ; 샘플 데이터 입력
            newSheet.Cells(1, 1).Value := "테스트 데이터"
            newSheet.Cells(2, 1).Value := cellValue
            newSheet.Cells(3, 1).Value := excelTitle

            ; 파일 저장
            SaveAndCloseExcelFile(excel, newWorkbook, excelTitle)

            ; 잠시 대기
            Sleep, 2000
        }

        MsgBox, 테스트가 완료되었습니다.
    }
    catch e {
        MsgBox, 에러가 발생했습니다: %e%
    }
    finally {
        ; 원본 워크북 닫기
        workbook.Close(false)
        ; Excel 종료
        excel.Quit()
        ; 메모리 해제
        ObjRelease(excel)
    }

    ; 모든 Excel 프로세스 종료
    CloseAllExcelInstances()
}

SaveAndCloseExcelFile(excel, workbook, excelTitle) {
    fileName := A_ScriptDir . "\" . excelTitle . ".xlsx"
    workbook.SaveAs(fileName)
    workbook.Close(false)
}

CloseAllExcelInstances() {
    Process, Close, EXCEL.EXE
    Sleep, 2000 ; Excel 프로세스가 완전히 종료될 때까지 대기
}