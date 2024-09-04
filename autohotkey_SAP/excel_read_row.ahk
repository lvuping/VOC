#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.

; 핫키 설정 (Ctrl+Shift+F로 스크립트 실행)
^+F::
    ; 클립보드 내용 가져오기
    ClipboardContent := Clipboard

    ; Excel 애플리케이션 객체 생성
    excel := ComObject("Excel.Application")
    excel.Visible := true ; Excel 창을 보이게 설정 (선택사항)

    ; 워크북 열기 (파일 경로를 적절히 수정하세요)
    workbook := excel.Workbooks.Open("C:\Path\To\Your\Excel\File.xlsx")
    sheet := workbook.Sheets(1) ; 첫 번째 시트 선택

    ; A열의 사용된 마지막 행 찾기
    lastRow := sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row

    ; A열 검색 및 C열 값 추출
    found := false
    Loop, %lastRow%
    {
        cellValue := sheet.Cells(A_Index, 1).Value
        if (cellValue = ClipboardContent)
        {
            cColumnValue := sheet.Cells(A_Index, 3).Value
            MsgBox, 찾았습니다!`n`nA열 값: %cellValue%`nC열 값: %cColumnValue%
            found := true
            break
        }
    }

    ; 결과가 없을 경우 메시지 표시
    if (!found)
        MsgBox, 클립보드의 값과 일치하는 항목을 A열에서 찾을 수 없습니다.

    ; Excel 종료
    workbook.Close(false) ; 변경사항 저장하지 않음
    excel.Quit()

return