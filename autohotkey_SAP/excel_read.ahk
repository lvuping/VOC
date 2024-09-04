#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.

; Excel 애플리케이션 객체 생성
excel := ComObject("Excel.Application")
excel.Visible := true ; Excel 창을 보이게 설정 (선택사항)

; 새 워크북 생성 또는 기존 파일 열기
workbook := excel.Workbooks.Open("C:\Path\To\Your\Excel\File.xlsx") ; 파일 경로를 적절히 수정하세요
sheet := workbook.Sheets(1) ; 첫 번째 시트 선택

; B2 셀의 값 읽기
valueB2 := sheet.Range("B2").Value

; C3 셀에 값 쓰기
sheet.Range("C3").Value := valueB2

; 변경사항 저장
workbook.Save()

; Excel 종료
excel.Quit()

MsgBox, 작업이 완료되었습니다. B2의 값 "%valueB2%"을(를) C3에 복사했습니다.