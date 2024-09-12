; 파일에서 데이터를 읽고 StrSplit()을 사용하여 파싱하는 AutoHotkey 스크립트
#Warn  ; 경고 활성화로 디버깅 용이
SetWorkingDir %A_ScriptDir%  ; 스크립트 디렉토리를 작업 디렉토리로 설정

FileSelectFile, SelectedFile, 3,, Select a text file, Text Files (*.txt)
if (SelectedFile = "")
    return

FileRead, FileContent, %SelectedFile%
if ErrorLevel
{
    MsgBox, 파일을 읽는 데 실패했습니다.
    return
}

Loop, Parse, FileContent, `n, `r  ; 각 줄을 순회
{
    Fields := StrSplit(A_LoopField, ",")  ; 콤마로 필드 분리

    if (Fields.Length() >= 2)  ; 최소 2개의 필드가 있는지 확인
    {
        SecondValue := Fields[2]
        MsgBox, 줄 번호: %A_Index%`n두 번째 값: %SecondValue%

        ; 예: 특정 날짜에 대한 조건 확인
        if (SecondValue = "2024.07.05")
        {
            MsgBox, 특정 날짜를 찾았습니다!
        }
    }
}