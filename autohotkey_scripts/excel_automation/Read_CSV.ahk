#Warn  ; 경고 활성화
SetWorkingDir %A_ScriptDir%  ; 스크립트 디렉토리를 작업 디렉토리로 설정

; CSV 파일 선택
FileSelectFile, SelectedFile, 3,, Select a CSV file, CSV Files (*.csv)
if (SelectedFile = "")
    return

; 출력 파일 이름 생성
SplitPath, SelectedFile,, OutDir, OutExtension, OutNameNoExt
OutputFile := OutDir "\" OutNameNoExt "_modified." OutExtension

; 파일 읽기
FileRead, FileContent, %SelectedFile%
if ErrorLevel
{
    MsgBox, 파일을 읽는 데 실패했습니다.
    return
}

; 새 파일에 쓰기 시작 (헤더 수정)
FileAppend, % "Company code(C310),Date,Naming rule,Status`n", %OutputFile%

; 줄 단위로 처리
LineNum := 0
Loop, Parse, FileContent, `n, `r
{
    LineNum++
    if (LineNum == 1)  ; 헤더 줄 건너뛰기
        continue

    ; 빈 줄 체크
    if (A_LoopField = "")
        break

    ; CSV 줄 파싱
    Fields := StrSplit(A_LoopField, ",")

    ; 필요한 데이터 처리 (예: 두 번째 필드 출력)
    if (Fields.Length() >= 2)
    {
        SecondValue := Fields[2]
        MsgBox, 줄 번호: %LineNum%`n두 번째 값: %SecondValue%

        ; 여기에 추가 처리 로직 구현
        ; 예: if (SecondValue = "특정값") { ... }
    }

    ; 'X' 추가 및 새 파일에 쓰기
    NewLine := A_LoopField . ",X"
    FileAppend, %NewLine%`n, %OutputFile%
}

MsgBox, CSV 파일 처리가 완료되었습니다. 수정된 파일: %OutputFile%