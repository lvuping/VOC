; AutoHotkey 스크립트

; VBScript 코드를 문자열로 정의
vbscode := "
    ' VBScript 코드
    Set SapGui = GetObject(""SAPGUI"")
    Set Appl = SapGui.GetScriptingEngine
    Set Connection = Appl.OpenConnection(""Your SAP System Description"", True)
    Set Session = Connection.Children(0)

    ' 요소가 준비될 때까지 대기하는 함수
    Function WaitForElement(id, maxWaitTime)
        Dim startTime, element
        startTime = Timer()
        Do
            On Error Resume Next
            Set element = Session.FindById(id)
            On Error GoTo 0
            If Not element Is Nothing Then
                WaitForElement = True
                Exit Function
            End If
            WScript.Sleep 100 ' 100ms 대기
        Loop Until Timer() - startTime > maxWaitTime
        WaitForElement = False
    End Function

    ' 로그인 화면 대기
    If Not WaitForElement(""wnd[0]/usr/txtRSYST-BNAME"", 10) Then
        MsgBox ""로그인 화면이 나타나지 않았습니다.""
        WScript.Quit
    End If

    ' SAP 로그인
    Session.FindById(""wnd[0]/usr/txtRSYST-BNAME"").Text = ""YourUsername""
    Session.FindById(""wnd[0]/usr/pwdRSYST-BCODE"").Text = ""YourPassword""
    Session.FindById(""wnd[0]/tbar[0]/btn[0]"").press

    ' 메인 화면 대기
    If Not WaitForElement(""wnd[0]/tbar[0]/okcd"", 20) Then
        MsgBox ""메인 화면이 나타나지 않았습니다.""
        WScript.Quit
    End If

    ' 트랜잭션 실행 (예: VA01)
    Session.FindById(""wnd[0]/tbar[0]/okcd"").Text = ""VA01""
    Session.FindById(""wnd[0]"").SendVKey 0

    ' VA01 화면 대기 (예: 판매문서 유형 필드)
    If Not WaitForElement(""wnd[0]/usr/ctxtVBAK-AUART"", 15) Then
        MsgBox ""VA01 화면이 나타나지 않았습니다.""
        WScript.Quit
    End If

    ' 여기에 추가 작업을 넣을 수 있습니다
    ' 각 작업 후에 적절한 WaitForElement 호출을 추가하세요

    ' 세션 종료
    Session.FindById(""wnd[0]/tbar[0]/btn[15]"").press
"

; 임시 VBS 파일 생성
FileSelectFile, vbsFile, S16, %A_Temp%\sap_automation.vbs, Create VBS file, VBS Files (*.vbs)
if (vbsFile = "")
    return

FileAppend, %vbscode%, %vbsFile%

; VBScript 실행
RunWait, wscript.exe "%vbsFile%"

; 임시 파일 삭제
FileDelete, %vbsFile%

MsgBox, SAP 자동화가 완료되었습니다.
