^!t::  ; Ctrl+Alt+T 단축키로 실행
    ; 클립보드 내용을 변수에 저장
    ClipboardOld := ClipboardAll
    
    ; 클립보드를 텍스트로 변환
    Clipboard := Clipboard
    
    ; 앞뒤 공백 제거 (내부 공백은 유지)
    Clipboard := Trim(Clipboard)
    
    ; 변경된 내용을 클립보드에 다시 저장
    ClipboardNew := ClipboardAll
    
    ; 원래 클립보드 내용 복원
    Clipboard := ClipboardOld
    
    ; 새로운 내용을 클립보드에 복사
    Clipboard := ClipboardNew
    
    ; 작업 완료 메시지 표시 (예시 포함)
    MsgBox, 클립보드의 앞뒤 공백이 제거되었습니다.`n`n예시:`n"      Donghyeon Kim             " ->`n"%Clipboard%"
return
