#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#InstallKeybdHook
 
Capslock::
Send {LControl Down}
KeyWait, CapsLock
Send {LControl Up}
if ( A_PriorKey = "CapsLock" )
{
    Send {Backspace}
}
return
 
 
 
; Shift + Ctrl 단축키 위한 스크립트
+CapsLock::+Ctrl
Return
 
 
; Ctrl + CapsLock 토글 방지
^CapsLock::Ctrl
Return
 
 
; Always on Top
^SPACE:: Winset, Alwaysontop, , A ; ctrl + space
Return
 
 
 
; \키를 backspace로
\::Backspace 
Return
 
 
 ; 윈도우키 + \ 키로 \ 입력
#\::\
Return
 
 
; 마우스 클릭
#LAlt::
MouseClick              
return
 
; 마우스 우클릭
#space::MouseClick, right
 
return 
