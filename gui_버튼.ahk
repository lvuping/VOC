#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


Gui,Show,x50 y50 h500 w500, test

Gui,Add,Button,x20 y20 h100 w100 ,버튼


gui,add,button,x200 y200 h250 w250,내가만든버튼
return

Button버튼:

msgbox,버튼 눌림!

return

button내가만든버튼:

msgbox, 만든거 

