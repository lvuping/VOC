#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

F3::Pause
F4::
Loop{
 Imagesearch,vx,vy, 0,0, 1440,900,*50 전투시작.png
IF ErrorLevel = 0
{
 vx:=vx+20
 vy:=vy+20
mouseclick, left, 284, 769
sleep 1000

}
Imagesearch,vx,vy, 0,0, 1440,900,*50 종료.png
IF ErrorLevel = 0
{
 vx:=vx+20
 vy:=vy+20
mouseclick, left, 76,144
sleep 1000

}
Imagesearch,vx,vy, 0,0, 1440,900,*50 용병선택.png
IF ErrorLevel = 0
{
 vx:=vx+20
 vy:=vy+20
mouseclick, left, 306, 307
sleep 1000

}
Imagesearch,vx,vy, 0,0, 1440,900,*50 고.png
IF ErrorLevel = 0
{
 vx:=vx+20
 vy:=vy+20
mouseclick, left, 427, 560
sleep 1000


}
Imagesearch,vx,vy, 0,0,620, 636,*50 팔로우안함.png
IF ErrorLevel = 0
{
 vx:=vx+20
 vy:=vy+20
mouseclick, left, 427, 560
sleep 1000


}


}

;팔로우안함 Relative:	620, 636 (default)
;넥스트       Relative:	472, 779 (default)
;전투시작    Relative:	284, 769 (default)
;다시하기    Relative:	76, 144 (default)
;용병선택                          947, -179
;고              Relative:	427, 560