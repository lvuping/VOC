#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


Run,%comspec% /c ipconfig /all > ip정보.txt



/* 해설

 

이 스크립틀르 실행함으로써 a.txt 에  ipconfig 실행결과가 저장이 될것입니다.

%comspec% 는 cmd 를 의미한다고 생각하시면 되고

/c 는 실행이 끝나면 꺼지는것이고

/k 는 안꺼지게 하는것입니다.

 

그리고 cmd를 숨기고 싶으면

Run,%comspec% /c ipconfig > a.txt,,hide

하시면 됩니다.

 */