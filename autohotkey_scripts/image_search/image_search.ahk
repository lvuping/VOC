;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;스크립트 시작 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#Include Gdip_all.ahk

#Include Gdip_ImageSearch.ahk

#Include Search_img.ahk

global Title

Title=NAVER - Chrome

gui,show,w100 h100 center,chapter11

gui,Add,Button,x0 y0 w100 h50 gStart,Start

gui,Add,Button,x0 y50 w100 h50 gStop,Stop

return

Start:

    WinGet,winid,ID,%Title%

    if(Search_img("test.bmp",winid,x,y)){

        MsgBox, success!! x=%x% y=%y%

        postclick(x,y)

    }

    else

    msgbox,못찾음

return

Stop:

ExitApp

return

GuiClose:

ExitApp

return

PostClick(FoundX,FoundY){

    lparam:=FoundX|FoundY<<16

    PostMessage,0x201,1,%lparam%,,%Title%

    PostMessage,0x202,0,%lparam%,,%Title%

    Sleep, 1000

}

;;;;;;;;;;;;;;;;;;;;;;;;;;;; 스크립트 끝 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;