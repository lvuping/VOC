#Warn
SetWorkingDir %A_ScriptDir%
SendMode Input
#SingleInstance Force

global isRunning := true

;~ Esc::
;~ isRunning := false
;~ MsgBox, 중지
;~ return

While_Search_Image_from_monitor(MonitorNo, ImagePath){

    isRunning := true
    SysGet, MonitorCount, MonitorCount
    SysGet, Monitor, Monitor, %MonitorNo%
    ;~ MsgBox, 4, 모니터 선택 확인, %MonitorLeft%  %MonitorTop% %MonitorRight% %MonitorBottom%
    ;~ MsgBox, %MonitorNo%
    while(isRunning){
        ImageSearch,FoundX, FoundY, %MonitorLeft%, %MonitorTop% , %MonitorRight%, %MonitorBottom%, *100 %ImagePath%

        if (ErrorLevel = 0){

            ClickX := FoundX + 20
            ClickY := FoundY + 20
            Click, %ClickX%, %ClickY%
            MsgBox, Found. X: %FoundX% | Y: %FoundY%
            isRunning := false
        }
        else if (ErrorLevel = 1){
            ;~ 화면에서 이미지 못찾음. 3초후 재시도
            Sleep, 3000
        }

        else if (ErrorLevel = 2){
            Sleep, 3000
            ;~ MsgBox, 이미지 경로가 잘못됨
        }

    }
    return

}

Search_Image_from_monitor(MonitorNo, ImagePath){

    SysGet, MonitorCount, MonitorCount
    SysGet, Monitor, Monitor, %MonitorNo%
    ImageSearch,FoundX, FoundY, %MonitorLeft%, %MonitorTop% , %MonitorRight%, %MonitorBottom%, *100 %ImagePath%
    ;~ if (ErrorLevel = 0)
    ;~ {
    ;~ ClickX := FoundX + 20
    ;~ ClickY := FoundY + 20
    ;~ Click, %ClickX%, %ClickY%
    ;~ MsgBox, Found.  X: %FoundX%  |  Y: %FoundY%
    ;~ }
    return ErrorLevel

}

;~ F2::While_Search_Image_from_monitor(2, "finished.png")
;~ F3::
;~ Result:=Search_Image_from_monitor(3, "finished.png")
;~ MsgBox, %Result%