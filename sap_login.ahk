SetTitleMatchMode, 2
WinActivate, SAP Logon 760
WinWaitActive, SAP Logon 760
;>>>>>>>>>>>>>>>>>>>>>>>>>>
_oSAP := ComObjGet("SAPGUI").GetScriptingEngine  ; Get the Already Running Instance
_oSAP.OpenConnection("server", 1)

Session := _oSAP.ActiveSession
;>>>>>>>>>>>>>>>>>>>>>>>>>>

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text := "/nse37"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtRS38L-NAME").text := "STRUCTURE_BUILD"
session.findById("wnd[0]/usr/ctxtRS38L-NAME").caretPosition := 15
session.findById("wnd[0]").sendVKey(7)
Sleep, 1000
