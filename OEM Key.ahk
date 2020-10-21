#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;GUI - OEM Key ( Export, activate and show )
Gui, Add, Groupbox, x10 y10 w400 h50, OEM key
Gui, Add, Edit, ReadOnly x15 y30 w300 h20 vOEMkey
Gui, Add, Button, gExportOEMkey x316 y30 w90 h20, Export
Gui, Add, Button, gTest x136 y80 w90 h20, Test
Gui, Add, Button, gSearchKey x226 y80 w90 h20, Search
Gui, Add, Button, gActivateOEM x316 y80 w90 h20, Activate OEM
Gui, Show, w420 h120, OEM Key
Return

ExportOEMkey:
return

Test:
GetCurrKey := ComObjCreate("WScript.Shell").Exec("wmic path softwarelicensingservice get OA3xOriginalProductKey").StdOut.Readall()
CurrKey := SubStr(GetCurrKey, 35, 29)
MsgBox,, Key, %CurrKey%
return

SearchKey:
Gui, Submit, NoHide
GetOEMkey := ComObjCreate("WScript.Shell").Exec("wmic path softwarelicensingservice get OA3xOriginalProductKey").StdOut.Readall()
OEMkey := SubStr(GetOEMkey, 35, 29)
GuiControl,, OEMkey, %OEMkey%
return

ActivateOEM:
return



GuiClose:
ExitApp