#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;GUI - OEM Key ( Show, Export and Activate )
Gui, Add, Groupbox, x10 y10 w400 h50, OEM key
Gui, Add, Edit, ReadOnly x15 y30 w300 h20 vOEMkey
Gui, Add, Button, gExportOEMkey x316 y30 w90 h20, Export
Gui, Add, Button, gAbout x10 y80 w20 h20, ?
;Gui, Add, Button, gTest x136 y80 w90 h20, Test
Gui, Add, Button, gSearchKey x226 y80 w90 h20, Grab Key!
Gui, Add, Button, gActivateOEM x316 y80 w90 h20, Activate OEM
Gui, Show, w420 h120, OEM Key Grabber v1.0
Return

ExportOEMkey:
if ( OEMkey == "" )
{
	MsgBox,,Error, No key grabbed!
	return
}
FileSelectFile, FileSelectVar, S8, %A_Desktop%\OEM-key.txt, Save OEM key
FileDelete, %FileSelectVar%
FileAppend, 
(
################################################
#Grabbed with OEM Key Grabber!                 #
#https://github.com/Pyrree/OEM-Key-Grabber     #
################################################

OEM key:.......... %OEMkey%
PC Name:.......... %A_ComputerName%
User Name:........ %A_UserName%
Windows Version:.. %A_OSVersion%
),%FileSelectVar%
return

About:
MsgBox,, About,
(
Welcome! :)

This is a small utility that looks up the OEM key on your motherboard,
you can also activate Windows with the key or export it.

Creator: Pyrre
Github: https://github.com/Pyrree/OEM-Key-Grabber
)
return

Test:
MsgBox,, Test, %A_ComputerName%
return

SearchKey:
Gui, Submit, NoHide
GetOEMkey := ComObjCreate("WScript.Shell").Exec("wmic path softwarelicensingservice get OA3xOriginalProductKey").StdOut.Readall()
OEMkey := SubStr(GetOEMkey, 35, 29)
GuiControl,, OEMkey, %OEMkey%
return

ActivateOEM:
if ( OEMkey == "" )
{
	MsgBox,,Error, No key grabbed!
	return
}
ComObjCreate("WScript.Shell").Exec("cscript /b C:\Windows\System32\slmgr.vbs -ipk %OEMkey%")
sleep, 5000
ComObjCreate("WScript.Shell").Exec("cscript /b C:\Windows\System32\slmgr.vbs -ato")
return


GuiClose:
ExitApp
