#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


GUI_TrackingNumber(){
	

	; Generated using SmartGUI Creator for SciTE
	Gui, Show, w287 h88, JODIFL

	Gui, Add, Text, x12 y28 w60 h30 +Center, Tracking Number
	Gui, Add, Edit, x72 y28 w180 h30 vTrackingNumber

	;GUI시작 시 포커스를 Invoice_No 입력칸에 위치
	GuiControl, Focus, TrackingNumber

	GuiClose:
	ExitApp	
	
	 return
}