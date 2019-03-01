#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk

;해외 배송인지 확인하는 것 먼저


	GetInfoFromJODIFL_WEB(PO_W){
		MsgBox, Go to JODIFL_WEB
		Reload
		return
	}