#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



FilePath := A_MyDocuments "\test.txt"

ol := ComObjCreate("Outlook.Application")
olMsg := ol.CreateItem(0) ; 0 = olMailItem

ol.Subject := "Text Report"
ol.Body := "The File in this emails path is " . FilePath
ol.To := "han1002@daum.net"

ol.Attachments.Add(FilePath) ; Add an Attachment
ol.Display ; Disply Email Item


ExitApp

Esc::
	ExitApp
