#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk

GUI_UserNameInput(){

	; 사용자 이름을 입력받기 위한 GUI
	Gui, Show, w370 h88, User Name

	Gui, Add, Text, x12 y28 w60 h30 +Center, User Name
	Gui, Add, Edit, x72 y28 w180 h30 vUserName
	Gui, Add, Button, x260 y28 w50 h50, Enter_Name


	;GUI시작 시 포커스를 UserName 입력칸에 위치
	GuiControl, Focus, UserName
	
	; 창이 종료될때까지 대기
	WinWaitClose, User Name
	


	;Enter 버튼 누르면 User Name창 종료
	ButtonEnter_Name:
	WinClose ;, User Name
	
	
;	ExitApp, User Name
	
;	GuiClose:
;	ExitApp	
	
;	MsgBox, infunction`n%UserName%


	; GUI에서 입력받은 값을 사용하기 위해
	Gui Submit, nohide



	
	return
}
