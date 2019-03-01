#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.




GUI_UserNameAndNumberOfDisapproval(){
	
	; 카드 승인거절 횟수, 사용자 이름을 입력받기 위한 GUI
	Gui, Show, w340 h140, User Name And Number Of Disapproval	
	
	Gui, Add, Text, x52 y75 w160 h20 , The Number of Attempt Approval
	Gui, Add, CheckBox, x52 y99 w50 h30 vDecline1st, 1st
	Gui, Add, CheckBox, x112 y99 w50 h30 vDecline2nd, 2nd
	Gui, Add, CheckBox, x172 y99 w50 h30 vDecline3rd, 3rd
	
	Gui, Add, Button, x260 y28 w50 h50, Enter_Name


	Gui, Add, Text, x12 y38 w60 h30 +Center, User Name
	Gui, Add, Edit, x72 y28 w180 h20 vUserName
	



	;GUI시작 시 포커스를 UserName 입력칸에 위치
	GuiControl, Focus, UserName

	; 창이 종료될때까지 대기
	WinWaitClose, User Name And Number Of Disapproval
	
	
/*	
	;1st, 2nd, 3rd 버튼 누르면 Credit Card Attempt GUI창 종료
	;1st 버튼 눌렀으면 1 리턴
	;2nd 버튼 눌렀으면 2 리턴
	;3rd 버튼 눌렀으면 3 리턴

	Click_1st:
	WinClose ;, User Name And Number Of Disapproval
	return 1
	
	Click_2nd:
	WinClose, User Name And Number Of Disapproval
	return 2
	
	Click_3rd:
	WinClose, User Name And Number Of Disapproval
	return 3
*/	
	;Enter 버튼 누르면 User Name창 종료
	ButtonEnter_Name:	
	WinClose, User Name And Number Of Disapproval
	
	; GUI에서 입력받은 값을 사용하기 위해
	Gui Submit, nohide	
	GUI, Hide
	















/*	
	
;	Gui, Add. Button, x92 y28 w30 h30, ASDF

;	Gui, Add, Button, x260 y28 w65 h30, Enter_Name


	;GUI시작 시 포커스를 UserName 입력칸에 위치
	GuiControl, Focus, UserName
	
	; 창이 종료될때까지 대기
	WinWaitClose, User Name And User Name And Number Of Disapproval
	


	;Enter 버튼 누르면 User Name창 종료
	ButtonEnter_Name:
	WinClose, User Name

;	ExitApp, User Name
	
;	GuiClose:
;	ExitApp	
	
;	MsgBox, infunction`n%UserName%



	



	
	
		; 몇번째 Decline인지 입력받기 위한 Gui 창
	Gui, Show, w196 h93, User Name And Number Of Disapproval
	
	Gui, Add, Text, x12 y19 w160 h20 , The Number of Attempt Approval
	Gui, Add, Button, x12 y49 w50 h30 gClick_1st, 1st
	Gui, Add, Button, x72 y49 w50 h30 gClick_2nd, 2nd
	Gui, Add, Button, x132 y49 w50 h30 gClick_3rd, 3rd


	;GUI시작 시 포커스를 1st 버튼에 위치
	GuiControl, Focus, 1st
	
	; 창이 종료될때까지 대기
	WinWaitClose, User Name And Number Of Disapproval
	
	;1st, 2nd, 3rd 버튼 누르면 Credit Card Attempt GUI창 종료
	;1st 버튼 눌렀으면 1 리턴
	;2nd 버튼 눌렀으면 2 리턴
	;3rd 버튼 눌렀으면 3 리턴

	Click_1st:
	WinClose, User Name And Number Of Disapproval
	return 1
	
	Click_2nd:
	WinClose, User Name And Number Of Disapproval
	return 2
	
	Click_3rd:
	WinClose, User Name And Number Of Disapproval
	return 3
	
*/	
	
	
	
	
	
	
	
	
	
	return
}
