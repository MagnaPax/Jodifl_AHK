#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



GUI_NumberOfDisapproval(){
	
	; 몇번째 Decline인지 입력받기 위한 Gui 창
	Gui, Show, w196 h93, Number of disapproval
	
	Gui, Add, Text, x12 y19 w160 h20 , The Number of Attempt Approval
	Gui, Add, Button, x12 y49 w50 h30 gClick_1st, 1st
	Gui, Add, Button, x72 y49 w50 h30 gClick_2nd, 2nd
	Gui, Add, Button, x132 y49 w50 h30 gClick_3rd, 3rd


	;GUI시작 시 포커스를 1st 버튼에 위치
	GuiControl, Focus, 1st
	
	; 창이 종료될때까지 대기
	WinWaitClose, Number of disapproval
	
	;1st, 2nd, 3rd 버튼 누르면 Credit Card Attempt GUI창 종료
	;1st 버튼 눌렀으면 1 리턴
	;2nd 버튼 눌렀으면 2 리턴
	;3rd 버튼 눌렀으면 3 리턴

	Click_1st:
	WinClose, Number of disapproval
	return 1
	
	Click_2nd:
	WinClose, Number of disapproval
	return 2
	
	Click_3rd:
	WinClose, Number of disapproval
	return 3
	
	; GUI에서 입력받은 값을 사용하기 위해
	Gui Submit, nohide
	
	return
}