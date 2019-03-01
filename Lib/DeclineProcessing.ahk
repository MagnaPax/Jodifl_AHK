#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk


;#Include GUI_UserNameAndNumberOfDisapproval.ahk


/*
#Include GUI_UserNameInput.ahk
#Include GUI_NumberOfDisapproval.ahk
*/

;global UserName


DeclineProcessing(){
	
	Start()
	
	CoordMode, mouse, Relative
	

	;security 버튼 풀기
	FindSecurityButtonAndClickItThenInputNumber1()

	


; 카드 승인거절 횟수, 사용자 이름을 입력받기 위한 GUI 함수 호출
	GUI_UserNameAndNumberOfDisapproval()
;	MsgBox, DECLINE`n%UserName%
	
	
	
	;각 Decline 횟수에 따라 Invoice Memo에 기록하는 함수 호출
	if(Decline1st)		
		InputDeclineAtInvoiceMemo(1, UserName)
	else if(Decline2nd)
		InputDeclineAtInvoiceMemo(2, UserName)
	else if(Decline3rd)
		InputDeclineAtInvoiceMemo(3, UserName)





	;LAMBS 저장
	SendInput, {F8}
	
	WinWaitClose, Status




	
/*	
	;저장 후 Status창이 나타나면 없어질때까지 대기
	IfWinExist, Status
		CheckTheWindowPresentAndWaitUntillItClose("Status")	
*/


	;security 버튼 풀기
	FindSecurityButtonAndClickItThenInputNumber1()	
	
	; 아웃룩 없으면 실행하도록 메세지로 알림
	IfWinNotExist, ahk_exe OUTLOOK.EXE
		MsgBox, Is Outlook Activated?
		
	;Alt+e 누르고 Alt+m 눌러서 아웃룩 호출
;	Send, !e
;	Send, !m


	
	; Email 체크 버튼 누르기
	;MouseMove, 955, 250
	;MouseGetPos, , , , control
	;ControlClick, %control%, LAMBS, l
	;ControlSend, %control%, {Space}, LAMBS
	;ControlSendRaw, %control%, {Space}, LAMBS
	;ControlSend, %control%, {Space}, LAMBS
	;ControlClick, %control%, LAMBS
	
	;ControlClick, x955 y250, LAMBS


;	emailSearch()



	;Email 버튼이 체크 안 되었는 지 찾아보고 체크 안 되어 있으면 체크하기
	;이메일 보낼 때 사용
	jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked.png
	CheckEmailButtonOrReleaseIt(jpgLocation)
	jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked_Activated.png
	CheckEmailButtonOrReleaseIt(jpgLocation)
	jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked_Activated2.png
	CheckEmailButtonOrReleaseIt(jpgLocation)










	; Invoice(M) 버튼 클릭하기
;	ControlClick, x609 y178, LAMBS -  Garment Manufacturer & Wholesale Software


	; Invoice(M) 버튼 클릭하기
	WinActivate, LAMBS
	Send, !m
	Sleep 100
	
	;아웃룩 창 이름 windowtitle 변수에 넣기
	windowtitle = Untitled - Message (HTML)	
	
	
;	WinWaitClose, Status ;, , 3











/* ; 이거 해보려다 결국 못했음

	; <아웃룩 없으면>
	; Status 로딩 -> (Status 안 없어진 상태에서) Information 창 뜸 -> 클릭하면 Status, Information 창 둘 다 사라짐
	
	; Information 창 나타나기를 6초동안 기다림
	WinWait, Information, , 6


	; 만약 아웃룩 실행하라는 경고창 뜨면
	IfWinExist, Information
	{
		;아웃룩 실행 후 기다렸다 처음부터 다시 시작
		
		; 엔터 눌러서 information창 없앰
		SendInput, {Enter}
		
		; 아웃룩 실행
		run, outlook
		
		; 아웃룩 활성화 될 때까지 대기
		WinWaitActive, %windowtitle%
		MsgBox, outlook activated!!
		sleep 2000
		
		
		WinWait, Windows Security, , 3
		IfWinExist, Windows Security
			WinClose
		ControlClick, x140 y284, Windows Security
		
		WinMinimize, Outlook
		
		WinMinimizeAll
		
		; Invoice(M) 버튼 클릭해서 이메일 출력 창 호출
		WinActivate, LAMBS
		Send, !m
		Sleep 100
	}
	
	
;	Sleep 1000
*/






	
	;아웃룩 창이 뜰 때까지 기다림
	WinWaitActive, %windowtitle%
	
	Sleep 1000


	;도착주소에 전역변수인 Email값 넣기
	ControlClick RichEdit20WPT1, %windowtitle%, , l, 3

	SendInput, %Email%
	SendInput, {Tab}
	
	
	

	
	;날짜를 입력하기 위한 처리
	FormatTime, CurrentDateTime, ,MM/dd/yyyy
	title = CREDIT CARD DECLINED NOTIFICATION %CurrentDateTime% %Invoice_No% JODIFL
	
	;Subject에 제목 넣기
	Control_InputText("RichEdit20WPT4", title, windowtitle)


	;INSERT Tab 버튼 클릭
	ControlClick, x176 y43, %windowtitle%
	;Signature 버튼 클릭
	ControlClick, x222 y97, %windowtitle%


	;decline2 찾아서 엔터
	Send, {Down}
	Sleep 200	
	Send, {Down}
	Sleep 200	
	Send, {Down}
	Sleep 200
	Send, {Down}
	Sleep 200	
	Send, {Down}
	Sleep 200	
	Send, {Enter}
	Sleep 200
	
	;Send Button 단축키 누르기
	SendInput, !s
	
	
	; 이메일 발송 창이 닫힐때까지 기다리기
	WinWaitClose, %windowtitle%


	
	; Decline 났으니까 UPS 재설정 하기
	Loop, 5
	{
		; Alt + Del 눌러서 잘못 입력한 것 취소하기. (작동 되는지는 모르겠음)
		ControlSend, , {!Delete}, UPS WorldShip - Remote Workstation
;		Sleep 200
		ControlSend, Button1, {Enter}, UPS WorldShip
;		Sleep 200
		ControlClick, Button1, UPS WorldShip
;		Sleep 200
		ControlSend, , {F2}, UPS WorldShip - Remote Workstation
;		Sleep 200
		ControlSend, , {F5}, UPS WorldShip - Remote Workstation
;		Sleep 200
	}
	
	
	; UPS WorldShip 에 경고창 나오면 없애기
	Loop, 5
	{
		; Alt + Del 눌러서 잘못 입력한 것 취소하기. (작동 되는지는 모르겠음)
		ControlSend, , {!Delete}, UPS WorldShip - Remote Workstation

		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip

		ControlSend, , {F2}, UPS WorldShip - Remote Workstation
		ControlSend, , {F5}, UPS WorldShip - Remote Workstation
		
		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip
	}
	
	
	
	


;버튼이 체크 되었는 지 찾아보고 체크 되어 있으면 체크를 풀기
;일반 인쇄 시 혹은 이메일 보낸 후 체크버튼 풀 때 사용
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked.png
CheckEmailButtonOrReleaseIt(jpgLocation)
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked_Activated.png
CheckEmailButtonOrReleaseIt(jpgLocation)


	; UPS WorldShip 에 경고창 나오면 없애기
	WinActivate, UPS WorldShip - Remote Workstation
	Loop, 5
	{
		; Alt + Del 눌러서 잘못 입력한 것 취소하기. (작동 되는지는 모르겠음)
		ControlSend, , {!Delete}, UPS WorldShip - Remote Workstation
;		Sleep 200

		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip

;		Sleep 200
		ControlSend, , {F2}, UPS WorldShip - Remote Workstation
;		Sleep 200
		ControlSend, , {F5}, UPS WorldShip - Remote Workstation
;		Sleep 200
		
;		key = F2
;		Control_SnedButton(Key, windowtitle)
;		Sleep 1000
;		key = F5
;		Control_SnedButton(Key, windowtitle)

		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip
	}
	
	
	
	; 이메일 내용을 파일에 저장해놨다가 해출 해볼까	

	return	
}