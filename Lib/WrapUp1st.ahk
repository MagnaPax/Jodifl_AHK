#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; LAMBS의 Invoice Memo에 Tracking Number 적어넣는 함수
WrapUp1st(Consolidation, loc_of_MostRecentPo, Invoice_Memo, Invoice_No, Delivery){
	
;	MsgBox, % loc_of_MostRecentPo
	
	windowtitle = LAMBS -  Garment Manufacturer & Wholesale Software
	
	;비활성으로 만들자. 이거 하는동안 웹 작업 하도록
	
	;LAMBS에 Tracking Number 적어넣기
	Start()
	FindSecurityButtonAndClickItThenInputNumber1()

	;No of Box로 이동한 뒤 탭키로 Invoice Memo 넘어가기
	MouseClick, l, 590, 498
	SendInput, {Tab}
	Sleep 200
	
	IF(Consolidation)
		SendInput, {Space}CONSOLIDATION
	else if(Delivery)
		SendInput, {Space}DELIVERY TO %Delivery%		
	else
		SendInput, {Space}%TrackingNumber%
		
	
;	MsgBox, % TrackingNumber
	
	;LAMBS 저장
	SendInput, {F8}
	
	
	;저장 후 Status창이 나타나면 없어질때까지 대기
	IfWinExist, Status
		CheckTheWindowPresentAndWaitUntillItClose("Status")		

/*
	;저장 후 1초 동안 Status창이 나타날 때까지 기다리다가
	;나타나면 창 종료
	;이렇게 짜려다가 말이 안돼서 관둠
	
	WinWait, Status, , 1
	if ErrorLevel
	{
		MsgBox, WinWait timed out.
		return
	}
	else
		WinClose  ; WinWait가 발견한 창을 최소화합니다.
*/	



	; Pay 버튼 클릭
	MouseClick, l, 745, 720
	
	; Confirm 창 나오면
	WinWait, Confirm
	SendInput, {Enter}
	Sleep 1500




	;버튼이 체크 되었는 지 찾아보고 체크 되어 있으면 체크를 풀기
	;일반 인쇄 시 혹은 이메일 보낸 후 체크버튼 풀 때 사용
	jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked.png
	CheckEmailButtonOrReleaseIt(jpgLocation)
	jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked_Activated.png
	CheckEmailButtonOrReleaseIt(jpgLocation)


	;WinActivate, LAMBS
	
	;Alt+m 눌러서 인쇄
	Send, !m


	
	;저장 후 Status창이 나타나면 없어질때까지 대기
;	WinWait, Status
	WinWaitClose, Status
	


	Sleep 500
	FindSecurityButtonAndClickItThenInputNumber1()
	FindSecurityButtonAndClickItThenInputNumber1()
	
	
	; 인보이스에서 패션고, JODIFL, LA쇼룸을 못 찾았으면 고객에게 이메일 발송
	; 만약 패션고, JODIFL, LA쇼룸의 인보이스를 찾았으면 인보이스 중에 쇼,전화,이메일 오더가 있는지 확인 후 있으면 고객에게 이메일 발송
	if(loc_of_MostRecentPo = 1)
		SendingEmailCustomerInvoice(Invoice_No)
	else
		FindingPOsOfShowPhoneEmail(Invoice_Memo, Invoice_No)


	
	
	
	
	
	return
}