#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk

CConLAMBS(loc_of_MostRecentPo){
	
		
	; 마우스 커서는 사용자의 물리적 마우스 이동에 반응하지 않습니다
	BlockInput, MouseMove
			
	
	
	CoordMode, mouse, relative
	
	;LAMBS활성화 후 화면 초기화
	Start()	
	
/*	이미 GetInfoFromLAMBS 에서 주소 정보 얻었음
	;Billing Add1
	MouseMove, 300, 483
	;MouseMove, 60, 483
	MouseGetPos, , , , control
	ControlGetText, BillingAdd1, %control%, LAMBS
		
	;Billing Zip Code
	MouseMove, 96, 523
	;MouseMove, 60, 523
	MouseGetPos, , , , control
	ControlGetText, BillingZip, %control%, LAMBS
*/

	; LAMBS에 있는 인보이스 밸런스 값 읽어오기
	GetInvoiceBalanceOnLAMBS()


/*
	;Invoice Balance 값 얻기
	MouseMove, 870, 594
	Sleep 100		
	MouseGetPos, , , , control
	ControlGetText, InvoiceBalance, %control%, LAMBS
*/		



	;CCNumbers := []
	

	
	;CC 버튼 찾고 클릭
	FindCCButtonAndClickIt()
	Sleep 2000
		
	
	windowtitle := "Credit Card ( P999131 )"
	; CC번호 저장
	ControlGetText, CCNumbers, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	; 만료일 저장
	ControlGetText, ExpDate, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	; CVS 저장
	ControlGetText, CVV, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate 를 넘겨서 Month 와 Year 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate, "")

	
;	MsgBox, %CCNumbers%`n`n%ExpDate%`n`n%CVV%

	
	; CCNumbers 받아서 4개씩 읽어서 PartiallyCCNum1,2,3,4,5에 넣기
	StartingPos = 1
	
	loop, 5{
		PartiallyCCNum%A_Index% := SubStr(CCNumbers, StartingPos, 4)
		StartingPos := StartingPos + 4
	}
	
	
	
	
	; 두 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	;MouseClick, l, 240, 390
	MouseClick, l, 200, 465
	
	Sleep 100
	
	ControlGetText, CCNumbers2, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate2, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV2, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate2 를 넘겨서 Month2 와 Year2 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate2, 2)

	
	StartingPos = 1
	
	loop, 5{
		2PartiallyCCNum%A_Index% := SubStr(CCNumbers2, StartingPos, 4)
		StartingPos := StartingPos + 4
	}
	
	
	
	
	
	; 세 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	;MouseClick, l, 240, 410
	MouseClick, l, 200, 480
	
	Sleep 100
	
	ControlGetText, CCNumbers3, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate3, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV3, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate3 를 넘겨서 Month3 와 Year3 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate3, 3)

	
	StartingPos = 1
	
	loop, 5{
		3PartiallyCCNum%A_Index% := SubStr(CCNumbers3, StartingPos, 4)
		StartingPos := StartingPos + 4
	}	
	
	
	
	
	; 네 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	;MouseClick, l, 240, 430
	MouseClick, l, 200, 495
	
	Sleep 100
	
	ControlGetText, CCNumbers4, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate4, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV4, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate4 를 넘겨서 Month4 와 Year4 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate4, 4)

	
	StartingPos = 1
	
	loop, 5{
		4PartiallyCCNum%A_Index% := SubStr(CCNumbers4, StartingPos, 4)
		StartingPos := StartingPos + 4
	}		
	
	



;	MsgBox, 같은 값 있나 비교`n`n%CCNumbers%   %CCNumbers2%   %CCNumbers3%   %CCNumbers4%`n`n`n`n%ExpDate%   %ExpDate2%   %ExpDate3%   %ExpDate4%`n`n`n`n%CVV%   %CVV2%   %CVV3%   %CVV4%
		
	
	
	
	; 이것이 중요한 처리.
	; 이렇게 값을 초기화 해 줌으로서 다른 함수에서 램스의 cc 정보를 얻어갈 때 각 CCNumbers2,3,4 변수에 값이 있는지 없는지 여부를 갖고 행동을 결정하는 데 사용된다
	; 두 번째 카드 번호가 첫 번째와 같으면 두 번째 카드정보 포함 나머지 cc 정보 값들을 초기화하기
	
	; 하지만 만약 두 번째 cc 정보는 있는데 세 번째 cc가 없으면 어떡하냐?
	; 할 수 없음. 그냥 3,4번째 cc 정보는 두 번째 cc 정보가 그대로 카피됨.
	; 그것까지 구현하기 귀찮음
	IfEqual, CCNumbers, %CCNumbers2%
	{
		Loop, 4{
			CCNumbers%A_Index% :=
			ExpDate%A_Index% :=
			CVV%A_Index% :=
		}
	}
	

	
	
	
	
	
	
	
	
	
	;MsgBox, %PartiallyCCNum1%    %PartiallyCCNum2%    %PartiallyCCNum3%    %PartiallyCCNum4%    %PartiallyCCNum5%
	
	
	;cc 창을 닫기
	WinClose, %windowtitle%
	
	

	; 사용자에게 마우스 커서를 이동하도록 허용합니다.
	BlockInput, MouseMoveOff
	
	
	
	

	;cc정보 띄워주기
;	MsgBox, Invoice Balance : %InvoiceBalance%`n`nCC number : %PartiallyCCNum1%    %PartiallyCCNum2%    %PartiallyCCNum3%    %PartiallyCCNum4%    %PartiallyCCNum5%`n`nExpiration Date : %ExpDate%`n`nSecurity Code : %CVV%`n`nAddress1 : %Address1%`n`nZip Code : %ZipCode%


	; CCNumbers 에 값이 있으면, 즉 첫 번째 CC 번호가 있으면 cc정보 띄워주기
	if(CCNumbers)
	{
		MsgBox, %InvoiceBalance%`n`n%PartiallyCCNum1%   %PartiallyCCNum2%   %PartiallyCCNum3%   %PartiallyCCNum4%   %PartiallyCCNum5%`n`n%ExpDate%`n`n%CVV%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%2PartiallyCCNum1%   %2PartiallyCCNum2%   %2PartiallyCCNum3%   %2PartiallyCCNum4%   %2PartiallyCCNum5%`n`n%ExpDate2%`n`n%CVV2%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%3PartiallyCCNum1%   %3PartiallyCCNum2%   %3PartiallyCCNum3%   %3PartiallyCCNum4%   %3PartiallyCCNum5%`n`n%ExpDate3%`n`n%CVV3%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%4PartiallyCCNum1%   %4PartiallyCCNum2%   %4PartiallyCCNum3%   %4PartiallyCCNum4%   %4PartiallyCCNum5%`n`n%ExpDate4%`n`n%CVV4%`n`n%BillingAdd1%`n`n%BillingZip%
		
		return
	}
	else{ ;CC창 안에 정보 없으면 안내문 띄운 뒤 5초 뒤 자동으로 나가기
		MsgBox, , , No Credit Card Information at LAMBS`n`n`nTHIS WINDOW WILL BE CLOSED IN 5 SECONDS, 5

		WinClose, %windowtitle%
		
		; 쇼,전화,이메일 주문에서 CConLAMBS호출했으면 Call for CC 띄우고 끝내기
		if(loc_of_MostRecentPo = 1){
			;MsgBox, , , It's Call for CC`n`n`nTHIS WINDOW WILL BE CLOSED IN 5 SECONDS, 5
			MsgBox, It's Call for CC
			Reload
		}
		else{ ;다른 곳 주문이면 웹에서 이미 한 번 디클라인 난 뒤 CConLAMBS 호출한거니까 LAMBS에 CC정보 없으면 디클라인 함수 호출 후 아예 끝내기
			DeclineProcessing()
			Reload
		}
		
		return
	}

		;MsgBox, %InvoiceBalance%`n`n%PartiallyCCNum1%   %PartiallyCCNum2%   %PartiallyCCNum3%   %PartiallyCCNum4%   %PartiallyCCNum5%`n`n%ExpDate%`n`n%CVV%`n`n%BillingAdd1%`n`n%BillingZip%, , , , 2





	; 두 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	MouseClick, l, 240, 390
	CCNumbers :=
	
	Sleep 100
	
	ControlGetText, CCNumbers2, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate2, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV2, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	StartingPos = 1
	
	loop, 5{
		2PartiallyCCNum%A_Index% := SubStr(CCNumbers2, StartingPos, 4)
		StartingPos := StartingPos + 4
	}
	; CCNumbers2 값이 있다면 보여주기
	if(CCNumbers2)
		MsgBox, %InvoiceBalance%`n`n%2PartiallyCCNum1%   %2PartiallyCCNum2%   %2PartiallyCCNum3%   %2PartiallyCCNum4%   %2PartiallyCCNum5%`n`n%ExpDate2%`n`n%CVV2%`n`n%BillingAdd1%`n`n%BillingZip%
	
	





	
	; 세 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	MouseClick, l, 240, 410
	CCNumbers :=
	
	Sleep 100
	
	ControlGetText, CCNumbers3, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate3, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV3, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	StartingPos = 1
	
	loop, 5{
		3PartiallyCCNum%A_Index% := SubStr(CCNumbers3, StartingPos, 4)
		StartingPos := StartingPos + 4
	}
	
	; CCNumbers3 값이 있다면 보여주기	
	if(CCNumbers3)
		MsgBox, %InvoiceBalance%`n`n%3PartiallyCCNum1%   %3PartiallyCCNum2%   %3PartiallyCCNum3%   %3PartiallyCCNum4%   %3PartiallyCCNum5%`n`n%ExpDate3%`n`n%CVV3%`n`n%BillingAdd1%`n`n%BillingZip%	








	;CC 창 닫힐 때까지 기다리기
	CheckTheWindowPresentAndWaitUntillItClose(windowtitle)	

	
	return
}

