﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


global counter_forN41

counter_forN41 = 0

;~ class CommN41{
class CommN41{

	; N41 로그인
	runN41(){
		
		Sleep 500
		
		N41_login_wintitle := " N41"
		N41_login_wintitle := "ahk_class FNWND3126"
		
		
;		WinWaitActive, %N41_login_wintitle%
;		WinMaximize, %N41_login_wintitle%

		WinMaximize, ahk_exe nvlt.exe


		; 이미 N41 열려있으면 메소드 중단하고 나오기
;		IfWinExist, %N41_login_wintitle%
		IfWinExist, ahk_exe nvlt.exe
			return


		Run, nvlt.exe, C:\NVLT
			
		; 로그인 창이 활성화 될 때까지 기다리기
		WinWaitActive, %N41_login_wintitle%
			
		; 아이디 입력
		ControlSend, Edit1, c123, %N41_login_wintitle%
		Sleep 200
			
		; Ok 버튼 클릭
		ControlClick, Button2, %N41_login_wintitle%, , l
			
		; 로그인 창이 닫힐때까지 기다리기
		WinWaitClose, %N41_login_wintitle%
		
		
		Sleep 500
		
		; 로그인 아이디 틀렸으면 나오는 경고창		
		IfWinActive, N41 Log In
			MsgBox, login again
		
		; 작업창 활성화 될때까지 기다리기
		WinWaitActive, %N41_login_wintitle%
		
		Sleep 4000
		
		while (A_cursor = "Wait")
			Sleep 3000
		
		Sleep 1000
		
		
		return
	}
	
	
	BasicN41Processing(){
		
		N41_login_wintitle := " N41"
		
		CommN41.runN41()
		
		IfWinActive, Connection Error
			MsgBox, 4100, Connection Error Warnning, Click Ok to continue
		
;		WinActivate, %N41_login_wintitle%
		WinActivate, ahk_exe nvlt.exe
		
	}
	














	; Customer Master Tab 클릭하기
	ClickCustomerMasterTab(){
		
		; Customer Master Tab 클릭하기
		Text:="|<Customer Master Tab>*170$81.Tzzzzzzzzzzzzw0D00000000000026000000000000kM000000000003y000000000011y8D00E0000Mk8n168020000361MEEUFSxnyQsRr42642+GPGIK2e0Eb0UFMG+GyULH1tg428mFGI42GUDMElHGHOGkUEI27z3vvniGHo22szzs00000000007zz00000000001zzs0000000000Dz000000000004"
		if ok:=FindText(609,104,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		else{
			; 못 찾으면 계속 재귀호출 해서 찾아보기
			Sleep 700
			CommN41.ClickCustomerMasterTab()			
			;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
			;~ MsgBox, 262144, No CM Tab Warnning, Please Open Customer Master
		}			
		
		return		
	}
	
	
	; Customer Master 에서 카드 정보 창 아이콘 클릭해서 열기
	OpenRegisterCreditCard(){
		
		; N41 로그인
		CommN41.BasicN41Processing()

		; Customer Master Tab 클릭하기
		CommN41.ClickCustomerMasterTab()
		
		Sleep 500
		
		; 카드 아이콘 클릭
		Text:="|<CC ICON>*183$17.zzzzlzs1w03s07m03Y0300600CDzQTyw1Zw03s0Dzzw"
		if ok:=FindText(697,130,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click			
		}		
		
	}

	
	; SO Manager 클릭
	OpenSOManager(){		
		
		; N41 로그인
		CommN41.BasicN41Processing()
		
		; SO Manager 클릭
		Text:="|<SO Manager>*161$75.zzzzzzzzzzzzzz000000000000A000000000001E00000000003t00000000001z47bkMk00000CTV5X36000001U4888RrbrXnbDUVl12e2m2mWkz41c8LHoHoLo1wU512GWWWWUUVY8gMEIoIoK4CQVsy22yWySSVz400000000E03kU0000000w00040000000000zzU0000000004"
		if ok:=FindText(275,104,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		
		else
		{
			; 못 찾으면 계속 재귀호출 해서 찾아보기
			Sleep 500
			CommN41.OpenSOManager()
		}		
		
	}
	
	; CustomerMaster 에 있는 뉴 버튼 클릭하기
	ClickNewButtonOnCustomerMaster(){
	
		Sleep 700		

		Text:="|<New Button on Customer Master>*188$35.zzzzzzzzzzzzzzzzzzzzzzzxz00Tznw00zzztzxzzznzvzvzbzrzqzDzzz9yTzzzTwztzyztznzxzs61zzTzzDzwzzyTzzzzzzzzzzzzzzzzzzzjzzzzyTzzzzw"
		if ok:=FindText(239,130,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 500
		}

		return
	}


	; Sales Order 클릭
	ClickSalesOrderOnTheMenuBar(){	

		; Sales Order on the Menu bar
		Text:="|<Sales Order on the Menu bar>*137$73.zw0000000000E3000000000081E0000000004DY0040000102Tl1s2007k0U1CTV41006A0E0a0EUwbD22ttnns8Q1IIV1F558z41bfv0UcWyY7m0IJ0MEIFEG8N2+OkYAO9g9CQVtxDS3t3nobwE000000000Ew80000000008040000000007zy0000000002"
		if ok:=FindText(122,57,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}
		else
		{
			; 못 찾으면 계속 재귀호출 해서 찾아보기
			Sleep 700
			CommN41.ClickSalesOrderOnTheMenuBar()
		}
		
		return
	}


	; Add(+) Button
	ClickAdd(){		

		; Add(+) Button
		Text:="|<Add(+) Button of Sales Order>*147$14.0000030180G05UDz4ztTyDz0S07U1s0A000008"
		if ok:=FindText(458,128,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}

				
		; 커서 상태가 작업처리중이면 끝날때까지 기다리기
		while (A_cursor = "Wait")
			Sleep 2000				

		Sleep 1000
		
		return
	}


	; Save Button
	ClickSave(){

		; Save Button
		Text:="|<Save Button of Sales Order>*167$58.k0k0w7U0TXU70401011j0w0G0Y041S7U182E0E4ww04U901TFzU0G0Y0w13w0182E2LoDk04zt090FzU0E040bzDD013sE201sS04EF080D0w0Fz40zzs1k14QE3zw0007lT0A00000Tzw0TzU"
		if ok:=FindText(381,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.ClickSave()
		}
		
		return
	}
	
	
	
	; 맨 위 메뉴바에 있는 Customer 클릭하기
	ClickCustomerOnTheMenuBar(){

		Text:="|<Customer on the top menu>*144$86.000000000000000000000000000000000000000000007U0000000000Dk2600000000002010k0000000000U09k00000000008E3l1s020000002S0EFW00U000000bU00E8jStzCRs09U2242+GPGIKG02Q0C10WkYIZx600X7YEE8X959EEM081u26+OGPGK4G028VUkyywvYYx7U0bMTw00000000009rzz000000000027zzk0000000000Uzw00000000000Dk0000000000000000000000000000U"
		if ok:=FindText(1969,56,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		
		return
	}
	
	
	; Customer Master 탭에 있는 List 클릭하기
	ClickListOnCustomerMaster(){
		
		MouseClick, l, 261, 164		

		Text:="|<List on Customer Master>*184$59.s0TzU00001k0zz000003U10200000702To220U0C040840100Q09zE8/r00s0E0UEQY01k0bx0Ui803U102117E0702To23aU0C04087ptU0Q083k00000s0E7U00001k0UC000003U1zs000004"
		if ok:=FindText(243,155,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}	
		
		return
	}
	
	
	; Customer Master 탭에 있는 검색조건이 이메일로 된 공란 클릭
	ClickTheBlankModeByEmail(){
		
		MouseClick, l, 1030, 133
		
		return
	}
	
	
	; Customer Master 탭에 있는 Create SO 클릭하기
	ClickCreateSO(){		

		Text:="|<Create So>*183$16.zzzs0z03tzD4AtUHaTCMAtkTbryNwtXnbA6TkNznU7zzzs"
		if ok:=FindText(581,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		
		return
	}
	
	; Pick Ticket 탭에 있는 Refresh Button 클릭하기
	ClickREfresh(){

		Text:="|<REFRESH BUTTON>*172$59.McA00U0000lEM0300001WU00C000035000zU03k6+000zU08EAI600lU0UTMcA00VU102lE00U10205WU010204Dv50032008U6+303600G0AI603y00c0Mc003y01U2lE000s01ztWU001U000351U0200008"
		if ok:=FindText(235,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 700
		}

		return
		
	}
	
	; Approved 됐는지 화면에서 찾아본 뒤 찾았으면 1을 리턴하고 못 찾았으면 0을 리턴
	DoesThisPickTicketApproved(){

		Text:="|<CC Approved>*202$48.6000000360000003D7bbSnST96qqHHHHNYoonOznTYoonSknkqqoHAPHUrbYSAST0440000004400000U"
		if ok:=FindText(1637,299,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  return 1
		}
		
		return 0
	}
	
	
	ClickCreatePickTicketButton(){
		

		Text:="|<Create Pick Ticket Button>*157$16.00E0200F028zl428Ll1Tw5rELB1zw5zELx1047zkU"
		if ok:=FindText(1132,127,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}	
		
		return
	}
	
	
	; SO Manager 에 있는 refresh 버튼 클릭
	ClickREfreshButtonOnSOManager(){
		
		Sleep 500

		Text:="|<REFRESH BUTTON ON SO Manager>*186$34.800000U0z00207y00808Q00U60s020w1U087k600U60y020M1k081k600U3VU0CzzyzzvzjrzzzzVzzzzzzzzU"
		if ok:=FindText(308,128,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 1000
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.ClickREfreshButtonOnSOManager()
		}
		
		return
		
	} ; END - ClickREfreshButtonOnSOManager	
	
	

	
	
	; Customer Master 에 있는 리프레쉬 버튼 클릭하기
	ClickRefreshButtonOnCustomerMaster(){
		
		Text:="|<Refresh Button>*182$65.zzzzzzzzzzzU07zy1zzJ0z00Dzs1zw01yTzTztVztzzwzyzyTlzm00tzxzwTnzY01nzzzkzXz/zvbzzzny3yLzrDyTzXyDwjziTwzz7wzxTzS1UTz7bzyzyzznzz0Dzxzxzzbzz0Tzvzvzzzzzbzzk07zzzzzzzzzzw"
		if ok:=FindText(261,130,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}
				
		
		
	}
	
	
	; Sales Order 에 있는 리프레쉬 버튼 클릭하기
	ClickRefreshButtonOnSalesOrder(){		

		Text:="|<Refresh Button On Sales Order Tab>*169$34.00200000M00003U00U0Tk0100zU0801X0100260404080E0E0U1U1V00603600M07w01U0Ds0600700800M00001002"
		if ok:=FindText(260,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.ClickRefreshButtonOnSalesOrder()
			
		}
	}
	
	
	; 고객 코드 얻기 위해 Sales Order 에 있는 Customer 표시 찾아가서 엔터쳐서 클릭보드에 복사 후 CustCode 변수에 넣어서 리턴하기
	GetCustomerCode(){
		
		Clipboard := ""
		CustCode := ""
		
		Text:="|<Customer on Sales Order>*181$55.zzzzzzzzzk00000000M00000000A0000000060D00800030AU040001U42HrSTrbk61/999+OM30YsgobxA0UG7KOHUa0NdNd99HH07bbawYj9U00000000k00000000M00000000Dzzzzzzzzy"
		if ok:=FindText(247,219,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+7, Y+H//2
			Sleep 150
			
			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			CustCode := Clipboard
			Sleep 150

;			MsgBox, % CustCode
			
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
			if(!CustCode){
			;~ if(CustCode == ""){
;				MsgBox, no value in variable, restart the method
				CommN41.GetCustomerCode()
			}
			
			
			return CustCode
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.GetCustomerCode()
		}

	}
	
	
	; Customer PO 번호 얻어서 리턴하기
	GetCustPONumber(){
		
		Clipboard := ""

		Text:="|<Cust PO on Sales Order>*177$49.zzzzzzzzzzzzzzzzk0000000M0000000A00000006S00ES3kjN0088XAJc4bi4F2zo2GG28VBO19l1sEZh0YCUU8LynGnEE6NHDDDA8FsdU0000000k0000000M0000000Dzzzzzzzz"
		if ok:=FindText(573,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+5, Y+H//2
			Sleep 150		  
			
			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			CustomerPO := Clipboard
			Sleep 150
			
			return CustomerPO
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.GetCustPONumber()
		}				
		
		
	}
	
	
	
	; 왼쪽 메뉴바에 있는 Customer 클릭하기
	ClickCustomerMarkOnTheLeftBar(){


		Text:="|<>*140$101.00000000000000000000000000000000000U00000000000000020000000000000000Dzk00000000000000zzk6Ekk000000000148EAXXU00000000068oUR7D0000000000BEd0uSK0000000000O0G1xgA0000000000oCY3vwM0000000000gx86kkk0000000001zzUBVVU0000000001zy000000000000001kQ0000000000000010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000w00000000000000026000000000000000860000000000000009k000000000000000S83k0080000000000EEAk00k000000004000E6nnnnzXnrU000221UBanBqvAaBU0005s30PD6lgqzgS0007YE30q7BXNBUMC000DkE6NgqNgmNglg000VUk7ntwttavtXs0033zU00000000000007zz00000000000000Tzy00000000000000zzw000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000E0000000000000009U000000000000000G0000000000000000A0000000000000001w03l0A00000000000005a0M00000000000008Tank000000000FU0QPBgU000000000X00QnnzU000000000700BbbU0000000004C06P6Bg0000000000A07bANk000000000UQ000M00000000001Vs003U000000000040c00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000004"

		if ok:=FindText(55,798,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 150
		}
	}
	
	; 왼쪽 메뉴바에 있는 SO Manager 클릭하기
	ClickSOManagerOnTheLeftBar(){

		Text:="|<SO Manager>*146$58.SD1X000002966A00000888RnbXXnbsUVJ1F1FFEO25Yx4x5x0c8GIIIII4WFV1FFFFEHks45x5wwx00000000E00000000S02"
		if ok:=FindText(45,154,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 150
		}

	}
	
	
	; SO Manager 에 있는 Customer 표시 찾기. 고객 코드 입력하기 위함
	FindCustomerMarkToFillInTheBlank(){
		
		Text:="|<Customer on the top menu bar of SO Manager>*147$52.U000000020000000081s020000UMU08000210Wxvbwts42+GPGIKUE8g959TG10WAYIZ186+OGPGK4UDjjCt9DG00000000800000000U00000002"
		if ok:=FindText(497,127,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+20, Y+H//2
			Click
			Sleep 150
		}
		else
		{
			; 못 찾으면 계속 재귀호출 해서 찾아보기
			Sleep 700
			CommN41.FindCustomerMarkToFillInTheBlank()
		}
	}

	
	; Pick Ticket 탭에 있는 House Memo 에 메모 넣기 위해
	PutMemoIntoHouseMemoOnPickTicket(){

		Text:="|<HOUSE MEMO>*166$29.W0001400029t9szmGGGMYob7l9d3cWGGnH4wwwQ00000000000000000000000000000n0001a0003AQzj6NB9G/HuGaKY4ZAh9d+FOCGLc"
		if ok:=FindText(1096,225,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+30, Y+H//2
			Sleep 500
			Click
			Sleep 150
			Send, {End}
			Sleep 150
			Send, {Enter}
			Sleep 150
		}
		else
		{
			; 못 찾으면 계속 재귀호출 해서 찾아보기
			Sleep 700
			CommN41.PutMemoIntoHouseMemoOnPickTicket()
		}
		
		return
	}


	; Open Allocation 에 있는 Create Pick Ticket 버튼 클릭하기
	ClickCreatePickTicketButtonOnOpenAllocation(){
		
;		MsgBox, 262144, Title, Open Allocation 에 있는 Create Pick Ticket 버튼 클릭하기


		Text:="|<Create Pick Ticket Button on Open Allocation Of SO Manager>*177$248.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzk0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000800000000000000000000000400000000000004005000000Dk00Dk000000Ds0002000000000000010DzM000007z007z000000A1U0014000001k000000E211000001yE01yE03zy0404000WDU000AQ00180040zxU00000Ty00Ty00jyE1Dt01zl6A0007j000k001081M033007zU07zU0BzA0U080E8VXzTz1vrXzjjTUE20+00NU05zu05zu03jb08zW05wEkSoykSxhUPPTM40U0U03k0100U100U0xnk200U1zwA77zAAznkyrba1080800M00Tzs0Tzs0D9w0XsE0TT1Xlkn3zwwRhttUE20200D007zy07zy03Yj040407nkMwyAknzhaPzSM40U0U06M01zzU1zzU0nxk10601zw3vxzAMTTTySza1080803300Tzs0Tzs09zg0Ay00Tz00k00000000000E3zy00000402040200zx018007zk0C000000000004000000000Ty00Ty00TzU0I00104000000000000000000000003z003z0000006000Tz2"
		if ok:=FindText(1185,332,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)-10, Y+H//2
			Click			
			Sleep 150
		}		
	}
	
	
	; Sales Order 탭에 있는 Memo
	; 메모값 읽기 위해
	MemoOnSalesOrderTab(){
		
		Clipboard := ""

		Text:="|<Customer Memo On Sales Order Screen>*191$45.zzzzzzzw0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U06M000w00n0007U06NtzSw00z9heTU07vx9nw00hM9CTU05dh9Hw00hD9/rU000000w0000007U000000w0000007U000000w0000007U000006w000001zU000006w000000zU000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000zzzzzzzzU"

		if ok:=FindText(1030,249,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, (X+W)+2, Y+H//2
		  Click
		  Sleep 150
		  
		  Send, ^c
		  Sleep 150
		  
;		  MsgBox, % Clipboard




		; 클립보드 내용 CustMemoOnSOTab 변수에 넣기
		Sleep 700
		CustMemoOnSOTab := Clipboard	


		; web 주문에 메모가 있다면 추출한 뒤 변수에 저장함
		CustMemoOnSOTab := RegExReplace(CustMemoOnSOTab, "(.*)CreditCard\sID\sfor\spayment:\s.*", "$1")  ; $1 역참조를 사용하여 메모 내용을 돌려준다

		; FG 주문에 메모가 있다면 추출한 뒤 변수에 저장함
		CustMemoOnSOTab := RegExReplace(CustMemoOnSOTab, ".*notes:(.*),\sOrder.*", "$1")  ; $1 역참조를 사용하여 메모 내용을 돌려준다

		; LAS 주문에 메모가 있다면 추출한 뒤 변수에 저장함
		CustMemoOnSOTab := RegExReplace(CustMemoOnSOTab, ".*LA,\s(.*).*", "$1")  ; $1 역참조를 사용하여 메모 내용을 돌려준다

		; LAS 주문의 메모를 저장한 변수에 None 이 있다면 메모가 없다는 뜻이니까 그냥 지운다
		if(CustMemoOnSOTab == "None"){
			CustMemoOnSOTab := ""
		}
		

		StringUpper, CustMemoOnSOTab, CustMemoOnSOTab ; 대문자로 바꾸기
		  
		return CustMemoOnSOTab
		  
		}

	} ; MemoOnSalesOrderTab() 메소드 끝


	
	; Sales Order 탭에 있는 House Memo 읽어서 리턴
	HouseMemoOnSalesOrderTab(){
		
		Clipboard := ""
		

		Text:="|<house memo>*140$32.W00008U00028t9ssyGGGG8YYb7m998N0WGGHG8X7bXU00000000000000000000000000000000240000n0000AlnyQ3AYYd0VDd+E/G2GY2oYYd0Z799c"
		if ok:=FindText(1033,289,150000,150000,0,0,Text)
		{
			
			Clipboard := ""
			
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+5, Y+H//2
			Click
			Sleep 150
			  
			Send, ^c
			Sleep 150
			
			
	;		MsgBox, % "house memo`n`n" . Clipboard



			; 클립보드 내용 CustMemoOnSOTab 변수에 넣기
			Sleep 700
			HouseMemoOnSOTab := Clipboard
			
			; 메모내용이 UPDATED PO# 이면 지우기. 쓸데 없는 값
			if(HouseMemoOnSOTab == "UPDATED PO#"){
;				MsgBox, % HouseMemoOnSOTab . "`n`nHouseMemoOnSOTab 지울예정"
				HouseMemoOnSOTab := ""
			}


			HouseMemoOnSOTab := RegExReplace(HouseMemoOnSOTab, "Staff only notes:(.*)", "$1")  ; $1 역참조를 사용하여 Staff only notes: 이외의 메모 내용이 있으면 변수에 저장
			
			StringUpper, HouseMemoOnSOTab, HouseMemoOnSOTab ; 대문자로 바꾸기

			  
			return HouseMemoOnSOTab
			  
			}

	} ; HouseMemoOnSalesOrderTab() 메소드 끝
	



	
	; 화면의 4분변에 있는 Opne Allocation 의 빈 공간 클릭하기
	; 정확히 클릭하기 위해 왼쪽 메뉴바에 있는 Customer 찾아서 거기서 오른쪽으로 이동해서 클릭하기
	ClickEmptySpaceOnOpenAllocationArea(){

		Text:="|<>*140$101.00000000000000000000000000000000000U00000000000000020000000000000000Dzk00000000000000zzk6Ekk000000000148EAXXU00000000068oUR7D0000000000BEd0uSK0000000000O0G1xgA0000000000oCY3vwM0000000000gx86kkk0000000001zzUBVVU0000000001zy000000000000001kQ0000000000000010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000w00000000000000026000000000000000860000000000000009k000000000000000S83k0080000000000EEAk00k000000004000E6nnnnzXnrU000221UBanBqvAaBU0005s30PD6lgqzgS0007YE30q7BXNBUMC000DkE6NgqNgmNglg000VUk7ntwttavtXs0033zU00000000000007zz00000000000000Tzy00000000000000zzw000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000E0000000000000009U000000000000000G0000000000000000A0000000000000001w03l0A00000000000005a0M00000000000008Tank000000000FU0QPBgU000000000X00QnnzU000000000700BbbU0000000004C06P6Bg0000000000A07bANk000000000UQ000M00000000001Vs003U000000000040c00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000004"

		if ok:=FindText(55,798,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, (X+W)+200, Y+H//2
		  Click
		  Sleep 150
		}
	}
		
	
	
	; Open Allocation 의 Chk 전체 선택하게 하기
	Click_Chk_On_OpenAllocation(){
		
		

		Text:="|<Chk on Open Allocation of SO Manager>*152$17.zzy00000000000000000000000DU0lU11T22a4548+8MokTD00E00U00000000000000001zzw007zzk00000000000003m2AY4ED/UGR0Ym19qOGbYYk"
		if ok:=FindText(953,342,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)-5
			Sleep 150
			Click, 2
			MouseMove, A_ScreenWidth / 2, A_ScreenHeight / 2
			
			Sleep 150
		}

	}
	

	; Open Allocation 화면의 체크박스가 체크됐을때
	Che_is_Checked(){
		

		Text:="|<Checked>*137$9.0830oAn3kA4"
		if ok:=FindText(953,377,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Sleep 300
			return 1
		}

/*
		Text:="|<Chk on Open Allocation of SO Manager is checked>*145$20.U008002D88aG290wiE9+Y2H90YuNd+XmGM002000U008002000U008002000zzzzk0Tw07z05zk3Tw1bzElzqMTww7z61zk0Tw07zzzs"
		if ok:=FindText(1065,368,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  ;~ MouseMove, X+W//2, Y+H//2
		  Sleep 300
		  ;~ MsgBox, the che 체크됐음 1 리턴함 - Che_is_Checked()
		  return 1
		}
*/

		return 0		
	}
	
	
	; SO Manager 화면에서 Last Inv. Dt 날짜 찾아서 리턴하기
	getLastInvDateOnSOManager(){
		
		Clipboard := ""

		Text:="|<Last Inv. Date>*127$80.k000A0k0001zVg00030A0000MAP0000k3000061ak3sSS0nwUk1UPw1XAn0AtgM0M6P0Mm0k3AP6061ak0wsA0n6NU1UNg1v7n0Alak0M6P0MkCk3ANg061ak4A1g0n6C01UNg1b8n0AlXU0MAPzDHss3AMMk7y7U"
		if ok:=FindText(587,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H//2)+30
			Sleep 150
/*			
			Sleep 2500			
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 1000
*/

			; 날짜 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			lastInvDate := Clipboard
			Sleep 150
			
			return lastInvDate
		}
	}
	
	
	; SO Manager 화면에서 priority 번호 읽어서 리턴하기
	getPriorityOnSOManager(){
		
		Clipboard := ""

		Text:="|<Priority On SO Manager>*126$27.0k00M000000007qDXvsn6QS6Mn3km3MS6EP3km3MS6Mn3kn6MS6DX3U"
		if ok:=FindText(359,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H//2)+30
			Sleep 500
			;~ Sleep 2000
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 1000
			
			; Priority 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			priority# := Clipboard
			Sleep 150
			
			return priority#
			
		}
	}
	
	; SO Manager 에 있는 Open SO 에 보낼 아이템 있는지 확인하기
	checkOpenSoIfThereAreItemsShipOut(){
		
		Clipboard := ""

		Text:="|<Color on Open SO of SO Manager>*109$46.3s00M000zk01U0073U06000M700M00308DVVw5w01z6DsTk0CCNllX00kNa36A031aMAEk0A6NUl30AkNa3461n1aMAEQCCCNll0zkTlXy40y0y67kEU"
		if ok:=FindText(1364,159,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+20
			Sleep 150
			
			; 마우스 오른쪽 버튼 클릭 후 Filter 메뉴 위에서 엔터치기
			Send, {RButton}			
			Loop, 10
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			; 필터 입력창 나올때까지 기다리기
			WinWait, Filter and Sort - SO Manager : d_so_manager_sod
			Sleep 200
			
			; 필터 입력창에서 조건 입력하기
			Text:="|<Apply Filter of Filter and Sort Window>*147$81.000000000000A0k00A0DhY0003UC001U1UBU000S9syTDAABySS03lD6PBtVthaH00nPAnNhsABhzM06PTaPBj1Vhg300zSAnNgkABgqM0CRkrnta1VhnX01Vc0kM0k000000000630Q000000004"
			if ok:=FindText(380,551,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, (Y+H//2)+50
				Sleep 100
				Send, {LButton}
				SendInput, ( available_allocate_qty > 0 )
				Sleep 100
			  
				; ok 버튼 클릭하기
				Text:="|<OK on Open Allocation>*177$67.zzzzzzzzzzzk0000000000M0000000000A000000000060000000000300000000001U0000000000k0003snU000M0003aPU000A0001XjU00060001krk00030000sPs0001U000QBw0000k0006Cr0000M0003iNk000A0000yAs00060000000000300000000001U0000000000k0000000000M0000000000A0000000000600000000003zzzzzzzzzzzU"
				if ok:=FindText(1018,616,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, X+W//2, Y+H//2
					Click
					Sleep 150
				  
				  
					Clipboard := ""
					isThereItemsOnOpenSo = 1

					Text:="|<Color on Open SO of SO Manager>*109$46.3s00M000zk01U0073U06000M700M00308DVVw5w01z6DsTk0CCNllX00kNa36A031aMAEk0A6NUl30AkNa3461n1aMAEQCCCNll0zkTlXy40y0y67kEU"
					if ok:=FindText(1364,159,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, (Y+H)+20
						Sleep 150
						
						; 마우스 오른쪽 버튼 클릭 후 Filter 메뉴 위에서 엔터치기
						Send, {RButton}			
						Loop, 4
						{
							Sleep 100
							Send, {Down}
							;~ Sleep 150
						}
						Send, {Enter}
						Sleep 150
						
						isThereItemsOnOpenSo := Clipboard
						Sleep 150
						
						; 화면에 아이템이 없으면 1 리턴하기
						if(!isThereItemsOnOpenSo){
							return 1
						}
					}
				  
				  
				  
				}
			}

		}
		
		return 0

	} ; checkOpenSoIfThereAreItemsShipOut() 함수 끝
	
	
	
	; SO Manager 의 3사분변의 Pick Ticket 에 오픈된 주문이 있는지 확인 한 뒤 있으면 1 리턴하기
	checkPickTicketSectionToFindIfPendingOrderExists(){
		
		;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
		;~ MsgBox, 262144, Title, 펜딩 오더가 있는지 확인합니다
		
		Sleep 500
		Clipboard := ""
		pickDate = 1


		Text:="|<Pick Date on Pick Ticket of SO Manager>*134$44.SE41s0U4E10F0815CK4PrCFIp16YYrZ1UFb9x1EQ4OGEEIJ14YYo4t8SDgu"
		if ok:=FindText(626,695,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			;~ MouseMove, X+W//2, Y+H//2
			MouseMove, X+W//2, (Y+H)+12

			
			; Pick Date 밑에서 마우스 오른쪽 버튼 클릭 후 Pick Date 위에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			pickDate := Clipboard
			Sleep 150
			
;			MsgBox, % pickDate
			
			; pickDate 변수에 값이 없으면(펜딩 오더가 없으면) 1 리턴하기
			if(!pickDate){
;				MsgBox, 펜딩 오더가 없으니 1 리턴하기
				return 1 
			}			
			
		}
		
		; 펜딩 오더가 있으면 0 리턴
		return 0
		
		
	} ; checkPickTicketSectionToFindIfPendingOrderExists() 메소드 끝
	
	
	; SO Manager 에서 카드가 있는지 없는지 확인 후 있으면 1 리턴하기
	checkCC(){
		
		Clipboard := ""
		#ofCC = 0
		

		Text:="|<CC on Customer List of SO Manager>*127$22.D0DUk3X20A6y1U0U6020M081U0U6020M780kMU3X203sU"
		if ok:=FindText(524,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+20
			
			; CC 밑에서 마우스 오른쪽 버튼 클릭 후 Copy 위에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			#ofCC := Clipboard
			Sleep 150
			
;			MsgBox, 262144, Title, #ofCC : %#ofCC%
			
			; #ofCC 변수에 값이 있으면 참 값인 1 리턴하기
			if(#ofCC){
				return 1
			}
			
			
			
		}
		
		return 0
		
	} ; checkCC() 메소드 끝
	
	
	
	
	; Cust DC Rate에 할인이 있는 지 확인 후 있으면 Discount Rate 복붙하기
	copyDiscountRateAndPasteIfExist(){
		
		discountRate = 0
		
		Text:="|<Cust DC Rate>*183$64.D008T3kT081Y00VWN140U42Hr6B04PrSk9N8Mw0F99D0YsVXk1wQjo2Eu6B04aGkNdNcMaMFP9AwwwlwD16yrc"
		if ok:=FindText(410,202,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+15, Y+H//2
			Sleep 150
			
			; 마우스 오른쪽 버튼 클릭 후 값 얻기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			discountRate := Clipboard
			Sleep 150
			
			; 할인값이 있으면 Discount Rate 에 붙여넣기
			if(discountRate){

				; 사용자의 마우스 이동 허용
				BlockInput, MouseMoveOff
				
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 262144, DISCOUNT CUSTOMER, 할인율 발견. 티파니 고객인지 확인 후 진행. OK 누르면 복사-붙이기 진행됨
				
				; 사용자의 마우스 이동 막음
				BlockInput, MouseMove

				Text:="|<Discnt Rate>*182$55.y80023s10MU001140UAOwQxkXSvqBmHGEF99D6i898DXZzXFo4Y4aGklCOOG2/N9z5st9V6yrc"
				if ok:=FindText(515,202,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, (X+W)+15, Y+H//2
					Click
					Sleep 150
					Send, ^v{Enter}
					
					; 2초동안 할인율을 중복으로 입력할건지 묻는 메세지창 나오길 기다림
					WinWaitActive, Sales Order, , 2
					if ErrorLevel
					{
						; 2초 안에 메세지 창이 안 나오면 새로운 값이 들어간 것이 아니라는 뜻. 리턴으로 함수 빠져나감
						return
					}
					else{
						Send, {Enter} ; 메세지 창이 활성화 됐으면 엔터 침
						Sleep 150
						Click
						Sleep 150
						Send, ^s
						Sleep 3000
						Send, ^s
						Sleep 3000
					}
				}
			}
			; 할인값이 없으면 함수 빠져나가기
			else
				return
		}
	}
	
	
	
	
	;Sales Order 에서 Order Type 값 얻기
	getOrderType(){
		
		Clipboard := ""
		orderType := ""
		


		Text:="|<Order Type on Sales Order>*166$53.S0600z001a0A00M0024vtnUq/lo94oo1YoosGNjc3B9jkYnEE6CHEn8aaUAMaawFwt0Mlss000001200000006408"
		if ok:=FindText(735,187,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+30, Y+H//2
			Sleep 150
			
			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			orderType := Clipboard
			Sleep 150

;			MsgBox, % CustCode
			
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
			if(!orderType){
			;~ if(CustCode == ""){
;				MsgBox, no value in variable, restart the method
				CommN41.getOrderType()
			}
			
			
			return orderType
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.getOrderType()
		}		
		
	} ; getOrderType 메소드 끝
	
	
	
	
	; PMT Method 를 FG-CC 나 LAS-CC 로 바꾸기
	changePMTMethodToFGorLAS(CustomerPO){
		
;	MsgBox, 262144, Title, CustomerPO : %CustomerPO%




		Text:="|<PMT Method>*175$57.wnTVa0G00QKMkAk2E03Wn61aCvnnwKMkAmOGGHwx61uTGGSQ5ck/G2GHnUh61OHGGGQ5ck/FnGSTU"
		if ok:=FindText(733,248,150000,150000,0,0,Text)
		{
			Sleep 1000
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			X_position = (X+W)+20
			Y_position = Y+H//2
			
			
			; PM Method 의 상태를 바꾸기전에 먼저 읽고 바꿀필요 없으면 나가기
			Loop{
				PMTMethod = ""
				
				MouseMove, (X+W)+20, Y+H//2
				
				; PM Method 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
				Send, {RButton}
				Loop, 4
				{
					Sleep 150
					Send, {Down}
					Sleep 200
				}
				Send, {Enter}
				Sleep 200
					
				PMTMethod := Clipboard
				Sleep 1000
				
				; 값을 읽었으면 반복 끝내기
				if(PMTMethod)
					break				
			}

			if CustomerPO contains MTR
			{
				if(PMTMethod == "FG-CC"){
;					MsgBox, 262144, Title, 바꿀필요 없어서 나감
					return
				}
			}
			else if CustomerPO contains OP
			{
				if(PMTMethod == "LAS-CC"){
;					MsgBox, 262144, Title, 바꿀필요 없어서 나감
					return
				}
			}
			else if CustomerPO contains Credit_Card
			{
				if(PMTMethod == "CREDIT CARD"){
;					MsgBox, 262144, Title, 바꿀필요 없어서 나감
					return
				}
			}
			
			
			; PMT 값이 원하는 값이 아니면 바꾸기
			MouseMove, (X+W)+20, Y+H//2 ; 마우스 PMT Method 로 이동시키기
			Sleep 150
			Click
			Sleep 150
				
			
			if CustomerPO contains MTR
			{
				Send, FG-CC
				Sleep 150
				Send, {Tab}			
				Sleep 150
			}
			else if CustomerPO contains OP
			{
				Send, LAS-CC
				Sleep 150
				Send, {Tab}
				Sleep 150			
			}
			else if CustomerPO contains Credit_Card
			{
				; Sale Order 에 있는 Customer PO 번호 읽어서 아무런 값도 없으면 쇼오더라는 뜻이고 결제방법을 신용카드로 변경할 필요 없으니 리턴해서 메소드 끝내기
				JustReadCustomerPO := CommN41.GetCustPONumber()
				if(!JustReadCustomerPO)
					return
				
				; 마우스 위치를 PMT Method 로 다시 옮기기
				MouseMove, (X+W)+20, Y+H//2
				Sleep 150
				Click
				Sleep 200
				Send, CREDIT ; CREDIT 을 쳐서 PMT Method 를 Credit Card로 바꾸기
				Sleep 200
				Send, {Tab}
				Sleep 200			
			}			
			
			Sleep 500
			
			; 저장하기
			send, ^s
			Sleep 150
			CommN41.ClickSave()
			Sleep 700
			IfWinActive, Sales Order
			{
				Send, {Enter}
				Sleep 300
			}
			Sleep 2000


			; PM Method 의 상태값 읽기
			Loop{
				PMTMethod = ""
				
				MouseMove, (X+W)+20, Y+H//2	; 마우스 위치 다시 이동하기
				
				; PM Method 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
				Send, {RButton}
				Loop, 4
				{
					Sleep 150
					Send, {Down}
					Sleep 200
				}
				Send, {Enter}
				Sleep 200
					
				PMTMethod := Clipboard
				Sleep 1000
				
				; 값을 읽었으면 반복 끝내기
				if(PMTMethod)
					break
				
			}
			
			
			; PM Method 값이 원하는 것으로 제대로 입력됐는지 확인하기
			if CustomerPO not in Credit_Card	; 신용카드 결제하는 Pick Ticket이 아닐때
			{
				if(PMTMethod != "FG-CC")		; FG 도 아니고
				{
					if(PMTMethod != "LAS-CC")	; LAS 도 아니면						
					{
						; 신용카드 결제하는것도 아닌데 FG-CC 도 아니고 LAS-CC 도 아니면 제대로 바뀐것이 아니니 재귀호출로 다시 시작
;	MsgBox, 262144, Title, FG-CC 혹은 LAS-CC로 바뀌지 않았음. 재귀호출로 바꾸기 다시 시작함.
						Sleep 2000
						CommN41.changePMTMethodToFGorLAS(CustomerPO)
					}
				}
			}
			else if CustomerPO in Credit_Card	; 신용카드 결제하는 Pick Ticket 일때
			{
				if(PMTMethod != "CREDIT CARD")	; PMTMethod 이 CREDIT CARD 로 바뀌지 않았으면 제대로 바뀐것이 아니므로 재귀호출로 다시 시작
				{
;	MsgBox, 262144, Title, Credit Card로 바뀌지 않았음. 재귀호출로 바꾸기 다시 시작함.
					Sleep 2000
					CommN41.changePMTMethodToFGorLAS(CustomerPO)
				}
			}

			
		} ; END끝 - if ok:=FindText(733,248,150000,150000,0,0,Text)
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 300
			CommN41.changePMTMethodToFGorLAS(CustomerPO)
		}

		; 저장하기
		;~ send, ^s
		;~ Sleep 700
		;~ IfWinActive, Sales Order
		;~ {
			;~ Send, {Enter}
			;~ Sleep 300
		;~ }
		;~ CommN41.ClickSave()
		;~ Sleep 500
		
;		MsgBox, 262144, 무제, 다 끝났음 리턴하기 직전

		return
		
	} ; END끝 - changePMTMethodToFGorLAS(CustomerPO)
	
	
	
	
	
	; SO Manager 화면의 왼쪽 밑 Pick Ticket 섹션에서 Pick Date 날짜 가져오기
	getPickDateOnPickTicketSectionOfSOManager(){
		
		
;		MsgBox, 262144, Title, 픽티켓 섹션에서 픽티켓 날짜 가져옵니다.
		
		i = 1
		Clipboard := ""
		pickDate := ""
		
		

		Text:="|<Pick Date>*122$43.wU83k10F04140U8dmkWStoJBEFd9/mUk8nYz1EQ4GGEUce2999EHYVswnc"
		if ok:=FindText(626,695,150000,150000,0,0,Text)
		{
			
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+13
			Sleep 150

			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			pickDate := Clipboard
			Sleep 150

;			MsgBox, % pickDate
/*			
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
			if(!pickDate){
				
				if(i == 3){
					CommN41.getPickDateOnPickTicketSectionOfSOManager()
					i++					
					
				}

			}
*/			
			
			return pickDate
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.getPickDateOnPickTicketSectionOfSOManager()
		}						
		
		
		
		
	} ; getPickDateOnPickTicketSectionOfSOManager 메소드 끝
	
	
	
	; Sales Order 탭에서 고객의 할인율 찾아서 할인받았으면 1 리턴하기
	FindCustDCRate(){
		
		Clipboard := ""

		Text:="|<Cust DC Rate>*177$67.zzzzzzzzzzzk0000000000M0000000000A00000000006D008D3kT083AU044H88U41Y2Hr2B04PrCm19916U2999t0YsUXE1wQbwUG7EFc0YGG6NdNc8aMF99D7bba7Vs8rqRU0000000000k0000000000Tzzzzzzzzzzw"
		if ok:=FindText(409,201,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			;~ MouseMove, X+W//2, Y+H//2
			MouseMove, (X+W)+10, Y+H//2
			Sleep 150
			
			
			; 고객 코드명 위에서 할인율 얻기 위해 마우스 오른쪽 버튼 클릭 후 밑으로 4칸 내려서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			discountRate := Clipboard
			Sleep 150

			; 할인 받았으면 1 리턴하기
			if(discountRate){
;				MsgBox, % discountRate
				return 1
			}
			; 할인 안 받았으면 0 리턴하기
			else
				return 0			
			
		}

		
		
	} ; FindCustDCRate 메소드 끝
	
	
	
	
	
	; pick ticket 화면에서 정보 읽은 뒤 값들 리턴하기
	getInfoOnPickTicket_Then_ReturnThem(){
		
		Sleep 2000
		
		
		WinActivate, ahk_class FNWND3126		
		
		Clipboard := ""
		
		
		; pick # 찾기
		Text:="|<pick# on Pick Ticket>*174$30.000070000700007wU81TW081LWb/7zWdi3LwcA2rUcC7zUce2bUb92bU"
		if ok:=FindText(276,154,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+5, Y+H//2
			Sleep 150
			
			; pick #  위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			pick# := Clipboard
			Sleep 150

;			MsgBox, % pick#

;~ /*
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
;			if(!pick#){
;				MsgBox, 262144, Title, pick# 변수에 값이 없음. 다시 시작
;				CommN41.getInfoOnPickTicket_Then_ReturnThem()
;			}			
*/			
			
			loop{ ; 첫번째
				
				
				Clipboard := ""
				
				; Customer Code 얻기
				Text:="|<Customer on Pick Ticket>*174$48.zzzzzzzz000000070000000700000007S00E0007m00E0007UGSvnyQzUGGGGGabUGQGOHybUG7GOHUbnGnGGGabSSSPmGQb000000070000000700000007zzzzzzzzU"
				if ok:=FindText(267,203,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, (X+W)+5, Y+H//2
					Sleep 150
					
					; pick #  위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
					Send, {RButton}
					Loop, 4
					{
						Sleep 150
						Send, {Down}
						;~ Sleep 150
					}
					Send, {Enter}
					Sleep 150
					
					CustCode := Clipboard
					Sleep 150

;					MsgBox, % CustCode

;~ /*
					; 만약 변수에 값이 안 들어갔으면 루프 처음으로 돌아가서 다시 시작하기
;					if(!CustCode){
;						MsgBox, 262144, Title, CustCode 변수에 값이 없음. 다시 시작
;						CommN41.getInfoOnPickTicket_Then_ReturnThem()
;					}
*/					
					
					loop{ ; 두 번째
						
						Clipboard := ""
							
						; 작성한 날짜, 시간 얻기
						Text:="|<Update Date on Pick Ticket>*163$45.W033k10QE0MF083WST2BvbQGOMFd9DWHn2AtDwGSMFd93aHH2999vXnsSDgv0E00000M2000003U"
						if ok:=FindText(1658,200,150000,150000,0,0,Text)
						{
							CoordMode, Mouse
							X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
							MouseMove, (X+W)+5, Y+H//2
							Sleep 150
							
							; pick #  위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
							Send, {RButton}
							Loop, 4
							{
								Sleep 150
								Send, {Down}
								;~ Sleep 150
							}
							Send, {Enter}
							Sleep 150
							
							updDate := Clipboard
							Sleep 150

;							MsgBox, % updDate

;~ /*
							; 만약 변수에 값이 안 들어갔으면 루프 처음으로 돌아가서 다시 시작하기
;							if(!updDate){
;								MsgBox, 262144, Title, updDate 변수에 값이 없음. 다시 시작
;								CommN41.getInfoOnPickTicket_Then_ReturnThem()								
;							}
*/							
							
							; 변수들 중 한 개에라도 값이 아무것도 없으면 다시 시작
;							if(!pick# && !CustCode && !updDate){
;							MsgBox, 262144, Title, 변수들 중 값이 없는 변수가 있음. 다시 시작
;							CommN41.getInfoOnPickTicket_Then_ReturnThem()							
;							}
							
;		MsgBox, 262144, Title, 찾은값들 리턴하기 전에 확인해보기`n`n`n%pick#%`n`n`n%custCode%`n`n`n%updDate%
							
							return [pick#, CustCode, updDate]
							
							
							
						} ; if ends - 작성한 날짜, 시간 얻기						
				
						; 만약 그림 못찾았으면 재귀호출해서 다시 처음부터 시작하기
						;~ else if(!updDate){
;						else{
;							MsgBox, 262144, Title, updDate 그림 못 찾았음
;							CommN41.getInfoOnPickTicket_Then_ReturnThem()
;						}
						
						
					} ; loop ends - 두 번째

				} ; if ends - Customer Code 얻기
		
				; 만약 그림 못찾았으면 재귀호출해서 다시 처음부터 시작하기
				;~ else if(!CustCode){
;				else{
;					MsgBox, 262144, Title, CustCode 그림 못 찾았음
;					CommN41.getInfoOnPickTicket_Then_ReturnThem()
;				}			
							
				
				

			} ; loop ends - 첫번째

		} ; if ends - pick # 찾기
		
		; 만약 그림 못찾았으면 재귀호출해서 다시 처음부터 시작하기		
		;~ else if(!pick#){
		else{
			MsgBox, 262144, Title, pick# 그림 못 찾았음
			CommN41.getInfoOnPickTicket_Then_ReturnThem()
		}			
			
		
		; 중간에 리턴되지 않고 이 코드가 실행되면 뭔가 이상한것. 무한반복 되려나?
		CommN41.getInfoOnPickTicket_Then_ReturnThem()
		
	} ; getInfoOnPickTicket_Then_ReturnThem 메소드 끝
	


	; SO Manager 화면에서 주소를 읽어서 리턴
	getADDr(){
		
		Clipboard := ""
		addr := ""
		
		WinActivate, ahk_class FNWND3126
		

		Text:="|<Address 1 on SO Manager>*117$71.3U1UA00000027030M000000AP060k000000sq3wTbnsS7U3l4AtbCAN6FU1aAEm6MMm0U03ANVgAkUa1U06zv3MNXz7Vs0BUq6kn201UM0O1YAVa6410E0w1gNXAAN6FU1s3DlyMDXsy03U"
		if ok:=FindText(411,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+20
			Sleep 150
							
			; 주소 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
							
			addr := Clipboard
			Sleep 150
			
			; 주소 못 읽었으면 재귀호출
;			if(!addr){
;				CommN41.getADDr()
;			}
			
			return addr
			
		}
				
		
		
		
	} ; getADDr() 메소드 끝
	
	
	; Pick Ticket 버튼 클릭하는 메소드
	clickPickTicketButton(){		

		Text:="|<Pick_Ticket_Button>*133$68.s0000000000200000000000U0000U00400DzwSE87y10128V4E2080E0EW8F5DY29wbDDzwFI+0WVG928W7Z308cMyEW8V1Es2+784DzsEI/0WVO10U044yM8bnSQzzk00000000C0s000000003UC000000002"
		if ok:=FindText(203,57,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 500
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.clickPickTicketButton()
		
		return
		
	} ; clickPickTicketButton() 메소드 끝
	
	
	
	
	; Customer Master 에 있는 Priority 클릭하기
	clickPriority(){
		
		

		Text:="|<Priority_on_CustomerMaster>*150$39.y0U04E6M00060nSXXpySP4mMaGyMYP4mK34XMaSkMaH4la34QMXA000001U00000MU"
		if ok:=FindText(491,314,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+35, Y+H//2
			
			Click
			
			Sleep 500
		}

		return
		
	} ; clickPriority() 메소드 끝
	
	
	; Pick Ticket 화면에 있는 Customer PO 값 가져오기
	getCustPONumberOnPickTicketScreen(){	
		
		Clipboard := ""
		CustomerPO# := ""

		Text:="|<Cust_PO#>*179$49.zzzzzzzzU0000001k0000000s0000000Rs011sD2zY00UWAlLUGSsF4/zk9988W4ps4b47V2Kw2Eu20VTzB/B10NZCwwwkV7Wb00000003U0000001k0000000zzzzzzzzy"
		if ok:=FindText(612,186,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+15, Y+H//2
			Sleep 150
			
			; Customer PO 값 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
							
			CustomerPO# := Clipboard
			Sleep 700
			
			; Customer PO 값 못 읽었으면 재귀호출
			if(!CustomerPO#){
				
				; 이미 2번 반복했으면 반복 끝내고 그냥 리턴하기
				if(counter_forN41 >= 2)
					return CustomerPO#
				
;				MsgBox, 262144, Title, 값 못 읽었음. 재귀호출로 다시 시작함. %counter_forN41% 번째 반복하는 것
				counter_forN41++
				CommN41.getCustPONumberOnPickTicketScreen()
			}
			
			return CustomerPO#
			
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.getCustPONumberOnPickTicketScreen()
		
		return
		
	} ; getCustPO#OnPickTicketScreen() 메소드 끝
	
	
	
	
	; YES 클릭하기
	clickYES(){	
		
		MsgBox, YES 클릭하기

		Text:="|<YES>*184$15.m06k0StnhcBxVg6BUNbCU"
		if ok:=FindText(1031,665,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Sleep 150
			Click
			Sleep 150
		}
		
		return	
		
	} ; clickYES() 메소드 끝
	
	
	
	; 체크박스 선택되지 않았다는 에러 메세지 창 문구
	noSelectedCheckboxes_ErrorMessage(){

		Text:="|<NO_SELECTED_CHECKBOXES>*141$199.Q0U08042000U0000Y003U000F003Vs4U8k0E04021000E0000E00100008U0216204M78szVvkXbXS3Xss9DXVnb7wDT711118re4YYl04UGEG439YY4aGEFaHG2AYUEUUYV4nmSEUCE9stW14WS2G9s8W9914Hk6EEGMW9188E984UY90WF8194U4F4YUW9018892E4ka684Y2MG4UN8a0YWM28mGEF4k0WA4V8yCFlr3nV7DCQ74FkGF714C98CWC3kw2HbM"
		if ok:=FindText(978,537,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  ;~ MouseMove, X+W//2, Y+H//2
		  
		  ; 에러 메세지 창이 떴으면 1 반환
		  return 1
		}
		
		return 0
		
	} ; noSelectedCheckboxes_ErrorMessage() 메소드 끝
	




	
	; 픽티켓 마지막 메세지(= 결제와 프린팅 준비 됐다는 메세지) 떴는지 확인. 
	lastPTMessage_ReadyToGoCCAndPrinting(){
		
		MouseMove, 100, 100 ; 마우스 포인터가 찾으려는 곳에 겹쳐져서 못 찾는 것을 방지하기 위해 일단 임의의 위치로 마우스 포인터 옮김
		
		Sleep 500
		
;MsgBox, 262144, Title, 마지막 메세지 창 떴는지 확인하는 메소드 lastPTMessage_ReadyToGoCCAndPrinting 들어옴
;Sleep 100

		Text:="|<Last Message Window(Totla Pick Ticket Made)>*147$117.y0E11t0E7m0U0U00080102088U20804040001008Qxt157G12CYRsTnlss96G188dWU8H54Y3G2N918WEt1t8M12EkwUGFm9s14G98893U8G7442GGF808mF911AK12MgkUGGH9U93Xj888uM8Fonb2GSD71U"
		if ok:=FindText(2968,410,150000,150000,0,0,Text)
		{

;MsgBox, 262144, Title, 마지막 메세지 창 찾았음
;Sleep 500

			; 마지막 메세지 창의 OK 버튼 찾아서 누르기
			Text:="|<OK>*157$14.D4wN+2IUq89m2IlYblC"
			if ok:=FindText(2756,453,150000,150000,0,0,Text)
			{
;MsgBox, 262144, Title, 마지막 메세지 창 찾았으니 OK 버튼 클릭 후 1을 리턴하면서 lastPTMessage_ReadyToGoCCAndPrinting 메소드 빠져나감
;Sleep 100
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Sleep 100
				Click
				
;				Sleep 2000 ; 이렇게 OK 버튼 클릭 뒤 쉬어주지 않으면 그 찰나의 시간에 OK 버튼을 또 찾아서 코드가 꼬여버리는것 같다.
				
				return 1
			}
			
			;~ return 1
		}
		
;MsgBox, 262144, Title, 마지막 메세지 없음

		return 0
	} ; lastPTMessage_ReadyToGoCCAndPrinting 메소드 끝
	
	
	
	
	
	; OK 버튼 클릭하기
	click_OK_Button(){
		
		MouseMove, 100, 100 ; 마우스 포인터가 찾으려는 곳에 겹쳐져서 못 찾는 것을 방지하기 위해 일단 임의의 위치로 마우스 포인터 옮김
		
		Text:="|<OK>*157$14.D4wN+2IUq89m2IlYblC"
		if ok:=FindText(2756,453,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			
;MsgBox, 262144, Title, 이 메세지 창에 있는 OK 버튼 찾았음
;Sleep 200

;MsgBox, 262144, Title, OK 버튼 누르기 전에 이 창이 픽티켓 마지막 메세지 창인지 확인하기 위해 메소드 호출할 예정
			isThereTheWindowExist := CommN41.lastPTMessage_ReadyToGoCCAndPrinting()
;MsgBox, 262144, Title, isThereTheWindowExist : %isThereTheWindowExist%`n`n1 : 이 창은 마지막 창`n0 : 이 창은 마지막 창이 아님
			if(isThereTheWindowExist == 1){
			; 이 창이 마지막 픽티켓 확인창이니 클릭하지 않고 그냥 메소드 나감
;MsgBox, 262144, Title, OK버튼 찾았으나 이 창이 마지막 픽티켓 확인창이니 클릭하지 않고 그냥 메소드 click_OK_Button() 나감. 여기서 클릭해버리면 이 창이 마지막 픽티켓 창인지 그냥 일반 메세지 창인지 구분할 수 있는 방법이 없기 때문
;Sleep 100
				return 1
			}

			; 이 창이 마지막 픽티켓 확인창이 아니기 때문에 화면에 나타난 OK 버튼 그냥 클릭함.
			MouseMove, X+W//2, Y+H//2
			Sleep 100
			Click
		}
		
		return		
	} ; click_OK_Button 메소드 끝
	
	
	; Yes 버튼 클릭하기
	click_Yes_Button(){
		
		MouseMove, 100, 100 ; 마우스 포인터가 찾으려는 곳에 겹쳐져서 못 찾는 것을 방지하기 위해 일단 임의의 위치로 마우스 포인터 옮김
		
		Text:="|<Yes>*152$14.m05U1HbRd3TMY39UGCS"
		if ok:=FindText(2621,492,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
;MsgBox, 262144, Title, Yes 버튼 누를 예정
;Sleep 200
			MouseMove, X+W//2, Y+H//2
			Sleep 100
			Click
		}
		
		return		
	} ; click_Yes_Button 메소드 끝
	

	
	
	; Style 클릭하기
	click_Style(){
		
		Sleep 100		

		Text:="|<Style Button>*148$40.08000009U00000Y000002k00040DU3m0E000F8100E11yQsNU729IF6038hT0Q04XZ0Vk4G6K270SCFDES00101Vs00A081E0000U"
		if ok:=FindText(467,56,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Sleep 200
		  Click
		  Sleep 200
		}
		; Style 버튼 못 찾았으면 재귀호출해서 계속 찾기
		else
			CommN41.click_Style()
		
		
		return
	} ; click_Style() 끝
	
	
	; Style 탭 안에 있는 Style 클릭하기
	; 엄밀히 말하자면 연필 그림 클릭 Style 글자가 클릭 했을때는 두꺼워지고 클릭 안 했을때는 연해져서 상황에 따라 못찾는 경우가 생겨서
	click_StyleButtonInTheStyleTab(){
		
		Sleep 100
		
		

		Text:="|<s in s>*160$13.0707k1s4Q4Y4U4Y4Y4Y4Y2Y140w0E"
		if ok:=FindText(304,153,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Sleep 200
		  Click
		  Sleep 1000
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.click_StyleButtonInTheStyleTab()
		
		
		return
	} ; click_StyleButtonInTheStyleTab() 끝
	
	
	
	
	
	; 스타일 입력칸 찾아서 클릭하기
	findTheSearchBoxOfStyle(){
		
		Sleep 100
		

		Text:="|<The Search Box of Style>*182$26.VryTvNzbyw39VXamHQNZY3mMNTAbCHMQnaDzwzzzyTzs"
		if ok:=FindText(806,156,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, (X+W)+10, Y+H//2
		  Sleep 200
		  Click
		  Sleep 500
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.findTheSearchBoxOfStyle()
		
		
		return
	}
	
	
	
	; 색깔 입력칸 찾아서 클릭하기
	findTheSearchBoxOfColor(){
		
		Sleep 100
		

		Text:="|<color>*182$29.VztzyRznzxzVa63yN9YrxmL9bvYiHCnNBb1kn3C"
		if ok:=FindText(947,155,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, (X+W)+15, Y+H//2
		  Sleep 200
		  Click
		  Sleep 500
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.findTheSearchBoxOfColor()
		
		
		
		return
	}
	
	
	
	
	; 화면에 제대로 들어왔는 확인키 위해 화면에서 스타일 번호 읽기
	getStyleNumberOnStyleTab(){
		

		Text:="|<Style# on the screen>*147$35.S8100IYk200d3woQ7vX9944XaGHs91gwY0znMlAUcwNWC1E0300000A0008"
		if ok:=FindText(280,220,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+50, Y+H//2
			Sleep 150
			
			; Customer PO 값 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
							
			Style# := Clipboard
			Sleep 150
			
			; Style# 값 못 읽었으면 재귀호출
			if(!Style#){
				CommN41.getStyleNumberOnStyleTab()
			}
			
			return Style#			
		}
		
		return Style#
		
	} ; getStyleNumber 끝
	



	; 화면에 제대로 들어왔는 확인키 위해 화면에서 색깔 읽기
	getColorOnStyleTab(){

		Text:="|<color on Style tab>*145$29.S0401W08020QFlw1YaH82B8qE4OFglAYmMwC8sm"
		if ok:=FindText(277,241,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+50, Y+H//2
			Sleep 150
			
			; Customer PO 값 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
							
			OOS_Color := Clipboard
			Sleep 150
			
			; Color 값 못 읽었으면 재귀호출
			if(!OOS_Color){
				CommN41.getColorOnStyleTab()
			}
			
			return OOS_Color
		}
		
		return OOS_Color
		
	} ; getStyleNumber 끝







	; On Order 의 돋보기 버튼 눌러서 창 열기
	clickMagnifierNextTo_OnOrder(){
		

		Text:="|<Magnifier Next to On Order>*159$72.0y00000000003z000000000071U0w00w020060k1a01a020060k12T12SSCD60k33P33MqFA61U33N33NWzA73U12N12NWkADz01aN1aMqPAM000wN0wMSCAk00000000000U00000000000U"
		if ok:=FindText(272,585,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			;~ MouseMove, X+W//2, Y+H//2
			MouseMove, X+4, Y+H//2
			Sleep 150
			Click
			Sleep 150
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.clickMagnifierNextTo_OnOrder()		
		
		return
		
	} ; clickMagnifierNextTo_OnOrder() 끝
	
	
	
	
	; 적당한 창 위치에서 (이 경우에는 Close 버튼을 찾은 뒤 그 위에 마우스를 위치했음) 마우스 오른쪽 버튼 클릭 후 Excel 파일 만들기 메뉴에서 엔터치기
	activeDownloadExcelFileFunction(){

		Text:="|<Close Button>*181$27.000000000000000000000000000000000000000000000zzzzs0000000000000000000000000000000D8003B000E/nnq1HmHkCPbu1n7kNeSPNtSSS00000000000000000000000000000007zzzz00000000000000000000000000000004"
		if ok:=FindText(780,658,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)-50
			Sleep 150
			
			; 적당한 창 위치에서 (이 경우에는 Close 버튼을 찾은 뒤 그 위에 마우스를 위치했음) 마우스 오른쪽 버튼 클릭 후 Excel 파일 만들기 메뉴에서 엔터치기
			Send, {RButton}
			Loop, 7
			{
				Sleep 150
				Send, {Down}
			}
			Send, {Enter}
			Sleep 150
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.activeDownloadExcelFileFunction()

		return
	}


	; 엑셀파일을 바탕화면에 저장하기
	makeTheExcelFileOfTheStyleOnDesktopScreen(OOSStyle#, OOSColor){

		; 엑셀 저장 창의 Desktop 버튼 클릭하기
		Text:="|<desktop Button>*165$42.0zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00jzzxI00zzzzw00000000000000000000000000000000000000000000000000000000000000000000000y00E000X00EE00V00EE00VbDHwwyVccKFYmVDgQF6XV87QF6XXA3KFYmw7iHQsy000000U000000U000000UU"
		if ok:=FindText(54,198,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 150
			
			; File name 으로 넘어가서 해당스타일번호-색깔 형식으로 파일명을 만들어서 (바탕화면에) 저장 후 Done! 윈도우 창이 나오면 엔터치기
			Send, !n
			Sleep 100
			StringReplace, OOSColor, OOSColor, /, , All
			Send, %OOSStyle#%-%OOSColor%
			Sleep 100
			Send, {Enter}
			Sleep 100
			
			WinWaitActive, Done!			
			IfWinActive, Done!
				Send, {Enter}
			Sleep 400
			
			; 남아있는 On Order List 윈도우 닫기
			Text:="|<Close Button>*181$27.000000000000000000000000000000000000000000000zzzzs0000000000000000000000000000000D8003B000E/nnq1HmHkCPbu1n7kNeSPNtSSS00000000000000000000000000000007zzzz00000000000000000000000000000004"
			if ok:=FindText(780,658,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Click
				Sleep 150
			}
			else
				MsgBox, 262144, Title, PLEASE CLOSE 'On Order List' WINDOW ON THE SCREEN.
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.makeTheExcelFileOfTheStyleOnDesktopScreen(OOSStyle#, OOSColor)				
		
		
		return
	}
	
	
	
	
	
	
	
	; Style탭에 있는 Style So Detail 버튼 누르기
	clickStyleSODetailButton(){

		Text:="|<Style SO Detail>*184$15.zzw01bzAU9Z9AktUNAktUVAkNU5AztU8Qz7U1w0Tzzw"
		if ok:=FindText(888,128,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 150
		}
		
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.clickStyleSODetailButton()
		
		return
	}







	; Sales Order Detail by Customer 새창에서 Color 입력칸 찾아서 클릭하기
	clickTheBlankOfColorOfSalesOrderDetailListbyCustomer(){


		Text:="|<blank of Color on Sales Order Detain List by Custoer>*210$41.zzzzzzzzzzzzzw00000DU00000TE00000yUS0E01x1a0U03u21tSQ7oA2OakDcMArB0TEENiO0yUnHIo1x0wwj83u000007o00000Dc00000TE00000yU00001w000003s000007nzzzzzzw"
		if ok:=FindText(500,66,150000,150000,0,0,Text)
		{
			Sleep 500
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			;~ MouseMove, X+W//2, Y+H//2
			MouseMove, X+W, Y+H//2
			Sleep 150
			Click
			Sleep 150
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.clickTheBlankOfColorOfSalesOrderDetailListbyCustomer()
		
		return
	}
	
	
	
	
	
	
	
	; 돋보기 아이콘 찾아서 클릭하기
	clickTheIconOfMagnifying_glass(){		

		Text:="|<magnifying glass>*190$12.1s2A42838383834264DsQ0w0s0U"
		if ok:=FindText(27,40,150000,150000,0,0,Text)
		{
			Sleep 300
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 150
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.clickTheIconOfMagnifying_glass()				
		
		
		return
	}


	
	
	
	
	
	; 적당한 창 위치에서 (이 경우에는 손가락 화살표를 찾은 뒤 그 위에 마우스를 위치했음) 마우스 오른쪽 버튼 클릭 후 Excel 파일 만들기 메뉴에서 엔터치기
	findFingerArrow_ThenActiveDownloadExcelFileFunction(){
		

		Text:="|<Finger arrow>*198$23.000U07n00ku027zU801zk0P20Dw60R0A1w0M1s0k3U1UDU3UD0bzw3zzs7rzU8"
		if ok:=FindText(15,127,150000,150000,0,0,Text)
		{
			Sleep 300
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2

			Sleep 150
			
			; 적당한 창 위치에서 (이 경우에는 Close 버튼을 찾은 뒤 그 위에 마우스를 위치했음) 마우스 오른쪽 버튼 클릭 후 Excel 파일 만들기 메뉴에서 엔터치기
			Send, {RButton}
			Loop, 7
			{
				Sleep 150
				Send, {Down}
			}
			Send, {Enter}
			Sleep 150
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.findFingerArrow_ThenActiveDownloadExcelFileFunction()

		return
	}






	; Sales Order Detail 창에서 엑셀파일을 만든 뒤 바탕화면에 저장하기
	makeExcelFileOfDropAllSelectedStyle(OOSStyle#){

		; 엑셀 저장 창의 Desktop 버튼 클릭하기
		Text:="|<desktop Button>*165$42.0zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00zzzzw00jzzxI00zzzzw00000000000000000000000000000000000000000000000000000000000000000000000y00E000X00EE00V00EE00VbDHwwyVccKFYmVDgQF6XV87QF6XXA3KFYmw7iHQsy000000U000000U000000UU"
		if ok:=FindText(54,198,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 300
			
			; File name 으로 넘어가서 해당스타일번호-색깔 형식으로 파일명을 만들어서 (바탕화면에) 저장 후 Done! 윈도우 창이 나오면 엔터치기
			Send, !n
			Sleep 100
			;~ Send, %OOSStyle#%-%OOSColor%
			Send, %OOSStyle#%-All
			Sleep 100
			Send, {Enter}
			Sleep 100

			; Done 창 닫기
			WinWaitActive, Done!			
			IfWinActive, Done!
				Send, {Enter}
			Sleep 500
			
			
			; Sales Order Detail List by Customer 창 닫기
			WinClose, Sales Order Detail List by Customer			

		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.makeExcelFileOfDropAllSelectedStyle(OOSStyle#)
		
		
		return
	}
	
	
	
	



	; Pick Ticket 탭에 있는 Sales Order 로 이동하기
	moveToSONumTab(){

		Text:="|<SO# on Pick Ticket Tab>*179$29.zzzzy0000Q0000s0001rUS0jl1a1LU24Dzk48BQM8EKsMEVzskn1Hz8w2b0000C0000Q0000zzzzzzzzzzzzzzzU"
		if ok:=FindText(276,187,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			;~ MouseMove, X+W//2, Y+H//2
			MouseMove, (X+W)+20, Y+H//2

		
			; SO# 번호로 이동하기
			Send, {RButton}
			Loop, 9
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.moveToSONumTab()
			
		return
		
	} ; END끝 - moveToSO#Tab()
	
	
	
	; pre authorized 버튼 클릭
	findAndClickPreAuthorizeCreditCardIcon(){
		
		
		Text:="|<pre-authorize Button>*205$16.001zzbzyTztzzY0SE1tzzbzyTztzzc01zzy"
		if ok:=FindText(718,129,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
				
			Sleep 2000
		}
		; 못 찾았으면 재귀호출로 계속 찾기
		else
			CommN41.findAndClickPreAuthorizeCreditCardIcon()
		
		return
		
	} ; end끝 - findAndClickPreAuthorizeCreditCardIcon()
	




	; Pick Ticket 에 있는 리프레쉬 버튼 클릭하기
	ClickRefreshButtonOnPickTicket(){
		

		Text:="|<refresh button on Pick Ticket Tab>*164$12.103070Ds7w3613U1U1kUMkDs7w0s0k0UU"
		if ok:=FindText(237,129,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 300
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.ClickRefreshButtonOnPickTicket()
			
		}
	} ; end - ClickRefreshButtonOnPickTicket()
	
	
	
	
	
	; SO Manager에 있는 Allocate 버튼 클릭하기
	clickAllocateButtonOfSOManager(){
		
;~ MsgBox, clickAllocateButtonOfSOManager 메소드 들어왔음

		Text:="|<Allocate Button on SO Manager>*154$39.000000U0000087y0002/zw000WTzk0Tsbw00428zzk0jW33y05zkUDk0iu40y05nEU000zy40005zE0000ju000040E0000zy4"
		if ok:=FindText(394,334,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+5, Y+H//2
			Sleep 150
			Click
			Sleep 150
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.clickAllocateButtonOfSOManager()			
		}
		return
	} ; end - clickAllocateButtonOfSOManager()
	
	
	
	
	; Check All 버튼 클릭하기
	clickCheckAllButtonOnTheAllocatinScreen(){
		

		Text:="|<Check All Button of Allocate by SO Detail>*127$50.SA003032QH000k0ka0yC7BUS9UAYGPk5WM39wUw38a0mE8D0z9lAaGPM8mLX8sQn66a"
		if ok:=FindText(335,35,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Sleep 150
			Click
			Sleep 1000
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.clickCheckAllButtonOnTheAllocatinScreen()			
		}
		
		return
		
	} ; end - clickCheckAllButtonOnTheAllocatinScreen
	
	
	
	; 체크박스 체크됐는지 확인
	isTheCheckBoxChecked(){
		

		Text:="|<check box checked>*162$13.zzzzzzw060305U6k6O6Ba6S361U0k0Tzw"
		if ok:=FindText(77,205,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Sleep 150
			
			; 체크박스 체크됐으면 1 리턴
			return 1
		}
		
		; 체크박스 체크 안 됐으면 0 리턴
		return 0		
		
	} ; end - isTheCheckBoxChecked()
	
	
	; Uncheck All 버튼 클릭하기
	clickUnCheckAllButtonOnTheAllocatinScreen(){

		Text:="|<Uncheck All button>*126$62.W00A003032MU03000k0ka9wQyC7BUS9WN9gYGPk5WMaG39wUw38a9YUmE8D0z9qN9gaGPM8mL6Fn8sQn66a"
		if ok:=FindText(335,35,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Sleep 150
			Click
			Sleep 150			
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.clickUnCheckAllButtonOnTheAllocatinScreen()
		}		
		return
		
	}
	
	
	
	
	; 화살표 버튼 클릭해서 이 화면 나가기
	clickTheArrowButtonOnTheAllocatinScreen(){
		

		Text:="|<arrow button>*174$22.zzzy000M001U0060U0M301UC060w0M3s1UDk60zUM3z1UDU60q0M2A1U0k601UM061U006000M001zzzy"
		if ok:=FindText(166,35,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 150
		}		
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.clickTheArrowButtonOnTheAllocatinScreen()
		}		
		
		return
	} ; end - clickTheArrowButtonOnTheAllocatinScreen
	
	
	
	; allocate 에러창 뜨는 지 확인 - 1
	checkAllocateErrorWindowAppears_1(){		

		Text:="|<allocate error message - 1>*146$172.49040U2E00000E000U0EV0w7WTST0MY0E201000001000201W44MUAF162WED7SwYMDCQ7Ytnlu8+8EUY1F448+91YYUGG04mE2H91AYUcV22E54SEV4Y4Hm79A1mD0t8wQWI4G4890WF127mEF88YY898U4YW2G9ET8EUY3t448F91YkWGEUYX0GGA9AX14V1688YEFa2Y3lnj9C3m71t8QwS88/rXkR1FxsE00000000000000000U000000000000000000000000000A00000000002"
		if ok:=FindText(993,548,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			
			return 1
		}
		else
			return 0
	} ; 끝 - checkAllocateErrorWindowAppears
	
	
	
	
	; 픽티켓 탭에 있는 Memo 찾아서 BTR 써넣기
	findMemoOnPickTicketTab_and_PutBTROnIt(){

		Text:="|<Memo on Pick Ticket>*184$31.n0003NU003wnnywqtB9GTxyYtVOkGQkh9d+EKbYZs8"
		if ok:=FindText(1099,174,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+30, Y+H//2
			Sleep 500
			Click
			Sleep 150
			Send, {End}
			Sleep 150
			Send, {Enter}
			Sleep 150			
			Send, BTR
			Send, ^s
			Sleep 2500
			Send, ^s
			Sleep 1000
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.findMemoOnPickTicketTab_and_PutBTROnIt()
		}		
		
		return
		
	} ; 끝 - findMemoOnPickTicketTab_and_PutBTROnIt
	
	
	
	
	
	; PICK TICKET 탭에서 인쇄가 몇 번 됐는지 그 횟수 얻기
	getNumberOfPrintCountOnPckTicketTab(){
		
		Text:="|<>*200$55.y20US000Ln00ENU00/szzQ8bYrjwyQYA2OOGzuCO63BBBT17B11aaajUXaUnHHHLkFnMDDDdjs"
		if ok:=FindText(1105,312,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+30, Y+H//2
			Sleep 200
			
			; Print Count 값 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
							
			#ofPrintCount := Clipboard
			Sleep 150			
			
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.getNumberOfPrintCountOnPckTicketTab()
		}				
		
		return #ofPrintCount
		
	}
	
	
	
	; pick ticket 탭에 들어가기 위해 버튼 클릭하기
	clickPickTicketTabToGoToThere(){		

		Text:="|<PICK TICKET>*133$68.s0000000000200000000000U0000U00400DzwSE87y10128V4E2080E0EW8F5DY29wbDDzwFI+0WVG928W7Z308cMyEW8V1Es2+784DzsEI/0WVO10U044yM8bnSQzzk00000000C0s000000003UC000000002"
		if ok:=FindText(203,57,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 1000
		}
		
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.clickPickTicketTabToGoToThere()
		}					
		
		return
	}
	
	
	



	
	
	
	
	
	
} ; 전체 클래스 class CommN41 끝
