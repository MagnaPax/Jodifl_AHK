;~ #NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;~ SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
;~ SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;~ #Include MethodsForAll.ahk
;~ #Include MethodsForPT.ahk
#Include CommN41.ahk


class N41 extends CommN41{

	; ##########
	;
	; ##########
	PutCCInfoInN41(Arr_CC, Arr_BillingAdd){
/*		
		Loop % Arr_CC.Maxindex(){
			MsgBox % "in the method CC Element number " . A_Index . " is " . Arr_CC[A_Index]
		}
				
		Loop % Arr_BillingAdd.Maxindex(){
			MsgBox % "in the method Add Element number " . A_Index . " is " . Arr_BillingAdd[A_Index]
		}
*/		
		
		; 혹시 이전 작업하던 창이 열려있을지도 모르니 그냥 닫고 시작하기
		WinClose, Credit Card Management
		
		; N41 로그인
		CommN41.BasicN41Processing()
		
		
		; ##################### 카드 정보 입력하려는 Customer Master 탭이 이미 열려있어야 됨 안그럼 에러 발생 ########################
		

		; Customer Master Tab 클릭하기
		CommN41.ClickCustomerMasterTab()

		
		

		
		WinActivate, Credit Card Management
		

		
		; 카드 입력창이 닫혀있으면 열기
		IfWinNotExist, Credit Card Management
		{
			; Register Credit Card 아이콘 클릭
			CommN41.OpenRegisterCreditCard()
			
			WinWaitActive, Credit Card Management
			Sleep, 500			
		}
		

		
		; 카드 정보 추가 하기 위해 New 버튼 클릭
		Text:="|<New Button>*129$37.3k00001800000Y0MU00G0AE03vw59mH7y2Z5Nbz1+yezz0bER0w0Fg6kS08Hm8D00000700000U"
		if ok:=FindText(578,284,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  
		  Sleep 500
		}
		

		; Credit Card 글자 찾아서 그곳에서부터 Name을 찾아 입력하기
		Text:="|<Credit Card>*123$52.Q009EC007800V1Y00M7CCS41nbUF99EE98a17oZ10QWM4EGI42G9nF99ENd8Zt3XZUwwVs"
		if ok:=FindText(652,352,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 200
			Send, {Tab}
		  
			; 카드 정보 넣기
			Loop % Arr_CC.Maxindex(){
				Sleep 200
				Send, % Arr_CC[A_Index]
				Send, {Tab}				
				;~ MsgBox % "Element number " . A_Index . " is " . Arr_CC[A_Index]
			}	  

			; Billing Add 정보 넣기
			Loop % Arr_BillingAdd.Maxindex(){
				Sleep 200
				Send, % Arr_BillingAdd[A_Index]
				Send, {Tab}				
				;~ MsgBox % "Element number " . A_Index . " is " . Arr_CC[A_Index]
			}
			
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, Complete CC Info, Click Ok to Save
			
			
			; Save Button 클릭
			Text:="|<Save Button>*124$41.2200000000000200U0004013k008028U00E04ESFCU08s2WX00EAwdy00U+9G4DV4InY8UWDDW7l140000W2800014IE0001zzU0004"
			if ok:=FindText(676,284,150000,150000,0,0,Text)
			{
			  CoordMode, Mouse
			  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			  MouseMove, X+W//2, Y+H//2
			  Sleep 100
			  Click
			  Sleep 100
			  Click
			  Sleep 200
			}
			
			Send, ^s
/*			
			; Close Button 클릭
			Text:="|<Close Button>*143$44.zzw0000DUT00003U1k0000sYA0800Atn1u003CQFWU00bbYE9nndtt42qZ6SCF0cgTXX4E+8oAsl6+qZX6MkybDDs0Q0000D0D00003wDk0000zzw00008"
			if ok:=FindText(1051,284,150000,150000,0,0,Text)
			{
			  CoordMode, Mouse
			  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			  MouseMove, X+W//2, Y+H//2
			  Click
			  Sleep 200
			}
*/			
			; cc 창 닫고 끝내기
			WinClose, Credit Card Management
			
		}
		
		return



/* 이거 구현해야 됨
		IfWinExist, Credit Card Management
		{
			WinActivate
			Sleep 500
			SendInput, {Esc}
			ControlClick, Button2, %N41_login_wintitle%, , l
		}
*/

		
		Loop % Arr_CC.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Arr_CC[A_Index]
		}
		
		return
	}
	
	
	
	; #######################################################
	; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
	; #######################################################
	UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC){
		
		; Customer Master Tab 클릭하기
		CommN41.ClickCustomerMasterTab()
		Sleep 300
		
		; To get Contact Name
		Text:="|<Contact>*149$41.S001005W00600O0QTSwRw1YmN9Bc2BYlm3E4P9YY6lAaHN9gwCAXTCC"
		if ok:=FindText(284,611,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+80, Y+H//2
			Click
			Sleep 100
			Send, ^a^c
			Sleep 100
			
			;~ if Arr_CC[1] not contains %Clipboard%
			Name := % Arr_CC[1]
			; N41에 있는 이름이 웹이나 램스에서 가져온 값과 맞지 않으면 경고창 띄우기
			if Name not contains %Clipboard%
			{
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Alert, CONTACT NAME ON N41 IS NOT MATCHED WITH THE NAME ON FG`n`n`n`nWOULD YOU LIKE TO CONTINUE TO CHANGE THE INFO ON THIS PAGE?`nIF YOU CLICK No, IT WILL STOP CHANGING THE INFO ON N41.
				
				IfMsgBox, No
				{
					;~ MsgBox, 262144, Title, Reload PROGRAM. THIS MESSAGE FROM UpdateInfoOnCustomerInformationEdit_Tab METHOD IN N41.ahk
					;~ Reload
					return
				}
				
				; Yes 눌렀으면 Contact에 값 입력하기
				IfMsgBox, Yes
				{
					MouseMove, (X+W)+80, Y+H//2
					Click
					Send, % Name
				}				
			}
				
		}

		
		; Addr1 Update
		Text:="|<Addr1>*144$31.60E80X0840nkwSStMqPAJYF8a2z8YH1FanNUsNswkM"
		if ok:=FindText(278,461,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+250, Y+H//2
			Click
			Sleep 100
			Send, {End}
			
			Loop, 100{
				Send, {BackSpace}
			}
			
			Send, % Arr_ShippingAdd[1]

			Sleep 200
		}
		
		; Addr2 Update
		Text:="|<Addr2>*146$33.60E83kk210GD3ltsNMqPA6N4G9VnwWFAAFanNX66SDATU"
		if ok:=FindText(279,484,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+150, Y+H//2
			Click
			Sleep 100
			Send, {End}
			
			Loop, 100{
				Send, {BackSpace}
			}
			
			Send, % Arr_ShippingAdd[2]
			
			Sleep 200
		  
		}
		
		; Zip Update
		Text:="|<Zip>*143$17.Tc0600AbklAX2Fg4XE9ZzHs040088"
		if ok:=FindText(479,508,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+100, Y+H//2
			Click
			Sleep 100
			Send, {End}
			
			Loop, 100{
				Send, {BackSpace}
			}
			
			Send, % Arr_ShippingAdd[5]
			Sleep 500
			Send, {Tab}
;			Sleep 1000
			
;			IfWinActive, City & State
			WinWaitActive, City & State, , 2
			{
;				MsgBox, City & State 활성화됨
				Sleep 500
				Send, {Left}
				Sleep 200
				Send, {Enter}
				Sleep 200
			}
			
			Sleep 200
			
			; 혹시 ZIP 코드가 없는 잘못된 주소가 들어갔나 확인하기
			Clipboard :=
			MouseMove, (X+W)+100, Y+H//2
			Click
			Send, ^a^c
			Sleep 200
			if(Clipboard == ""){
				MsgBox, 262144, Alert, WRONG ADDRESS. PLEASE CHECK.`n`nCLICK OK TO CONTINUE
			}
		}

		Sleep 100
		
		; Auto Email, Divison, Warehouse 그리고 Ship via, Terms, Pay.Method, Priority 업데이트 하기
		N41.UpdateInfoForAutoEmail()
		
		Sleep 150


		; 바뀐 내용 저장하기
		Send, ^s
		Sleep 100
		Send, ^s
		
		
		; Save to click the Save Icon
		; 저장 제대로 하기 위해 아이콘 3번 클릭해서 저장하자
		Loop, 3
		{
			Text:="|<Save Icon>*190$17.zzz00600AzwNzsnzlbzXDz600A00MzsnzlaTXAz700Dzzs"
			if ok:=FindText(384,130,150000,150000,0,0,Text)
			{
			  CoordMode, Mouse
			  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			  MouseMove, X+W//2, Y+H//2
			  Click
			}
			
			Sleep 150
		}


		; 저장 후 경고창 뜨면 고치도록 알려주기
		IfWinActive, User Setting - Column Check
		{
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 4096, Alert, PLEASE FILL IN THE INFORMATION OF ACCOUNTING INFO
		}
		
		return
	}
	
	
	00000000000000NEW__UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC){
		
		; Customer Master Tab 클릭하기
		CommN41.ClickCustomerMasterTab()
		Sleep 300
		
		; 이름 읽은 뒤 화면의 이름과 배열의 이름이 같지 않으면 경고창 띄우기
		Text:="|<Contact>*149$41.S001005W00600O0QTSwRw1YmN9Bc2BYlm3E4P9YY6lAaHN9gwCAXTCC"
		if ok:=FindText(284,611,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+80, Y+H//2
			Click
			Sleep 100
			Send, ^a^c
			Sleep 100
			
			Name := % Arr_CC[1]
			; N41에 있는 이름이 웹이나 램스에서 가져온 값과 맞지 않으면 경고창 띄우기
			if Name not contains %Clipboard%
			{
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Alert, CONTACT NAME ON N41 IS NOT MATCHED WITH THE NAME ON FG`n`n`n`nWOULD YOU LIKE TO CONTINUE TO CHANGE THE INFO ON THIS PAGE?`nIF YOU CLICK No, IT WILL STOP CHANGING THE INFO ON N41.
				
				IfMsgBox, No
				{
					;~ MsgBox, 262144, Title, Reload PROGRAM. THIS MESSAGE FROM UpdateInfoOnCustomerInformationEdit_Tab METHOD IN N41.ahk
					;~ Reload
					return
				}
				
				; Yes 눌렀으면 Contact에 값 입력하기
				IfMsgBox, Yes
				{
					MouseMove, (X+W)+80, Y+H//2
					Click
					Send, % Name
				}				
			}				
		}
		
		
		
		
		return
	}
	
	
	
	; #################################################################
	; Auto Email, Divison, Warehouse 그리고 Ship via, Terms, Pay.Method, Priority 업데이트 하기
	; #################################################################
	UpdateInfoForAutoEmail(){
		
		Clipboard := ""
		
		; 마지막인지 알아내기 위해 이름 확인하기 
		Text:="|<Name>*173$68.00000000000E000000000040000000000100000000000QU000000000780000000001uSTwS000000TYavAU000007sxarw000001yvNBU000000RhqHBU000007NxYlk00000100000000000E000000000040000000000100000000000E00000000006"
		if ok:=FindText(298,244,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5			
			MouseMove, (X+W)+5, Y+H//2
			
			Click
			Sleep 150
			Send, ^a^c
			Sleep 150
			
;			MsgBox, % Clipboard
			
			
			; 만약 고객 이름이 ZYNERGY 면 모든 고객의 정보를 바꿨으니 끝내기
			;~ if(%Clipboard% == ZYNERGY){
			if Clipboard in ZYNERGY
			{			
				MsgBox, 262144, 끝, it's end of the list
				ExitApp
			}
			
;			MsgBox, % Clipboard
			
			
			; To read email
			Text:="|<email>*174$68.00000000000E000000000040000000000100000000000TU001g00000600003000001U3zbak00000T0rN9g0000060AqDP000001Vv9iqk00000M0mPRg000007sAaTP00000100000000000E000000000040000000000100000000000E00000000006"
			if ok:=FindText(298,657,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, (X+W)+2, Y+H//2
				
				Click
				Sleep 150
				Send, ^a^c
				Sleep 150
				
				email := Clipboard
				
;				MsgBox, % "email : " . email
				
				; 이메일에 이메일 주소가 아닌 이상한 값이 들어있으면 그냥 넘어가기
				if email not contains @
				;~ if email contains User Setting
				{
					; 오른쪽 화살표 눌러서 다음으로 넘어가기
					Text:="|<Arrow To Right>*191$74.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzrzzzTz7zzztzwTzznzlzzzsTz1zzwDwTzzs7zkDzz1z7zzs1zw0zzk7lzzw0Tz03zw0QTzw07zk0Dz037zy01zw03zk0Fzzs0Tz01zw0QTzz07zk1zz0D7zzw1zw1zzkDlzzzkTz1zzw7wTzzz7zkzzz7z7zzzxzwzzznzlzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzU"
					if ok:=FindText(505,130,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, Y+H//2
						Click
						Sleep 1000
						; 저장할 지 묻는 경고창에 No 클릭 하기
						IfWinActive, Customer Master - Customer Edit
						{
							Sleep 150
							Send, {Enter}
						}
						Sleep 1000
					}
					
					; 재귀 호출로 처음부터 다시 시작하기
					N41.UpdateInfoForAutoEmail()
;					MsgBox, garbage value in the email, so moved to the next cust info
				}
				
				
				; Auto Inv.Email Adr
				Text:="|<Auto Inv.Email Adr>*143$99.60100U00w00183080k0M0400400010M10D4rb0bqMUztt87VttMaNY4mG7aP990gPAN4n8kaHkUW8t8AW9XwaN64mS44F991yFAFanAUaFYUW/988nNa6yAs4mAboFDd33DAU"
				if ok:=FindText(765,359,150000,150000,0,0,Text)
				{					
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, (X+W)+10, Y+H//2
					
					Click
					Sleep 150
					Send, ^a^c
					Sleep 150
					
					Send, %email%
					
					Sleep 150
					
					
					; 체크박스 찾아서 클릭
					Text:="|<Auto Inv. Email CHECKBOX>*145$173.00000000000000000000000000000000000000000000000000000007zw0000000000000000000000000080800000000000000000000000000E0Ek0804000w00180000000000000U0VU0k080010000E0000000000001017WPnUHvA23zbYU000000000000202/4nAUaGE7aP990000000000000404m9aFVAbU88WCG0000000000000809yHAX2ND0EF4YY0000000000000E0GAqNY4mAUUW/980000000000000U0gBwNk9YN1x4HuE00000000000010100000000000000000000000000202000000000000000000000000007zy"
					if ok:=FindText(982,335,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						;~ MouseMove, X+W//2, Y+H//2
						MouseMove, (X+W)-8, Y+H//2
						Sleep 150
						Click
						Sleep 150
					  
					  

						; Checked Auto Inv.Email checkbox
						Text:="|<Auto Inv.Email CHECKED>*148$172.000000000000000000000000007zw00000000000000000000000000E0E000000000000000000000000010160100U007U009000000000000040IM0A02000E00040000000000000E3HlBtk9xa11znmE00000000000010N/4nAUaGE7aP990000000000000535YHAX2ND0EF4QY0000000000000KMLtAmA9Yw114GGE0000000000001D1FanAUaFY44FN900000000000004M73T6Q2N6ETF4yY0000000000000E0E0000000000000000000000000101000000000000000000000000007zy"
						if ok:=FindText(983,335,150000,150000,0,0,Text)
						{
						  CoordMode, Mouse
						  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						  MouseMove, X+W//2, Y+H//2
						  
;						  MsgBox, check box checked, ok to finish
						}
						else
						{
;							MsgBox, check box unchecked. should be checked.					
							
							; 체크박스 찾아서 클릭
							Text:="|<Auto Inv. Email CHECKBOX>*145$173.00000000000000000000000000000000000000000000000000000007zw0000000000000000000000000080800000000000000000000000000E0Ek0804000w00180000000000000U0VU0k080010000E0000000000001017WPnUHvA23zbYU000000000000202/4nAUaGE7aP990000000000000404m9aFVAbU88WCG0000000000000809yHAX2ND0EF4YY0000000000000E0GAqNY4mAUUW/980000000000000U0gBwNk9YN1x4HuE00000000000010100000000000000000000000000202000000000000000000000000007zy"
							if ok:=FindText(982,335,150000,150000,0,0,Text)
							{
							  CoordMode, Mouse
							  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
							  ;~ MouseMove, X+W//2, Y+H//2
							  MouseMove, (X+W)-8, Y+H//2
							  
							  Click
							  Sleep 150
							  
							}				
						
						}  
					  
					}				
					
				}

				
				; 위의 신용카드 정보가 업데이트가 안 되도 여기부터는 동작함
				; Division 찾아서 업데이트 한 뒤 탭키 눌러서 Warehouse 업데이트
				Text:="|<Division>*147$41.wE208014000002BNcsXXwOGHNAaMowbWFgldt3YXNWFWN9YnsX4SFla"
				if ok:=FindText(284,359,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					;~ MouseMove, X+W//2, Y+H//2
					MouseMove, (X+W)+30, Y+H//2
					
					Click
					Sleep 150
					SendInput, J
					;~ Send, JODIFL
					Sleep 150
					Send, {Tab}
					Sleep 150
					SendInput, W
					;~ Send, WH
					Sleep 150
					Send, {Tab}
					Sleep 150
					
				}
				
				
				Sleep 1500
				; Warehouse 에 WH 제대로 입력 안 됐을 때 처리
				IfWinExist, WAREHOUSE
				{					
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, WAREHOUSE
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, THE STATUS OF Warehouse IS NOT CHANGED.`nPLEASE CHANGE THE STATUS AS 'WH' MANUALLY THEN CLICK OK BUTTON TO CONTINUE THE PROCESS.`n`n상태를 수동으로 WH로 바꿔주세요
				}				
				
				
				
				
				; Memo 내용 있는지 확인하기
				Clipboard := ""
				Text:="|<Memo>*199$34.ss0003XU000CSSTwSxv9anDpjqN8xqUN4XqP9YH/Nb6F7U00000000000000000000000000000000000000000000000000000000000003zzzzzs"
				if ok:=FindText(1138,202,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					;~ MouseMove, X+W//2, Y+H//2
					;~ MouseMove, X+W//2, (Y+H)+5
					MouseMove, X+5, (Y+H)+5
				  
					Click
					Sleep 150
					Send, ^a^c
					Sleep 150
					
;					MsgBox, %Clipboard%

					
					; Memo 에 내용이 없으면 UPS 번호 있는지 확인 후 그것도 없으면
					; SHIP VIA, Terms, Pay.Method 업데이트 하기					
					if(Clipboard == ""){					
						
;						MsgBox, 262144, Empty, No value in Clipboard
						
						; UPS # 있는지 확인
						Clipboard := ""
						Text:="|<UPS ACCOUNT #>*143$85.WT7UA7VsS8aLw5FAmE669WNYH8E2caN07W0U8G9o87wHAs2l0E4B4u41+9wC38U826WL20Z4k1VyE412F9V0zqMAkXAH4nBYkU+CA3kknkwD3WME58"
						if ok:=FindText(759,463,150000,150000,0,0,Text)
						{							
							CoordMode, Mouse
							X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
							;~ MouseMove, X+W//2, Y+H//2						  
							MouseMove, (X+W)+25, Y+H//2
						  
							Click
							Sleep 150
							Send, ^a^c
							Sleep 150
							
;							MsgBox, %Clipboard%
							
							; UPS 번호가 없으면 세가지 값 체인지
							if(Clipboard == ""){
;								MsgBox, 262144, Empty, No value in Clipboard
								
								
								; Ship Via 찾아서 UPSG 로 바꾼 뒤 탭키로 넘어가서 Terms 와 Pay.Method 바꿔주기
								Text:="|<Ship Via>*145$43.SM80AB09A0024047mT1aHnX9AUn98tYYMB4Q6mGA7WGnN9Y3lNDAYy0kbk00E00000080002"
								if ok:=FindText(285,314,150000,150000,0,0,Text)
								{
									CoordMode, Mouse
									X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
									;~ MouseMove, X+W//2, Y+H//2
									MouseMove, (X+W)+30, Y+H//2
									
									Click
									Sleep 150
									Send, ^a^c
									Sleep 150
									
									Send, UPSG
									Sleep 150
									
									; Terms 까지 가기
									Loop, 24
									{
										Sleep 200
										Send, {Tab}
									}
									
									Send, ^a^c
									Sleep 150
									Send, DOR
									Sleep 150
									Send, {Tab}
									Sleep 150
									Send, {Tab}
									Send, CREDIT CARD
									Sleep 150
									
									
								}
							}							
						}
					} ; Memo 안에 내용 없으면 끝남
					
					
					; 만약 메모 내용이 Created automatically by Staging 이면 새 고객이므로 값들 업데이트
					if Clipboard in Created automatically by Staging
					{
						Send, {Delete}
						
;						MsgBox, 262144, New Customer, Created automatically by Staging is in the Memo
						
						; UPS # 있는지 확인
						Clipboard := ""
						Text:="|<UPS ACCOUNT #>*143$85.WT7UA7VsS8aLw5FAmE669WNYH8E2caN07W0U8G9o87wHAs2l0E4B4u41+9wC38U826WL20Z4k1VyE412F9V0zqMAkXAH4nBYkU+CA3kknkwD3WME58"
						if ok:=FindText(759,463,150000,150000,0,0,Text)
						{							
							CoordMode, Mouse
							X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
							;~ MouseMove, X+W//2, Y+H//2							
							MouseMove, (X+W)+25, Y+H//2
						  
							Click
							Sleep 150
							Send, ^a^c
							Sleep 150
							
;							MsgBox, 262144, UPS, %Clipboard%
							
							; UPS 번호가 없으면 세가지 값 체인지
							if(Clipboard == ""){
;								MsgBox, 262144, Empty, No UPS Account #


								; Priority 를 2로 바꾸기
								Text:="|<Priority>*179$68.00000000000E000000000040000000000100000000000TUA01Y000007A0003000001nSnnryM0000QqBilha00007tXlgPD00001kMwP6nk0000Q6BglgM000071XDAPa00001000001U0000E00001k000040000000000100000000000E00000000006"
								if ok:=FindText(506,313,150000,150000,0,0,Text)
								{
									CoordMode, Mouse
									X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
									;~ MouseMove, X+W//2, Y+H//2
									;~ MouseMove, X+5, (Y+H)+5
									MouseMove, (X+W)+5, Y+H//2
								  
									Click
									Sleep 150
									Send, ^a^c
									Sleep 150
									Send, 2
									Sleep 150								

								}

								
								
								; Ship Via 찾아서 UPSG 로 바꾼 뒤 탭키로 넘어가서 Terms 와 Pay.Method 바꿔주기
								Text:="|<Ship Via>*145$43.SM80AB09A0024047mT1aHnX9AUn98tYYMB4Q6mGA7WGnN9Y3lNDAYy0kbk00E00000080002"
								if ok:=FindText(285,314,150000,150000,0,0,Text)
								{
									CoordMode, Mouse
									X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
									;~ MouseMove, X+W//2, Y+H//2
									MouseMove, (X+W)+30, Y+H//2
									
									Click
									Sleep 150
									Send, ^a^c
									Sleep 150
									
									Send, UPSG
									Sleep 150
									
									; Terms 까지 가기
									Loop, 24
									{
										Sleep 200
										Send, {Tab}
									}
									
									Send, ^a^c
									Sleep 150
									Send, DOR
									Sleep 150
									Send, {Tab}
									Sleep 150
									Send, {Tab}
									Send, CREDIT CARD
									Sleep 150
									
									
								}
							}							
						}
					} ; Memo 내용이 Created automatically by Staging 인것 처리 끝남
					
					
					
					
					
					
					
				}
				
				
				; 에러창이 뜰 때 처리
				
/*				; Warehouse 에 WH 제대로 입력 안 됐을 때 처리
				IfWinExist, WAREHOUSE
				{					
					; 에러 메세지 없애기					
					Sleep 150
					WinActivate, WAREHOUSE
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT COULD NOT CHANGE THE STATUS OF Warehouse AS 'WH'.`nPLEASE CHANGE THE STATUS MANUALLY THEN CLICK OK BUTTON TO CONTINUE THE PROCESS.
				}
*/				
				IfWinExist, DIVISION
				{
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, DIVISION					
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT'S AN ERROR OCCURED					
				}
								
				IfWinExist, TERM
				{
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, TERM					
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT'S AN ERROR OCCURED					
				}
												
				IfWinExist, PAYMENTCODE
				{
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, PAYMENTCODE								
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT'S AN ERROR OCCURED					
				}
																
				IfWinExist, DataWindow Error
				{
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, DataWindow Error					
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT'S AN ERROR OCCURED					
				}

				IfWinExist, Connection Error
				{
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, Connection Error					
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT'S AN ERROR OCCURED					
				}
				
				IfWinExist, User Setting - Column Check
				{
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, User Setting - Column Check
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 리프레쉬 버튼 클릭
;					CommN41.ClickRefreshButtonOnCustomerMaster()
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150			
					
					; 재귀호출로 아예 처음부터 다시 시작
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT'S AN ERROR OCCURED					
				}
				
				
				

/*
				; 저장하기
				Send, ^s
				
				Sleep 150
				
				; Save to click the Save Icon
				Text:="|<Save Icon>*190$17.zzz00600AzwNzsnzlbzXDz600A00MzsnzlaTXAz700Dzzs"
				if ok:=FindText(384,130,150000,150000,0,0,Text)
				{
				  CoordMode, Mouse
				  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				  MouseMove, X+W//2, Y+H//2
				  Click
				}
				
				Sleep 200
*/


				; Customer Info 의 Type 에러
				; 이건 그냥 넘어가자
				IfWinExist, Customer Master - Customer Edit
				{
					; 에러 메세지 없애기
					Sleep 150
					WinActivate, Customer Master - Customer Edit						
					Sleep 150
					Send, {Enter}			
					Sleep 150
					
					; 이건 그냥 건너뛰기
					; 오른쪽 화살표 눌러서 다음으로 넘어가기
					Text:="|<Arrow To Right>*191$74.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzrzzzTz7zzztzwTzznzlzzzsTz1zzwDwTzzs7zkDzz1z7zzs1zw0zzk7lzzw0Tz03zw0QTzw07zk0Dz037zy01zw03zk0Fzzs0Tz01zw0QTzz07zk1zz0D7zzw1zw1zzkDlzzzkTz1zzw7wTzzz7zkzzz7z7zzzxzwzzznzlzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzU"
					if ok:=FindText(505,130,150000,150000,0,0,Text)
					{
					  CoordMode, Mouse
					  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					  MouseMove, X+W//2, Y+H//2
					  Click
					  Sleep 150
					}
					
					; 저장할거냐는 경고창에 No 선택
					Sleep 150
					Send, {Enter}
					Sleep 150
					
					; 건너뛴 뒤 다시 시작하기
;					N41.UpdateInfoForAutoEmail()
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, ERROR, IT'S AN ERROR OCCURED					
				}


/*
				; 오른쪽 화살표 눌러서 다음으로 넘어가기
				Text:="|<Arrow To Right>*191$74.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzrzzzTz7zzztzwTzznzlzzzsTz1zzwDwTzzs7zkDzz1z7zzs1zw0zzk7lzzw0Tz03zw0QTzw07zk0Dz037zy01zw03zk0Fzzs0Tz01zw0QTzz07zk1zz0D7zzw1zw1zzkDlzzzkTz1zzw7wTzzz7zkzzz7z7zzzxzwzzznzlzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzU"
				if ok:=FindText(505,130,150000,150000,0,0,Text)
				{
				  CoordMode, Mouse
				  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				  MouseMove, X+W//2, Y+H//2
				  Click
				  Sleep 1000
				}
*/				

			}

		}
		
		
;		MsgBox, 262144, recurrence, Call this method again


		; 재귀 호출로 처음부터 다시 시작하기
;		N41.UpdateInfoForAutoEmail()

		
	}





	; #################################################################
	; 램스에서 만든 엑셀파일에서 정보 얻은 배열을 받아서 Sales Order 만들기
	; #################################################################
	MakeSalesOrderUsingInfoFromLAMBS(Str_Having_1item){
		
/*
			Str_Having_1item[1] := RefinedOrderID
			Str_Having_1item[2] := PO#
			Str_Having_1item[3] := Name
			Str_Having_1item[4] := Style
			Str_Having_1item[5] := Color
			Str_Having_1item[6] := Qty
			Str_Having_1item[7] := OrderLocation			
			Str_Having_1item[8] := Email
			Str_Having_1item[9] := Memo			
			Str_Having_1item[10] := Billing Addr
			Str_Having_1item[11] := Shipping Addr
*/





		Name := Str_Having_1item[3]
		BillingAddr := Str_Having_1item[10]
		ShippingAddr := Str_Having_1item[11]
		

		; N41 로그인
		CommN41.BasicN41Processing()

		
		; Sales Order 클릭
		CommN41.ClickSalesOrderOnTheMenuBar()
		
		
		; preOrderId 값과 Str_Having_1item[1] 값이 같지 않으면 Sales Order 탭에 처음 들어온것이니 고객 찾는 동작 실행하기
		; 만약 같다면 이미 이전에 이 if문 실행했고 그 고객이 갖고있는 다른 아이템이라는 뜻이므로 if문 실행하지 않고 건너뜀
		if (preOrderId != Str_Having_1item[1]){
			


;~ /*

			; ############################################################################################################
			; 커스터머 마스터 탭에 재대로 들어왔는지 확인하는 절차 필요
			; 안 들어왔으면 else 문 써서 수동으로 다시 들어오게끔 한 뒤에 자동처리 계속 하도록 goto문 작성
			; 그냥 이거 하지 말자
			; ############################################################################################################
			
			Clipboard =
			
			; 맨 위 메뉴바에 있는 Customer 클릭하기
			CommN41.ClickCustomerOnTheMenuBar()
			
			; Customer 안에 있는 Customer Master 클릭하기
			CommN41.ClickCustomerMasterTab()


			; Customer Master 탭에 있는 List 클릭하기
			CommN41.ClickListOnCustomerMaster()
			
			; 이메일 상태창의 공란 클릭
			CommN41.ClickTheBlankModeByEmail()
			
			Send, ^a
			Send, % Str_Having_1item[8] ; email 입력
			send, {Enter}
			
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 3000
			Sleep 2000
			
			
			; 이메일이 일치하는 고객 정보가 있는지 확인하기
			MouseClick, r, 279, 210
			Sleep 100
			Loop, 4{
				Send, {Down}
				Sleep 300
			}
			Send, {Enter}
			Sleep 1000
			
			; 고객 정보가 없으면 경고창 띄우기
			if(Clipboard == ""){
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 262144, Alert, COMPANY NAME IS `n`n%Name%`n`n`n`nSHIPPING ADDRESS IS `n`n%ShippingAddr%`n`n`n`nBILLING ADDRESS IS `n`n%BillingAddr%`n`n`n`nTHERE IS NO COMPANY NAME SEARCHED`nPLEASE SEARCH COMPANY THEN CLIKC OK
				Reload
			}
			
;MsgBox, Pause

			
			
;			SoundPlay, %A_WinDir%\Media\Ring06.wav
;			MsgBox, 262144, Notification, COMPANY NAME IS `n`n%Name%`n`n`n`nSHIPPING ADDRESS IS `n`n%ShippingAddr%`n`n`n`nBILLING ADDRESS IS `n`n%BillingAddr%`n`n`n`nOK BUTTON TO CONTINUE PROCESSING
			
			
			; Customer Master 탭에 있는 Create SO 클릭하기
			CommN41.ClickCreateSO()


			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 3000
			Sleep 2000



			; New Button on Sales Order Tab
			; 아이템 입력창인 Sales Order Tab 에 제대로 들어왔는지 확인하는 절차
			; 만약 Sales Order Tab에 있는 New 버튼을 찾았다면 제대로 들어온 것이므로 아래의 동작들 실행
			Loop{
				Text:="|<New Button On Sales Order>*178$37.000000M00000A001000001E000Dtg0207V100k3yP00M1050000U000U0E080E08040BU40203k20100U100U000U0E000Q7s03000001U00000000000000000Dzzzzzw"
				if ok:=FindText(238,130,150000,150000,0,0,Text)
				{

					; 이전에 수정한 SO 정보를 저장하지 않았다는 경고창이 나올 경우
					IfWinActive, Sales Order
					{
						SoundPlay, %A_WinDir%\Media\Ring06.wav
									
						MsgBox, 4100, Alert, Previous SO info has not saved, Would you like to save it?
						
						; Yes 눌렀으면 저장하기
						IfMsgBox, Yes
						{
							WinActivate, Sales Order
							Send, {Left}					
							Sleep 100
							Send, {Enter}					
						}				
						; No 눌렀으면 저장하지 않기
						IfMsgBox, No
						{
							WinActivate, Sales Order
							Send, {Enter}					
						}
					}

					
					; Division 으로 이동
					Loop, 12{
						Send, {Tab}
						Sleep 150
					}
					
					Send, JODIFL ; Division 입력
					Send, {Tab}
					Sleep 100
					Send, % Str_Having_1item[7] ; Order Type 입력
					
					
					; Memo 로 이동
					Loop, 24{
						Send, {Tab}
						Sleep 150
					}

					Send, % Str_Having_1item[9] ; Memo 입력
					Sleep 1000
					
					
					; PO# 로 이동
					Loop, 27{
						Send, {Tab}
						Sleep 150
					}
					Sleep 500
					Send, % Str_Having_1item[2] ; PO# 입력				

				}
				; 만약 Create SO 버튼 안 눌렀으면 다시 누르게 한 뒤 데이터 입력 동작 다시 시작하기
				else{
					SoundPlay, %A_WinDir%\Media\Windows Exclamation.wav					
					MsgBox, 262144, Alert, 1st : CLICK CREATE SO BUTTON`n`n2nd : CLICK OK BUTTON OF THIS WINDOW`n`n3rd : MOVE MOUSE CURSOR TO 'Ship Date' THEN CLICK THAT`n`n`n`nCOMPANY NAME IS `n%Name%`n`nSHIPPING ADDRESS IS `n%ShippingAddr%`n`nBILLING ADDRESS IS `n%BillingAddr%`n`n
					Sleep, 5000
					continue
				}
			
			break
			}


;			MsgBox, pause
			

*/

/*
			; 맨 위 메뉴바에 있는 Customer 클릭하기
			CommN41.ClickCustomerOnTheMenuBar()
			
			; Customer 안에 있는 Customer Master 클릭하기
			CommN41.ClickCustomerMasterTab()


			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, Move to the cursor, CHANGE SEARCH CONDITION AS EMAIL`n`nCLICK OK THEN MOVE MOUSE CURSOR INTO THE BLANK AND CLICK IT			
			
			Sleep 5000
			
			
			
			Send, ^a
			Send, % Str_Having_1item[8] ; email 입력
			send, {Enter}


			
			MsgBox, 262144, Notification, COMPANY NAME IS `n`n%Name%`n`n`n`nSHIPPING ADDRESS IS `n`n%ShippingAddr%`n`n`n`nBILLING ADDRESS IS `n`n%BillingAddr%`n`n`n`nCLICK OK BUTTON THEN CLICK Create SO BUTTON			
			
			Sleep, 5000


			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 3000
			Sleep 2000


			; New Button on Sales Order Tab
			; 아이템 입력창인 Sales Order Tab 에 제대로 들어왔는지 확인하는 절차
			; 만약 Sales Order Tab에 있는 New 버튼을 찾았다면 제대로 들어온 것이므로 아래의 동작들 실행
			Loop{
				Text:="|<New Button On Sales Order>*178$37.000000M00000A001000001E000Dtg0207V100k3yP00M1050000U000U0E080E08040BU40203k20100U100U000U0E000Q7s03000001U00000000000000000Dzzzzzw"
				if ok:=FindText(238,130,150000,150000,0,0,Text)
				{

					; 이전에 수정한 SO 정보를 저장하지 않았다는 경고창이 나올 경우
					IfWinActive, Sales Order
					{
						SoundPlay, %A_WinDir%\Media\Ring06.wav
									
						MsgBox, 4100, Alert, Previous SO info has not saved, Would you like to save it?
						
						; Yes 눌렀으면 저장하기
						IfMsgBox, Yes
						{
							WinActivate, Sales Order
							Send, {Left}					
							Sleep 100
							Send, {Enter}					
						}				
						; No 눌렀으면 저장하지 않기
						IfMsgBox, No
						{
							WinActivate, Sales Order
							Send, {Enter}					
						}
					}

					
					; Division 으로 이동
					Loop, 12{
						Send, {Tab}
						Sleep 100
					}
					
					Send, JODIFL ; Division 입력
					Send, {Tab}
					Sleep 100
					Send, % Str_Having_1item[7] ; Order Type 입력
					
					
					; Memo 로 이동
					Loop, 24{
						Send, {Tab}
						Sleep 100
					}

					Send, % Str_Having_1item[9] ; Memo 입력
					
					
					; PO# 로 이동
					Loop, 27{
						Send, {Tab}
						Sleep 100
					}
					
					Send, % Str_Having_1item[2] ; PO# 입력				

				}
				; 만약 Create SO 버튼 안 눌렀으면 다시 누르게 한 뒤 데이터 입력 동작 다시 시작하기
				else{
					SoundPlay, %A_WinDir%\Media\Windows Exclamation.wav
					MsgBox, 262144, Alert, 1st : CLICK CREATE SO BUTTON`n`n2nd : CLICK OK BUTTON OF THIS WINDOW`n`n3rd : MOVE MOUSE CURSOR TO 'Ship Date' THEN CLICK THAT`n`n`n`nCOMPANY NAME IS `n%Name%`n`nSHIPPING ADDRESS IS `n%ShippingAddr%`n`nBILLING ADDRESS IS `n%BillingAddr%`n`n
					Sleep, 5000
					continue
				}
			
			break
			}
*/







		}




		preOrderId := Str_Having_1item[1] ; 한 오더에 여러 아이템이 있는지 확인하기 위해




		; S로 시작하는 아이템은 스웨터일 가능성이 높으니 옮기지 않기
		Style := Str_Having_1item[4]		

		if Style contains S
		{
			;~ MsgBox, It's sweaters
			return
		}



		; ADD(+) 클릭
		CommN41.ClickAdd()

		
		Send, % Str_Having_1item[4] ; Style
		Sleep 100
		Send, {tab}
		
		Send, % Str_Having_1item[5] ; Color
		Sleep 100
		Send, {tab}
		Send, {tab}
		
		Send, % Str_Having_1item[6] ; Qty
		Sleep 100
		Send, {tab}
		
;~ /*
		Send, ^s
		; Save Button
		CommN41.ClickSave()
*/		
		
		Sleep 1000
		
		
		IfWinActive, Sales Order
		{
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, Alert, AN ERROR MESSAGE APPEARS`n`nIF OK BUTTON CLICKED, THE APPLICATION WILL BE RELOADED
			Reload
		}
		
		; inactive 아이템 있다는 경고창 떴을 때
		IfWinActive, Inactive Style
		{
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, inactive Item, delete inactive item`n`nCLICK OK TO CONTINUE
			
			;~ WinActivate, Inactive Style
			;~ Send, {Enter}
			;~ Sleep 100
			;~ Send, ^d
			;~ Sleep 500
			;~ WinWaitActive, Sales Order
			;~ WinActivate, Sales Order
			;~ Send, {Left}
			;~ Sleep 100
			;~ Send, {Enter}
			;~ Sleep 200			
			
		}
				

		
		
		
		/* 배열로부터 읽기 첫 번째 방법
		Loop % Str_Having_1item.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Str_Having_1item[A_Index]
		}
		*/


		;~ MsgBox, pause on n41 method	
		
		;~ Sleep 500
		; 새로고침 누르면서 끝내기
		;~ Send, ^n
		;~ Sleep 1000		
		
		
		return
		
		

	}
	






	; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지
	FromClickingPreAuthorizedButton_To_PrintOutPickTicket(){
		
			CN41_driver := New CommN41
		
			; pre authorized 버튼 클릭
			Text:="|<pre-authorize Button>*205$16.001zzbzyTztzzY0SE1tzzbzyTztzzc01zzy"
			if ok:=FindText(718,129,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Click
				
				Sleep 1000
				
				; Pre-Authorized 통과 됐거나 Declined 됐을 때
				WinWaitActive Credit Card Processing, , 4
				IfWinActive, Credit Card Processing
				{
					Sleep 500
					Send, {Enter}
					Sleep 500
					
					; Print 버튼 클릭
					Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
					if ok:=FindText(359,129,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, Y+H//2
						Click					
						
						Sleep 1000
						
						; 프린트 창 최대화 하기
						WinWaitActive, Pick Ticket Print
						WinMaximize, Pick Ticket Print
	;~ /*					
						; 안에 있는 프린트 버튼 클릭
						Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
						if ok:=FindText(199,44,150000,150000,0,0,Text)
						{
							CoordMode, Mouse
							X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
							MouseMove, X+W//2, Y+H//2
							Click
							
							Sleep 500
							
							; 에러창 나오면 프로그램 다시 시작하기
							IfWinActive, Microsoft Visual C++ Runtime Library
							{						
								MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
								return
	;							Reload
							}
							
							Send, {Down}
							Sleep 200
							Send, {Down}
							Sleep 200
							Send, {Enter} ; Print Now 눌러서 인쇄하기
							
							Sleep 3000
							WinActivate, Pick Ticket Print
							WinClose, Pick Ticket Print ; 프린트 창 닫기
							Sleep 700						
							
							;~ CommN41.runN41() ; N31 활성화 한 뒤 
							;~ CommN41.OpenSOManager() ; SO Manager 탭 열고 끝내기						
							;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
							
							; SO MANAGER 탭 누르고 끝내기						
							CN41_driver.ClickREfresh()
							
							Send, {Enter} ; 리프레쉬 버튼 누른 뒤 
							Sleep 700
							
							result := CN41_driver.DoesThisPickTicketApproved() ; Approved 됐는지 화면에서 찾아본 뒤 찾았으면 1을 리턴하고 못 찾았으면 0을 리턴
							if(result == 0){					
								MsgBox, Does This Pick Ticket Approved?							
							}						
							
							; SO Manager 탭 클릭해서 pick ticket 탭에서 나오기
							CN41_driver.OpenSOManager()
							
							; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
							; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
							CN41_driver.ClickREfreshButtonOnSOManager()
							

	;						Reload
							return
							
						}
	*/									
						
						
					}
								
					
				}
					
				; CC 가 없어서 업데이트 할거냐고 물을 때
				IfWinActive, Pick Ticket
				{
	;				SoundPlay, %A_WinDir%\Media\Ring06.wav
	;				MsgBox, 262144, Title, CC update`n`nCHECK THE SHIP VIA`n`nCLICK OK TO CONTINUE
					
					Sleep 300				

					Send, {Right}
					Sleep 200
					Send, {Enter}
					Sleep 500			


					
					; Print 버튼 클릭
					Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
					if ok:=FindText(359,129,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, Y+H//2
						Click					
						
						Sleep 1000
						
						; 프린트 창 최대화 하기
						WinWaitActive, Pick Ticket Print
						WinMaximize, Pick Ticket Print
	;~ /*					
						; 안에 있는 프린트 버튼 클릭
						Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
						if ok:=FindText(199,44,150000,150000,0,0,Text)
						{
							CoordMode, Mouse
							X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
							MouseMove, X+W//2, Y+H//2
							Click
							
							Sleep 500
							
							; 에러창 나오면 프로그램 다시 시작하기
							IfWinActive, Microsoft Visual C++ Runtime Library
							{						
								MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
	;							Reload
								return
							}
							
							Send, {Down}
							Sleep 200
							Send, {Down}
							Sleep 200
							Send, {Enter} ; Print Now 눌러서 인쇄하기
							
							Sleep 3000
							WinActivate, Pick Ticket Print
							WinClose, Pick Ticket Print ; 프린트 창 닫기
							Sleep 700						
							
							;~ CommN41.runN41() ; N31 활성화 한 뒤 
							;~ CommN41.OpenSOManager() ; SO Manager 탭 열고 끝내기						
							;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
							
							SoundPlay, %A_WinDir%\Media\Ring06.wav
							;~ MsgBox, 262144, Title, NO CC INFO ON THIS CUSTOMER`n`n`n`nCHECK THE SHIP VIA`n`nCLICK OK TO CONTINUE
							MsgBox, 4100, Wintitle, NO CC INFO ON THIS CUSTOMER`n`nWOULD YOU LIKE TO CONTINUE OR OPEN THE WEB PAGE OF THIS ORDER?`nIF CLICKE No, THE WEB PAGE WILL BE OPENED
							IfMsgBox, No
							{
								FGFromAll_driver := New FG_ProcessingOfPT
								WinClose, ahk_class Chrome_WidgetWin_1 ; 이상하게 열린 창 다시 사용할때 에러난다 차라리 다 닫고 새로 시작하자
								IsItFromNewOrder = 0
								FGFromAll_driver.ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, SO#FromBOExcelFile)
							}

							; SO MANAGER 탭 누르고 끝내기
							CN41_driver := New CommN41
							CN41_driver.OpenSOManager()
							
							
							; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
							; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
							CN41_driver.ClickREfreshButtonOnSOManager()						
							
	;						Reload
							return
							
						}
					}

					;~ Reload
					
				}

			}

		return
	} ; FromClickingPreAuthorizedButton_To_PrintOutPickTicket 메소드 끝















	; Allocation 에서 pre authorized 버튼 누르고 인쇄하기
	Alloc_PrintWithPreAuthorized(){
		
		MsgBox, method in
		
		CN41_driver := New CommN41

		; Allocation 오더에서 온 것일때는 왼쪽 밑의 Open Allocation 에 있는 버튼 누르기
		CN41_driver.ClickCreatePickTicketButtonOnOpenAllocation()
		
		
	
		; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
;		Sleep 3000	
			
		
		WinWaitActive, Pick Ticket ; 확인창
		IfWinActive, Pick Ticket
		{
			Sleep 200
			Send, {Enter}
			Sleep 200
			
			; ; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지
			N41.FromClickingPreAuthorizedButton_To_PrintOutPickTicket()
		}

	
	} ; Alloc_PrintWithPreAuthorized 메소드 끝











} ; 클래스 끝


