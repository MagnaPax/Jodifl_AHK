#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk


UPS_InputAdd(NextDay, 3rdDay, 2ndDay, CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts, NoDeclare){
	
	


	; Invoice_Wts값에 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 Wts_of_Boxes에 저장
	StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.
	
	; Wts_of_Boxes를 루프 돌려서 들어있는 값 개수만큼 No_of_Boxes에 저장
	Loop, %Wts_of_Boxes0%
	{
		
		this_Wts := Wts_of_Boxes%a_index%
	;	MsgBox, The Weight Box number %a_index% is %this_Wts%
		No_of_Boxes := % A_Index
	}
	
	; ups에 입력 시 이메일 주소 앞에 공란에 있으면 에러 발생.
	; 에러 없애기 위한 조치
	; Email 값에 공란(스페이스) 나올 때마다 문자열 나누어 저장. 공란(스페이스)은 제외해서 EliminateSpaceInEmail에 저장
	StringSplit, EliminateSpaceInEmail, Email, %A_Space%, %A_Space%  ; 점이나 콤마는 제외합니다.
	
	; EliminateSpaceInEmail를 루프 돌려서 들어있는 값 개수만큼 Email에 저장
	Loop, %EliminateSpaceInEmail0%{
		;MsgBox, % EliminateSpaceInEmail%A_Index%
		Email := EliminateSpaceInEmail%a_index%		
	}
	;MsgBox, Email:%Email%

	windowtitle = UPS WorldShip - Remote Workstation


/*	
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	ControlSend, Button1, {Enter}, UPS WorldShip
;	Control_SnedButton(Enter, UPS WorldShip)


	;UPS WorldShip 경고창이 활성화로 뜨면 그냥 무시하고 엔터치기
	IfWinActive, UPS WorldShip
		Send, {Enter}
*/
	
	
	WinActivate, % windowtitle
	
	; UPS WorldShip 에 경고창 나오면 없애기
	Loop, 5
	{
		; Alt + Del 눌러서 잘못 입력한 것 취소하기. (작동 되는지는 모르겠음)
		ControlSend, , {!Delete}, UPS WorldShip - Remote Workstation
;		Sleep 200

		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip

		IfWinExist, Software Update
			ControlSend, Button1, {Enter}, Software Update
		
		IfWinExist, Software Update
			WinClose


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
		
		Sleep 200
	}
	
	
	MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nState : %State%`n`nPhone : %Phone%`n`nEmail : %Email%


	MsgBox, 이름 : %CompanyName%`n수령인 : %Attention%`n주소1 : %Address1%`n주소2 : %Address2%`n우편번호 : %ZipCode%`n도시명 : %City%`n전번 : %Phone%`n가격 : %SubTotal%`n이멜 : %Email%`n가격 : %SubTotal%`n`n`n청구소주소 :  %BillingAdd1%`n`n`n청구소ZipCode:  %BillingZip%
	
	

/*
	; F2로 배송 메뉴로 넘긴 뒤 F5로 새로고침 해주기
	key = F2
	Control_SnedButton(Key, windowtitle)
	Sleep 1000
	key = F5
	Control_SnedButton(Key, windowtitle)
	Sleep 1000
	
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	if(WinWait, UPS WorldShip, , 1){
		WinActivate
		ignoreUPSWorldShip()
	}	
	
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	ControlSend, Button1, {Enter}, UPS WorldShip

*/


	
	


	; 회사명 입력
	ControlClick Edit4, %windowtitle%, , l, 2
	Sleep 1000
;	Send, ^a
	SendInput, %CompanyName%
	Sleep 1000
	SendInput, {Del} ;가끔 'A123' 이런 회사명이 'A12'라는 회사명을 입력할 때 대신 입력되는 경우 있어서 Del키 눌러서 방지하기
	Sleep 200
	SendInput, {Enter}
	Sleep 2000


	; Attention 입력	
	Control_InputText("Edit5", Attention, windowtitle)
	
	; Add 1 입력	
	Control_InputText("Edit6", Address1, windowtitle)

	; Add 2 입력	
	Control_InputText("Edit7", Address2, windowtitle)
	
	; Zip Code 입력
	ControlClick Edit9, %windowtitle%, , l, 3
	loop, 10
		SendInput, {BackSpace}
	Sleep 1000
	Send, ^a
	SendInput, %ZipCode%
	Sleep 1000
	SendInput, {Enter}
	Sleep 2000
	
	ignoreUPSWorldShip()
	
	
	
	; 도시명 입력	
	Control_InputText("Edit10", City, windowtitle)

/*
	; 도시명 입력
	ControlSend, Edit10, {^a}, %windowtitle%
;	ControlClick Edit10, %windowtitle%, , l, 3
	loop, 10
		SendInput, {BackSpace}
;	Sleep 1000
;	Send, ^a
	Send, %City%
	Sleep 1000
	SendInput, {Tab}
	Sleep 2000
	
	ignoreUPSWorldShip()
	
	; 주 입력
	ControlSend, ComboBox6, {^a}, %windowtitle%
;	ControlClick Edit10, %windowtitle%, , l, 3
	loop, 5
		SendInput, {BackSpace}
;	Sleep 1000
;	Send, ^a
	Send, %State%
	Sleep 1000
	SendInput, {Tab}
;	Sleep 2000
	
	ignoreUPSWorldShip()	
*/	
	

	; 전화번호 입력	
	Control_InputText("Edit13", Phone, windowtitle)
	

	; 이메일 입력	
	Control_InputText("Edit14", Email, windowtitle)

	
	; 만약 CustomerUPSAccount에 값이 들어있으면 고객 UPS Account 입력 후 Receiver로 바꾸기
	if(CustomerUPSAccount){
		;UPS Account Number공란에 CustomerUPSAccount값 넣기
		Control_InputText("Edit15", CustomerUPSAccount, windowtitle)
		
		;Bill Transportation To: 스크롤 다운 바 클릭
		ControlClick, ComboBox17, %windowtitle%
		
		;아래 화살표 누르기
		ControlSend, ComboBox17, {Down}, %windowtitle%

		;엔터 누르기
		ControlSend, ComboBox17, {Enter}, %windowtitle%
		
		Sleep 500
	}
	
	
	; 만약 3rd Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(3rdDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%

		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%

		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
	
	; 만약 2nd Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(2ndDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%

		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%

		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
		

	; 만약 Next Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(NextDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%

		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		
		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
/*	
	; 무게 입력	
	ControlSend, Edit32, %Invoice_Wts%, %windowtitle%
	sleep 1000
*/	
/*
	; 무게 입력
	if (1 == No_of_Boxes)
		Control_InputText("Edit32", this_Wts, Invoice_Wts)
	else{
		; Invoice_Wts값에 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 Wts_of_Boxes에 저장
		StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.
		
		Loop, %Wts_of_Boxes0%{
			this_Wts := Wts_of_Boxes%a_index%
			;MsgBox, The Weight Box number %a_index% is %this_Wts%			
			Control_InputText("Edit32", this_Wts, windowtitle)
;			Sleep 1000
			ControlClick, Button29, %windowtitle%
			Sleep 1200
		}
	ControlClick, Button3, %windowtitle%
	}
*/	

	; 무게 입력
	if (1 == No_of_Boxes){
		Control_InputText("Edit32", this_Wts, windowtitle)
	}
	else{
		Loop, %Wts_of_Boxes0%{
			this_Wts := Wts_of_Boxes%a_index%
			; 무게 입력란에 값 입력
			Control_InputText("Edit32", this_Wts, windowtitle)
			; Add버튼 클릭
			ControlClick, x325 y642, %windowtitle%
			Sleep 1500
		}
	ControlClick, x371 y619, %windowtitle% ;마지막에 왼쪽 화살표 누르고 끝냄
	}
				
		
/*	
	; Declared Value 입력	
	ControlSend, Edit39, %SubTotal%, %windowtitle%
	sleep 1000
	

	; Declared Value 입력	
	ControlClick, Edit39, %windowtitle%
	Sleep 100
	Control_InputText("Edit39", SubTotal, windowtitle)
	Sleep 100
	ControlSend, Edit39, {Enter}, %windowtitle%
	Sleep 100


	; Declared Value입력
	ControlClick Edit39, %windowtitle%, , l, 3
	loop, 20
		SendInput, {BackSpace}
	Sleep 1000
	Send, ^a
	SendInput, %SubTotal%
	Sleep 1000
	SendInput, {Enter}
	Sleep 2000
	
	; Declared Value입력
	Control_InputText("Edit39", SubTotal, windowtitle)
*/	

	; 고객이 자기 UPS 계정 사용하면서 가격 넣지 말아달라고 요청 시
	if !(NoDeclare){

		; Declared Value입력
		ControlClick Edit39, %windowtitle%, , l, 3
	;	loop, 5
	;		SendInput, {BackSpace}
		Sleep 1000
	;	Send, ^a
		SendInput, %SubTotal%
		Sleep 1000
		SendInput, {Tab}
		Sleep 1000
	}
	
	
/*	
	; Reference number 1 입력
	ControlClick, Edit40, %windowtitle%
	Sleep 100
	
	Control_InputText("Edit40", Invoice_No, windowtitle)
	Sleep 100
	
;	Control_InputText("Edit40", Invoice_No, windowtitle)
;	ControlSend, Edit40, %Invoice_No%, %windowtitle%	

	ControlSend, Edit40, {Enter}, %windowtitle%	
	Sleep 1500
*/	
	
/*
	; 회사명 입력 창으로 가서 위아래로 커서 움직여서 리프레쉬 하기
	ControlSend, Edit4, {Up}, %windowtitle%
	Sleep 3000
	ControlSend, Edit4, {Down}, %windowtitle%
	Sleep 3000
*/

	; Reference number 1 입력
	ControlClick Edit40, %windowtitle%, , l, 3
;	loop, 5
;		SendInput, {BackSpace}
	Sleep 1000
;	Send, ^a
	SendInput, %Invoice_No%
	Sleep 1000
	SendInput, {Tab}
	Sleep 100
	


/*
이거 구현해야 됨
패션고 이외의 주문처에서 주소 읽어올 때 
Country 변수 값이 United States 인지 확인해야 됨
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	; 해외 주문 찾아내기
	IfEqual, State, PR
	{
		MsgBox, It's Order fromDDD Puerto Rico. Now converts to manual mode.
		Reload
	}
	
	IfNotEqual, Country, United States
	{
		MsgBox, It's Order from out of country. Now converts to manual mode.
		Reload
	}
	
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
*/
	

	
	; 배송비 가져오기
	;sleep 1500
	ControlGetText, OriginalShippingFee, Edit44, %windowtitle%
	;MsgBox, % ShippingFee


	; Negotiated 값을 올림하기
	RoundedShippingFee := % Ceil(OriginalShippingFee)
	;MsgBox, % RoundedShippingFee


	;올림 된 배송비 리턴하기
	return %RoundedShippingFee%	


return	
}




















2UPS_InputAdd(NextDay, 3rdDay, 2ndDay, CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts){
	
	windowtitle = UPS WorldShip - Remote Workstation	
	
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	ControlSend, Button1, {Enter}, UPS WorldShip
;	Control_SnedButton(Enter, UPS WorldShip)


	;UPS WorldShip 경고창이 활성화로 뜨면 그냥 무시하고 엔터치기
	IfWinActive, UPS WorldShip
		Send, {Enter}

	
	MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nPhone : %Phone%`n`nEmail : %Email%
	

	
	; F2로 배송 메뉴로 넘긴 뒤 F5로 새로고침 해주기
	key = F2
	Control_SnedButton(Key, windowtitle)
;	Sleep 1000
	key = F5
	Control_SnedButton(Key, windowtitle)
	Sleep 1000
	
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	ControlSend, Button1, {Enter}, UPS WorldShip
	
	
	; 회사명 입력
	ControlClick Edit4, %windowtitle%, , l, 2
	Sleep 1000
;	Send, ^a
	SendInput, %CompanyName%
	Sleep 1000
	SendInput, {Enter}
	Sleep 2000
	
	
/*
	; Attention 입력	
	Control_InputText("Edit5", Attention, windowtitle)
*/
	InputInfoAtUPS(Edit5, Attention, windowtitle)
	Sleep 100
/*	
	
	; Attention 입력	
	ControlClick, Edit5, %windowtitle%
	Sleep 100
	Send, ^a
	SendInput, %Attention%
	SendInput, {Enter}
	Sleep 100
*/	


	; Add 1 입력	
	Control_InputText("Edit6", Address1, windowtitle)	

	; Add 2 입력	
	Control_InputText("Edit7", Address2, windowtitle)

/*
	; Zip Code 입력	
	ControlSend, Edit9, %ZipCode%, %windowtitle%
	ControlSend, Edit9, {Enter}, %windowtitle%	
	sleep 2000
*/

	; Zip Code 입력
	ControlClick, Edit9, %windowtitle%
	Control_InputText("Edit9", ZipCode, windowtitle)
	Control_InputText("Edit9", ZipCode, windowtitle)
	ControlSend, Edit9, {Enter}, %windowtitle%	
	
/*	ControlClick, ComboBox6, %windowtitle%
	sleep 500
	ControlSend, Edit9, {Enter}, %windowtitle%
	Sleep 500
*/	

	;Control_InputText("Edit9", ZipCode, windowtitle)

	; 도시명 입력	
	Control_InputText("Edit10", City, windowtitle)
	ControlClick, Edit10, %windowtitle%
;	Sleep 1500
	
	; 전화번호 입력	
	Control_InputText("Edit13", Phone, windowtitle)

	; 이메일 입력	
	Control_InputText("Edit14", Email, windowtitle)




	
	; 만약 CustomerUPSAccount에 값이 들어있으면 고객 UPS Account 입력 후 Receiver로 바꾸기
	if(CustomerUPSAccount){
		;UPS Account Number공란에 CustomerUPSAccount값 넣기
		Control_InputText("Edit15", CustomerUPSAccount, windowtitle)
		
		;Bill Transportation To: 스크롤 다운 바 클릭
		ControlClick, ComboBox17, %windowtitle%
		
		;아래 화살표 누르기
		ControlSend, ComboBox17, {Down}, %windowtitle%

		;엔터 누르기
		ControlSend, ComboBox17, {Enter}, %windowtitle%
		
		Sleep 500
	}
	
	
	; 만약 3rd Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(3rdDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%

		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%

		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
	
	; 만약 2nd Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(2ndDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%

		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%

		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
		

	; 만약 Next Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(NextDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%

		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		
		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
		
	
	
	
	; 무꼐 입력	
	ControlSend, Edit32, %Invoice_Wts%, %windowtitle%
	sleep 1000
	;Control_InputText("Edit32", Invoice_Wts, windowtitle)
	
/*	
	; Declared Value 입력	
	ControlSend, Edit39, %SubTotal%, %windowtitle%
	sleep 1000
*/	

	; Declared Value 입력	
	ControlClick, Edit39, %windowtitle%
	Sleep 100
	Control_InputText("Edit39", SubTotal, windowtitle)
	Sleep 100
	ControlSend, Edit39, {Enter}, %windowtitle%
	Sleep 100

/*
	; Reference number 1 입력
;	ControlSend, Edit40, %Invoice_No%, %windowtitle%
	Control_InputText("Edit40", Invoice_No, windowtitle)
	sleep 1500
*/

	; Reference number 1 입력
	ControlClick, Edit40, %windowtitle%
	Sleep 100
	
	Control_InputText("Edit40", Invoice_No, windowtitle)
	Sleep 100
	
;	Control_InputText("Edit40", Invoice_No, windowtitle)
;	ControlSend, Edit40, %Invoice_No%, %windowtitle%	

	ControlSend, Edit40, {Enter}, %windowtitle%	
	Sleep 1500
	
	
/*
	; 회사명 입력 창으로 가서 위아래로 커서 움직여서 리프레쉬 하기
	ControlSend, Edit4, {Up}, %windowtitle%
	Sleep 3000
	ControlSend, Edit4, {Down}, %windowtitle%
	Sleep 3000
*/

	
	; 배송비 가져오기
	;sleep 1500
	ControlGetText, OriginalShippingFee, Edit44, %windowtitle%
	

	; 고객 UPS Account가 없는데(있다면 당연히 배송비는 0 되는게 맞지만)
	if(!CustomerUPSAccount){
		; 배송비 가격이 없다면
		if(!OriginalShippingFee){
			; 루프를 5번 돌아라. 만약 배송비를 읽었으면 5번 이전이라도 루프를 나와라
			Loop, 5
			{
				;MsgBox, no
				
				; Declared Value 다시 입력해보기
				ControlClick, Edit39, %windowtitle%
				Sleep 100
				Control_InputText("Edit39", SubTotal, windowtitle)
				Sleep 100
				ControlSend, Edit39, {Enter}, %windowtitle%
				Sleep 100
				
				; control 사용해서 배송비 얻기
				ControlGetText, OriginalShippingFee, Edit44, %windowtitle%
				
				; OriginalShippingFee 에 배송비가 저장됐으면 루프를 탈출해라
				if(OriginalShippingFee)
					break
				else if(!OriginalShippingFee){ ; 만약 계속 배송비가 OriginalShippingFee 에 없으면 콘트롤 함수 말고 마우스로 직접 Declared Value 입력 후 배송비 가져와라
					MouseClick, l, 577, 274
					SendInput, % SubTotal
					MouseClick, l, 578, 508, , , d
					Send, ^a
					OriginalShippingFee := % Clipboard
				}
				; OriginalShippingFee 에 배송비가 저장됐으면 루프를 탈출해라
				if(OriginalShippingFee)
					break
			}
		}
		; 이렇게 루프를 돌고 난리를 쳤는데도 OriginalShippingFee 에 배송비가 입력 안 됐으면 수동으로 입력 받아라
		if(!OriginalShippingFee){
			
			
			; OriginalShippingFee 넣기 위한 GUI
			Gui, Show, w330 h88, Original Shipping Fee

			Gui, Add, Text, x12 y25 w60 h30 +Center, Original Shipping Fee
			Gui, Add, Edit, x72 y28 w180 h20 vOriginalShippingFee
		;	Gui, Add. Button, x92 y28 w30 h30, ASDF

			Gui, Add, Button, x260 y28 w60 h20 +default, Enter>


			;GUI시작 시 포커스를 OriginalShippingFee 입력칸에 위치
			GuiControl, Focus, OriginalShippingFee
			
			; 창이 종료될때까지 대기
			WinWaitClose, Original Shipping Fee
			
			;Enter 버튼 누르면 Tracking Number GUI창 종료
			ButtonEnter>:
			WinClose, Original Shipping Fee
			
			; GUI에서 입력받은 값을 사용하기 위해
			Gui, Submit, nohide
		;	GUI, TrackingNumber:Destroy
			GUI, Hide
			
		}
	}

	
	;MsgBox, % ShippingFee


	; Negotiated 값을 올림하기
	RoundedShippingFee := % Ceil(OriginalShippingFee)
	;MsgBox, % RoundedShippingFee


	;올림 된 배송비 리턴하기
	return %RoundedShippingFee%
	



/*
	회사명4 수령인5 주소1-6 주소2-7 우편번호9 도시명10 전번13 이멜14 고객어카운트15
	고객어카운트Receiver할때 ComboBox17
	무게32
	무게추가Button6
	Declared Value 39
	Reference number1 40
	배송비 44
	인쇄버튼 Button10
*/	



/*
	ControlSend, , {F2}, UPS WorldShip - Remote Workstation
	Sleep 500

	ControlSend, , {F5}, UPS WorldShip - Remote Workstation
	Sleep 500

	ControlSetText, Edit4, CLOVER CHIC BOUTIQUE, UPS WorldShip - Remote Workstation
	ControlSetText, Edit5, ASHLEY HAZLEWOOD, UPS WorldShip - Remote Workstation
	

	;CustomerUPSAccount := 
	
	; 만약 CustomerUPSAccount에 값이 들어있으면 고객 UPS Account 입력 후 Receiver로 바꾸기
	if(CustomerUPSAccount){
		MouseClick, l, 82, 526, 2
		;SendInput, ^a
		SendInput, % CustomerUPSAccount
		
		MsgBox, you're in
		
		;return
	}



		CompanyName = CLOVER CHIC BOUTIQUE
		Attention = ASHLEY HAZLEWOOD
		Address1 = 20103 TREASURE OAKS CT
		Address2 = 
		ZipCode = 77450
		City = TX
		Phone = 6505754396
		Weight = 27
		SubTotal = 1,045.50
		Invoice_No = 71364
		



		
		DragAndPast(226, 242, 10, 242, CompanyName)
		sleep 1500



		DragAndPast(242, 282, 10, 282, Attention)

		DragAndPast(259, 322, 10, 322, Address1)

		DragAndPast(143, 364, 10, 364, Address2)
		
		DragAndPast(274, 403, 150, 403, ZipCode)
		
		DragAndPast(144, 443, 10, 443, City)
		
		DragAndPast(144, 526, 10, 526, Phone)
		

		

		MouseClickAndPaste(338, 534, Invoice_Wts)


		;Declared Value는 에러가 많이 나니까 3번 넣어주자
		loop, 3{
			MouseClickAndPaste(585, 274, SubTotal)
			sleep 1000
		}


		MouseClickAndPaste(537, 310, Invoice_No)
		sleep 1000
		
		;Negotiated 값 얻기
		OriginalShippingFee := ClickAndCtrlAll(580, 510)
		
		; Negotiated 값을 올림하기
		RoundedShippingFee := % Ceil(OriginalShippingFee)
		MsgBox, % RoundedShippingFee
		
		;올림 된 배송비 리턴하기
		return %RoundedShippingFee%
		
		MsgBox, Click OK Button To Continue


		;늘 주문 페이지로 넘어가서 끝낼까?
		SendInput, {F2}
		sleep 500
*/


		; 상자가 한 개 이상인 것 배열에 넣고 처리해야 됨
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		



	
	
return	
}































1UPS_InputAdd(NextDay, 3rdDay, 2ndDay, CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts){
	
	windowtitle = UPS WorldShip - Remote Workstation	
	
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	Control_SnedButton(Enter, UPS WorldShip)


	;UPS WorldShip 경고창이 활성화로 뜨면 그냥 무시하고 엔터치기
	IfWinActive, UPS WorldShip
		Send, {Enter}

	
	MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nPhone : %Phone%`n`nEmail : %Email%
	
	; F2로 배송 메뉴로 넘긴 뒤 F5로 새로고침 해주기
	key = F2
	Control_SnedButton(Key, windowtitle)
;	Sleep 1000
	key = F5
	Control_SnedButton(Key, windowtitle)
;	Sleep 1000
	
	


/*	
	; 회사명 입력	
	ControlClick, Edit4, %windowtitle%
	Control_InputText("Edit4", CompanyName, windowtitle)
	Control_InputText("Edit4", CompanyName, windowtitle)
	ControlSend, Edit4, {Enter}, %windowtitle%
*/

	; 회사명 입력	
	ControlClick, Edit4, %windowtitle%
	Sleep 100
	Send, ^a
	SendInput, %CompanyName%
	SendInput, {Enter}
	Sleep 2000
	
	
	
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;	Control_InputText("Edit4", CompanyName, windowtitle)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;	Sleep 100
	
;	Control_InputText("Edit4", CompanyName, windowtitle)
;	ControlSend, Edit4, {Enter}, %windowtitle%
;	Sleep 100
	
	
/*	
	; 회사명 입력
	ControlSend, Edit4, %CompanyName%, %windowtitle%
	Sleep 500
	ControlSend, Edit4, {Tab}, %windowtitle%
	ControlSend, Edit4, {Tab}, %windowtitle%
	SendInput, %CompanyName%
	Send, {Tab}
*/
;	sleep 3000

/*
	; 회사명 입력	
	Control_InputText("Edit4", CompanyName, windowtitle)
*/
/*
	; Attention 입력	
	Control_InputText("Edit5", Attention, windowtitle)
*/

	; Attention 입력	
	InputInfoAtUPS(Edit5, Attention, windowtitle)
	
/*	
	
	; Attention 입력	
	ControlClick, Edit5, %windowtitle%
	Sleep 100
	Send, ^a
	SendInput, %Attention%
	SendInput, {Enter}
	Sleep 100
*/	

/*
	; Add 1 입력	
	Control_InputText("Edit6", Address1, windowtitle)
*/

	; Add 1 입력	
	InputInfoAtUPS(Edit6, Address1, windowtitle)

/*
	; Add 2 입력	
	Control_InputText("Edit7", Address2, windowtitle)
*/

	; Add 2 입력	
	InputInfoAtUPS(Edit7, Address2, windowtitle)
	
	

/*
	; Zip Code 입력	
	ControlSend, Edit9, %ZipCode%, %windowtitle%
	ControlSend, Edit9, {Enter}, %windowtitle%	
	sleep 2000
*/
/*
	; Zip Code 입력
	ControlClick, Edit9, %windowtitle%
	Control_InputText("Edit9", ZipCode, windowtitle)
	Control_InputText("Edit9", ZipCode, windowtitle)
	ControlSend, Edit9, {Enter}, %windowtitle%	
*/

	; Zip Code 입력	
	InputInfoAtUPS(Edit9, ZipCode, windowtitle)
	Sleep 1500
	
	
	
/*	ControlClick, ComboBox6, %windowtitle%
	sleep 500
	ControlSend, Edit9, {Enter}, %windowtitle%
	Sleep 500
*/	

	;Control_InputText("Edit9", ZipCode, windowtitle)
/*
	; 도시명 입력	
	Control_InputText("Edit10", City, windowtitle)
	ControlClick, Edit10, %windowtitle%
;	Sleep 1500
*/

	; 도시명 입력	
	InputInfoAtUPS(Edit10, City, windowtitle)
	
/*	
	; 전화번호 입력	
	Control_InputText("Edit13", Phone, windowtitle)
*/	
	
	; 전화번호 입력	
	InputInfoAtUPS(Edit13, Phone, windowtitle)
	
/*
	; 이메일 입력	
	Control_InputText("Edit14", Email, windowtitle)
*/

	; 전화번호 입력	
	InputInfoAtUPS(Edit14, Email, windowtitle)	




	
	; 만약 CustomerUPSAccount에 값이 들어있으면 고객 UPS Account 입력 후 Receiver로 바꾸기
	if(CustomerUPSAccount){
		;UPS Account Number공란에 CustomerUPSAccount값 넣기
		Control_InputText("Edit15", CustomerUPSAccount, windowtitle)
		
		;Bill Transportation To: 스크롤 다운 바 클릭
		ControlClick, ComboBox17, %windowtitle%
		Sleep 100
		
		;아래 화살표 누르기
		ControlSend, ComboBox17, {Down}, %windowtitle%
		Sleep 100

		;엔터 누르기
		ControlSend, ComboBox17, {Enter}, %windowtitle%
		
		Sleep 500
	}
	
	
	; 만약 3rd Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(3rdDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%
		Sleep 100
		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100
		
		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
	
	; 만약 2nd Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(2ndDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%
		Sleep 100
		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100
		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100
		
		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
		

	; 만약 Next Day가 체크되어있으면 UPS Service 메뉴 바꾸기
	if(NextDay){
		;UPS Service 스크롤 다운 바 클릭
		ControlClick, ComboBox15, %windowtitle%
		Sleep 100
		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100		
		;위 화살표 누르기
		ControlSend, ComboBox15, {Up}, %windowtitle%
		Sleep 100	
		
		;엔터 누르기
		ControlSend, ComboBox15, {Enter}, %windowtitle%
		
		Sleep 500
	}
		
	
	
/*	
	; 무게 입력	
	ControlSend, Edit32, %Invoice_Wts%, %windowtitle%
	sleep 1000
	;Control_InputText("Edit32", Invoice_Wts, windowtitle)
*/

	; 무게 입력	
	InputInfoAtUPS(Edit32, Invoice_Wts, windowtitle)	

/*	
	; Declared Value 입력	
	ControlSend, Edit39, %SubTotal%, %windowtitle%
	sleep 1000
*/	
/*
	; Declared Value 입력	
	ControlClick, Edit39, %windowtitle%
	Sleep 100
	Control_InputText("Edit39", SubTotal, windowtitle)
	Sleep 100
	ControlSend, Edit39, {Enter}, %windowtitle%
	Sleep 100
*/	
	
	; Declared Value 입력	
	InputInfoAtUPS(Edit39, Invoice_Wts, windowtitle)	
	

/*
	; Reference number 1 입력
;	ControlSend, Edit40, %Invoice_No%, %windowtitle%
	Control_InputText("Edit40", Invoice_No, windowtitle)
	sleep 1500
*/
/*
	; Reference number 1 입력
	ControlClick, Edit40, %windowtitle%
	Sleep 100
	
	Control_InputText("Edit40", Invoice_No, windowtitle)
	Sleep 100
	
;	Control_InputText("Edit40", Invoice_No, windowtitle)
;	ControlSend, Edit40, %Invoice_No%, %windowtitle%	

	ControlSend, Edit40, {Enter}, %windowtitle%	
	Sleep 1500
*/	
	
	; Reference number 1  입력	
	InputInfoAtUPS(Edit40, Invoice_No, windowtitle)	

	
/*
	; 회사명 입력 창으로 가서 위아래로 커서 움직여서 리프레쉬 하기
	ControlSend, Edit4, {Up}, %windowtitle%
	Sleep 3000
	ControlSend, Edit4, {Down}, %windowtitle%
	Sleep 3000
*/

	
	; 배송비 가져오기
	;sleep 1500
	ControlGetText, OriginalShippingFee, Edit44, %windowtitle%
	;MsgBox, % ShippingFee


	; Negotiated 값을 올림하기
	RoundedShippingFee := % Ceil(OriginalShippingFee)
	;MsgBox, % RoundedShippingFee


	;올림 된 배송비 리턴하기
	return %RoundedShippingFee%
	



/*
	회사명4 수령인5 주소1-6 주소2-7 우편번호9 도시명10 전번13 이멜14 고객어카운트15
	고객어카운트Receiver할때 ComboBox17
	무게32
	무게추가Button6
	Declared Value 39
	Reference number1 40
	배송비 44
	인쇄버튼 Button10
*/	



/*
	ControlSend, , {F2}, UPS WorldShip - Remote Workstation
	Sleep 500

	ControlSend, , {F5}, UPS WorldShip - Remote Workstation
	Sleep 500

	ControlSetText, Edit4, CLOVER CHIC BOUTIQUE, UPS WorldShip - Remote Workstation
	ControlSetText, Edit5, ASHLEY HAZLEWOOD, UPS WorldShip - Remote Workstation
	

	;CustomerUPSAccount := 
	
	; 만약 CustomerUPSAccount에 값이 들어있으면 고객 UPS Account 입력 후 Receiver로 바꾸기
	if(CustomerUPSAccount){
		MouseClick, l, 82, 526, 2
		;SendInput, ^a
		SendInput, % CustomerUPSAccount
		
		MsgBox, you're in
		
		;return
	}



		CompanyName = CLOVER CHIC BOUTIQUE
		Attention = ASHLEY HAZLEWOOD
		Address1 = 20103 TREASURE OAKS CT
		Address2 = 
		ZipCode = 77450
		City = TX
		Phone = 6505754396
		Weight = 27
		SubTotal = 1,045.50
		Invoice_No = 71364
		



		
		DragAndPast(226, 242, 10, 242, CompanyName)
		sleep 1500



		DragAndPast(242, 282, 10, 282, Attention)

		DragAndPast(259, 322, 10, 322, Address1)

		DragAndPast(143, 364, 10, 364, Address2)
		
		DragAndPast(274, 403, 150, 403, ZipCode)
		
		DragAndPast(144, 443, 10, 443, City)
		
		DragAndPast(144, 526, 10, 526, Phone)
		

		

		MouseClickAndPaste(338, 534, Invoice_Wts)


		;Declared Value는 에러가 많이 나니까 3번 넣어주자
		loop, 3{
			MouseClickAndPaste(585, 274, SubTotal)
			sleep 1000
		}


		MouseClickAndPaste(537, 310, Invoice_No)
		sleep 1000
		
		;Negotiated 값 얻기
		OriginalShippingFee := ClickAndCtrlAll(580, 510)
		
		; Negotiated 값을 올림하기
		RoundedShippingFee := % Ceil(OriginalShippingFee)
		MsgBox, % RoundedShippingFee
		
		;올림 된 배송비 리턴하기
		return %RoundedShippingFee%
		
		MsgBox, Click OK Button To Continue


		;늘 주문 페이지로 넘어가서 끝낼까?
		SendInput, {F2}
		sleep 500
*/


		; 상자가 한 개 이상인 것 배열에 넣고 처리해야 됨
		

return	
}

