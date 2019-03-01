#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include GetActiveBrowserURL.ahk
#Include function.ahk



OnlinePayment(loc_of_MostRecentPo)
{


	;;;;;;;;;;;;;;
	;;;;;;;;;;;;;; 나는 여기 window title 이 url 값인 줄 알았는데 아니었음. 로그인 된 윈도우 창 제목은 그냥 단순하게 Virtual Terminal - Internet Explorer 이었음
	;;;;;;;;;;;;;; 
	; 이전에 열려있는 Virtual POS Terminal URL 주소 읽어오기
	; URLofVirtualPOSTerminal.txt 에 저장되어 있는 url 주소 BeingReadURLofVirtualPosTerminal 변수에 저장하기
	FileRead, BeingReadURLofVirtualPosTerminal, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt
	
	
	; 이미 열려있는 BeingReadURLofVirtualPosTerminal 값의 url 주소가 있는 IE 창 사용하기
	;Paymentwb := IEGet(BeingReadURLofVirtualPosTerminal)


	; 이미 열려있는 Virtual Terminal 의 IE 창 사용하기
	; 윈도우 타이틀 복잡할 줄 알았는데 아래처름 드릅게 단순함
	Paymentwb := IEGet("Virtual Terminal - Internet Explorer")
	
	;MsgBox, % Paymentwb

	

	; 만약 열려있는 IE 창이 없으면 로그인 하기
	if(!Paymentwb){


		; 이전에 있었던 이미 로그아웃 된 창은 닫고 시작하기
		WinClose, Login - Internet Explorer
		

		; 로그인 url
		;URL = https://secure.linkpt.net/lpc/servlet/LPCPage?pagetype=pos_main&_st=3.4F87A07E1C6E8B0A1E2AF11F9E08F4AB&_si=1.0B3AE51F901D9956E0E6CB7FBF604B3EAC94931E57D982E9A5F44A8E552EA71F&_sk=3.CDE807FD9A35D14CD565AF85CE2655A6
		URL = https://secure.linkpt.net/lpc/servlet/LPCLogin



		StoreNumber = 1001343442
		UserID = 343442
		Password = Jon12345!

		Paymentwb := ComObjCreate("InternetExplorer.Application")
		Paymentwb.Visible := True
		Paymentwb.Navigate(URL)
		While Paymentwb.readyState != 4 || Paymentwb.document.readyState != "complete" || Paymentwb.busy ; wait for the page to load
		   Sleep, 10
		   
		
		
		; 창 설정
		; Full screen으로 만든 뒤 글자 크기를 100%로 만들기
		WinMaximize, Login - Internet Explorer
		Send, ^0
		



		Loop{
			
			

			;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
			;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; 팀뷰어로 한거라 위치가 달라질테니 수정하기 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
			PointOfX = 1262
			PointOfY = 450
			PointOftoX = 1390
			PointOftoY = 495
			
			
			
			
			
			; 지정된 위치(현재 페이지)에서 Capture2Text 이용해서 값 얻기
			MouseMove, %PointOfX%, %PointOfY%
			SendInput, #q
			Sleep 200
			MouseMove, %PointOftoX%, %PointOftoY%
			Sleep 200
			SendInput, #q
			Sleep 200	


			; Capture2Text 창 닫기
			WinActivate, Capture2Text - OCR Text
			WinWaitActive, Capture2Text - OCR Text
			WinClose
			
			
			; Capture2Text 창 닫기
;			WinWaitActive, Capture2Text - OCR Text
;			IfWinExist, Capture2Text - OCR Text
;				WinClose


			HumanVerificationCode =: %Clipboard%
			Sleep 350


			; HumanVerificationCode 값에 특수문자 있으면 로그인 시 에러 발생
			; 그렇기 때문에 알파벳과 숫자만 HumanVerificationCode 에 저장
			HumanVerificationCode := RegExReplace(HumanVerificationCode, "[^a-zA-Z0-9]", "")
			

			;MsgBox, HumanVerificationCode is : `n%HumanVerificationCode%


			; 현재 url 얻는 부분
			nTime := A_TickCount
			sURL := GetActiveBrowserURL()
			WinGetClass, sClass, A
			If (sURL != ""){
				;MsgBox, % "The URL is """ sURL """`nEllapsed time: " (A_TickCount - nTime) " ms (" sClass ")"
			}
			Else If sClass In % ModernBrowsers "," LegacyBrowsers
				MsgBox, % "The URL couldn't be determined (" sClass ")"
			Else
				MsgBox, % "Not a browser or browser not supported (" sClass ")"


			; 얻은 현재 url이 로그인 화면이면 로그인 하기
			if(RegExMatch(sURL, "imU)LPCLogin")){
					
	;			Paymentwb.document.getElementById("tbUserID").value := UserID  ;UserID 입력

				Paymentwb.document.getElementsByTagName("INPUT")[1].innerText := StoreNumber  ;StoreNumber 입력
				Paymentwb.document.getElementsByTagName("INPUT")[2].innerText := UserID  ;UserID 입력
				Paymentwb.document.getElementsByTagName("INPUT")[3].innerText := Password  ;Password 입력
				Paymentwb.document.getElementsByTagName("INPUT")[5].innerText := HumanVerificationCode  ;Human Verification Code 입력
				Paymentwb.document.getElementsByTagName("INPUT")[6].Click() ; 로그인 버튼 누르기

						
	;			msgbox, Put in Human Verify Code
						
				While Paymentwb.readyState != 4 || Paymentwb.document.readyState != "complete" || Paymentwb.busy ; wait for the page to load
				   Sleep, 10
			}
				
			
			; 현재 url 얻는 부분
			nTime := A_TickCount
			sURL := GetActiveBrowserURL()
			WinGetClass, sClass, A
			If (sURL != ""){
				;MsgBox, % "The URL is """ sURL """`nEllapsed time: " (A_TickCount - nTime) " ms (" sClass ")"
			}
			Else If sClass In % ModernBrowsers "," LegacyBrowsers
				MsgBox, % "The URL couldn't be determined (" sClass ")"
			Else
				MsgBox, % "Not a browser or browser not supported (" sClass ")"

			
			

			
			
			; url이 로그인 화면이 아니면 payment 화면으로 넘어갔다는 뜻이므로
			; 현재 url 파일에 저장하고 루프를 탈출
			IfNotEqual, sURL, https://secure.linkpt.net/lpc/servlet/LPCLogin
			{
				
				; 우선 이전에 파일에 저장되어 있는 값을 초기화 한 뒤
				; sURL 안에 있는 현재 URL주소를 URLofVirtualPOSTerminal.txt 파일에 저장하기
				FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt, 1
				FileAppend, %sURL%, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt
				
				break
			}
			
			
		}	
		
	}
*/		


	Sleep 500
	WinActivate, Virtual Terminal - Internet Explorer

	
	
/*			
	; 현재 url 얻는 부분
	nTime := A_TickCount
	sURL := GetActiveBrowserURL()
	WinGetClass, sClass, A
	If (sURL != ""){
		;MsgBox, % "The URL is """ sURL """`nEllapsed time: " (A_TickCount - nTime) " ms (" sClass ")"
	}
	Else If sClass In % ModernBrowsers "," LegacyBrowsers
		MsgBox, % "The URL couldn't be determined (" sClass ")"
	Else
		MsgBox, % "Not a browser or browser not supported (" sClass ")"


	; 만약 현재 url 이 이전 결제 완료 창이라면 
	; 새로운 결제 창을 열기 위해 Next Transaction 버튼 클릭하기
	IfEqual, sURL, https://secure.linkpt.net/lpc/servlet/LPCTxn
	{
		Paymentwb.document.getElementsByTagName("INPUT")[57].Click()
		
		While Paymentwb.readyState != 4 || Paymentwb.document.readyState != "complete" || Paymentwb.busy ; wait for the page to load
			Sleep, 10
	}
*/	
	
	
	
	;MsgBox, CC info in the function`n`n %CCinfoOnFASHIONGO%
	
	; Virtual Terminal 실행
	; 1. 계속 바뀌는 url 어떻게 처리할건가
	; 2. 일정 시간 지나면 자동 로그아웃 되는 문제






	; 창 설정
	; Full screen으로 만든 뒤 글자 크기를 100%로 만들기
	;WinMaximize, %BeingReadURLofVirtualPosTerminal% - Internet Explorer
	WinMaximize, Virtual Terminal - Internet Explorer
	Send, ^0


   
	   
	; Invoice Number
	;Paymentwb.document.getElementsByTagName("INPUT")[18].value := Invoice_No ;이것도 작동됨
	;Paymentwb.document.getElementsByTagName("INPUT")[18].innerText := Invoice_No
	Paymentwb.document.getElementsByTagName("INPUT")[17].innerText := Invoice_No
	
	; InvoiceBalance
	; 일반적인 COM 함수로 입력하니 사람이 아닌 기계가 입력한 값이라고 여기는지 자동 계산이 안되고 경고창이 떴음
	; 그래서 일단 입력창에 포커스를 맞추고 수동으로 값 입력하는 것으로 바꿈
	;Paymentwb.document.getElementsByTagName("INPUT")[19].innerText := SubTotal ; 이렇게 하면 값이 입력은 되는데 기계입력으로 받아들이는 듯. 배송비 계산이 자동으로 안 됨
	;Paymentwb.document.getElementsByTagName("INPUT")[19].value := SubTotal ; 이렇게 하면 값이 입력은 되는데 기계입력으로 받아들이는 듯. 배송비 계산이 자동으로 안 됨
	Paymentwb.document.getElementsByTagName("INPUT")[19].focus()
	SendInput, % InvoiceBalance
	SendInput, {tab}
	
	; Shipping Amount
	; InvoiceBalance 에 이미 배송비까지 포함되어 계산됐으니 배송료를 따로 입력할 필요 없다
	;Paymentwb.document.getElementsByTagName("INPUT")[20].innerText := RoundedShippingFee; 이것도 작동됨
	;Paymentwb.document.getElementsByTagName("INPUT")[20].value := RoundedShippingFee

	; Tax Exempt Checkbox Button
	Sleep 100
;	Paymentwb.document.getElementsByTagName("INPUT")[22].checked := true
	Paymentwb.document.getElementsByTagName("INPUT")[22].focus()
	Sleep 100
	SendInput, {Space}
	Sleep 100
	;SendInput, {Enter}
	;Sleep 100
	


	; Transaction Origin DropdownList
	Paymentwb.document.getElementsByTagName("SELECT")[0].selectedIndex := 3
	; Transaction Type DropdownList
;	Paymentwb.document.getElementsByTagName("SELECT")[1].selectedIndex := 1



	; Credit Card Number
	Paymentwb.document.getElementsByTagName("INPUT")[26].innerText := CCNumbers




	; Expiration Month DropdownList
	;Paymentwb.document.getElementsByTagName("SELECT")[2].value := Month
	;Paymentwb.document.getElementsByTagName("SELECT")[2].value := "07" ;이렇게 하면 작동됨
	Paymentwb.document.getElementsByTagName("SELECT")[2].focus()
	SendInput, % Month
	SendInput, {Enter}
		


	; Expiration Year DropdownList
	;Paymentwb.document.getElementsByTagName("SELECT")[3].value := Year
	;Paymentwb.document.getElementsByTagName("SELECT")[3].value := 2027 ;이렇게 하면 작동됨
	Paymentwb.document.getElementsByTagName("SELECT")[3].focus()
	SendInput, % Year
	SendInput, {Enter}
	


	; Security Code
	Paymentwb.document.getElementsByTagName("INPUT")[29].innerText := CVV
	; Customer Name
	Paymentwb.document.getElementsByTagName("INPUT")[34].innerText := CompanyName
	; Billing Street Address
	Paymentwb.document.getElementsByTagName("INPUT")[36].innerText := BillingAdd1
	; City
	Paymentwb.document.getElementsByTagName("INPUT")[38].innerText := City



	; State DropdownList
	;Paymentwb.document.getElementsByTagName("SELECT")[4].selectedIndex := 8 ;이건 작동됨 다른건 다 안 됐음
	; 일반적인 COM 함수로 입력하니 사람이 아닌 기계가 입력한 값이라고 여기는지 주가 선택이 안 되고 계속 경고창이 떴음
	; 그래서 일단 입력창에 포커스를 맞추고 수동으로 값 입력하는 것으로 바꿈
	Paymentwb.document.getElementsByTagName("SELECT")[4].focus()
	SendInput, % State
	SendInput, {tab}	
	
	

	; Zip Code
	Paymentwb.document.getElementsByTagName("INPUT")[40].innerText := BillingZip

MsgBox, to click  Countinue Button

	; Countinue Button Click
	Paymentwb.document.getElementsByTagName("INPUT")[56].Click()
	
	
	; 버튼 클릭 후 다음 화면 뜰때까지 로딩 기다리기
	While Paymentwb.readyState != 4 || Paymentwb.document.readyState != "complete" || Paymentwb.busy ; wait for the page to load
		Sleep, 10



	; HUMAN VARIFICATION 코드 나오면 감지한 뒤 처리하기

jpgLocation = %A_ScripDir%PICTURES\HumanVarificationArrowSmallButton.png
ImageSearch, , , 853, 783, 950, 880, %jpgLocation%
if(ErrorLevel = 0){
	
		Loop{
			
			

			;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
			;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; 팀뷰어로 한거라 위치가 달라질테니 수정하기 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
/*			
			PointOfX = 918
			PointOfY = 809
			PointOftoX = 1048
			PointOftoY = 851
*/
			PointOfX = 918
			PointOfY = 837
			PointOftoX = 1048
			PointOftoY = 881

			
			
			; 지정된 위치(현재 페이지)에서 Capture2Text 이용해서 값 얻기
			MouseMove, %PointOfX%, %PointOfY%
			SendInput, #q
			Sleep 200
			MouseMove, %PointOftoX%, %PointOftoY%
			Sleep 200
			SendInput, #q
			Sleep 200	


			; Capture2Text 창 닫기
			WinWaitActive, Capture2Text - OCR Text
			WinClose
			
			
			; Capture2Text 창 닫기
;			WinWaitActive, Capture2Text - OCR Text
;			IfWinExist, Capture2Text - OCR Text
;				WinClose


			HumanVerificationCode =: %Clipboard%
			Sleep 350


			; HumanVerificationCode 값에 특수문자 있으면 로그인 시 에러 발생
			; 그렇기 때문에 알파벳과 숫자만 HumanVerificationCode 에 저장
			HumanVerificationCode := RegExReplace(HumanVerificationCode, "[^a-zA-Z0-9]", "")
			

			;MsgBox, HumanVerificationCode is : `n%HumanVerificationCode%


			jpgLocation = %A_ScripDir%PICTURES\PleaseConfirmOrder.png
			ImageSearch, , , 492, 240, 704, 291, %jpgLocation%
			
			; 현재 창이 Please Confirm Order 화면이면 자동 가입 방지 코드 입력
			if(ErrorLevel = 0){
				Paymentwb.document.getElementsByTagName("INPUT")[48].innerText := HumanVerificationCode  ;Human Verification Code 입력
				Paymentwb.document.getElementsByTagName("INPUT")[49].Click() ; Please Confirm Order 화면에서 Submit 버튼 클릭
						
				While Paymentwb.readyState != 4 || Paymentwb.document.readyState != "complete" || Paymentwb.busy ; wait for the page to load
				   Sleep, 10
			}
				
			
			

			jpgLocation = %A_ScripDir%PICTURES\TransactionResult.png
			ImageSearch, , , 492, 240, 704, 291, %jpgLocation%
			
			; 현재 창이 Transaction Result 화면이면 자동 가입 방지 코드 력이 성공했다는 뜻이므로
			; 루프를 탈출
			if(ErrorLevel = 0)
				break			
		}
		
}





	; Please Confirm Order 화면에서 Submit 버튼 클릭
	;Paymentwb.document.getElementsByTagName("INPUT")[46].Click()
	Paymentwb.document.getElementsByTagName("INPUT")[47].Click()
	;Paymentwb.document.getElementByName("submitbt").Click()
	
	
	
	; 버튼 클릭 후 다음 화면 뜰때까지 로딩 기다리기
	While Paymentwb.readyState != 4 || Paymentwb.document.readyState != "complete" || Paymentwb.busy ; wait for the page to load
		Sleep, 10
	
	; 카드 승인 결과 (APPROVED, DECLINED, FAILED 등) 표시되는 곳에서 승인결과 읽어와서 TransactionStatus 변수에 저장하기
	TransactionStatus := Paymentwb.document.getElementsByTagName("TD")[2].innerText
	
;	TransactionStatus = DECLINED
;	TransactionStatus = APPROVED
	
	MsgBox, TransactionStatus`n`n%TransactionStatus%
	
	


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
 ;;;;;;;;;;;;;;;;;;; 이곳 주석 부분 없애야 됨 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	IfEqual, TransactionStatus, FAILED
	{
		MsgBox, It's Invalid Credit Card Information
		; Next Transaction 버튼 클릭해서 새로운 입력창 열기
		Paymentwb.document.getElementsByTagName("INPUT")[58].Click()		
	}

	IfEqual, TransactionStatus, DECLINED
	{
		MsgBox, It's Declined
		
		;쇼, 전화, 이메일, jodifl.com 주문이라면
		if(loc_of_MostRecentPo = 1){
		
			; CCNumbers2 에 값이 있으면
			if(CCNumbers2){

				CCNumbers := % CCNumbers2
				;ExpDate := % ExpDate2
				CVV := % CVV2
				Month := % Month2
				Year := % Year2
				
				MsgBox, Month2 : %Month2%`n`nMonth : %Month%
				MsgBox, CCNumbers2 : %CCNumbers2%`n`CCNumbers : %CCNumbers%
				
				; 이 if 문은 CCNumbers2 안에 값이 있어야만 실행된다
				; CCNumbers2 값을 이용했는데도 또 디클라인 났을 때 if(CCNumbers2) 가 또 실행되는 것을 막기 위해
				; CCNumbers2 값을 없애서 if(CCNumbers2) 실행되는 것을 막기
				CCNumbers2 :=
				
				
				; 일단 디클라인 났으니 Next Transaction 버튼 클릭해서 새로운 입력창 열고
				; 위의 값으로 결제 다시 시도				
				Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
				OnlinePayment(loc_of_MostRecentPo)
			}

						
			if(CCNumbers3){

				CCNumbers := % CCNumbers3
				;ExpDate := % ExpDate3
				CVV := % CVV3
				Month := % Month3
				Year := % Year3
			
				; 이 if 문은 CCNumbers3 안에 값이 있어야만 실행된다
				; CCNumbers3 값을 이용했는데도 또 디클라인 났을 때 if(CCNumbers3) 가 또 실행되는 것을 막기 위해
				; CCNumbers3 값을 없애서 if(CCNumbers3) 실행되는 것을 막기
				CCNumbers3 :=
				
				
				; 일단 디클라인 났으니 Next Transaction 버튼 클릭해서 새로운 입력창 열고
				; 위의 값으로 결제 다시 시도				
				Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
				OnlinePayment(loc_of_MostRecentPo)
			}

						
			if(CCNumbers4){

				CCNumbers := % CCNumbers4
				;ExpDate := % ExpDate4
				CVV := % CVV4
				Month := % Month4
				Year := % Year4
			
				; 이 if 문은 CCNumbers4 안에 값이 있어야만 실행된다
				; CCNumbers4 값을 이용했는데도 또 디클라인 났을 때 if(CCNumbers4) 가 또 실행되는 것을 막기 위해
				; CCNumbers4 값을 없애서 if(CCNumbers4) 실행되는 것을 막기
				CCNumbers4 :=
				
				
				; 일단 디클라인 났으니 Next Transaction 버튼 클릭해서 새로운 입력창 열고
				; 위의 값으로 결제 다시 시도				
				Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
				OnlinePayment(loc_of_MostRecentPo)
			}
			
			; 위에서 LAMBS 예 잇는 모든 CC 값을 이용했지만 모두 디클라인 났으므로
			; Next Transaction 버튼 클릭해서 새로운 입력창 열고 디클라인 함수 호출
			Paymentwb.document.getElementsByTagName("INPUT")[58].Click()			
			DeclineProcessing()
		}




		; 쇼, 전화, 이메일, jodifl.com 주문이 아니라 패션고, LA쇼룸 주문이면
		; LAMBS 에서 cc정보를 얻어서 거기에 나온 cc 갯수 만큼 시도한 뒤 그래도 디클라인 나면 그때 
		; DeclineProcessing() 호출해야 됨
		
		; iCountForOnlinePayment 값이 1이 아니라면, 즉 CConLAMBS 함수를 호출한 적이 없다면
		; CConLAMBS 함수를 호출해서 LAMBS 에 있는 cc 값들 읽어오기
		IfNotEqual, iCountForOnlinePayment, 1
		{
			; 일단 디클라인 났으니 Next Transaction 버튼 클릭해서 새로운 입력창 열고
			Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
			
			; LAMBS 에서 cc값 얻기
			CConLAMBS(loc_of_MostRecentPo)
			
			; 만약 CConLAMBS 함수를 실행해서 LAMBS 의 CC 값을 읽어왔는데 또 디클라인 났을 때
			; 지금의 IfNotEqual 이 또 실행되는 것을 방지하기 위해 iCountForOnlinePayment 값을 1로 설정하기
			iCountForOnlinePayment = 1
			
			; CConLAMBS 함수를 이용해서 LAMBS 에서 얻어온 cc 값을 이용해서 다시 결제 시도하기
			OnlinePayment(loc_of_MostRecentPo)
		}

		
		; CCNumbers2 에 값이 있으면 if 문 실행
		; 여기서 얻어온 CCNumbers2 값은 패션고나 LA쇼룸 에서 얻은 cc 값이 한 번 디클라인 난 뒤 
		; 위의 
		; IfNotEqual, iCountForOnlinePayment, 1 를 실행해서 얻어온 LAMBS 의 CC 값에서
		; CCNumbers 값이 또 디클라인 난 뒤
		; CCNumbers2 값이 있다면 그것을 이용해서 결제를 시도하기 위한 것

		if(CCNumbers2){

			CCNumbers := % CCNumbers2
			;ExpDate := % ExpDate2
			CVV := % CVV2
			Month := % Month2
			Year := % Year2
			
				
			; 이 if 문은 CCNumbers2 안에 값이 있어야만 실행된다
			; CCNumbers2 값을 이용했는데도 또 디클라인 났을 때 if(CCNumbers2) 가 또 실행되는 것을 막기 위해
			; CCNumbers2 값을 없애서 if(CCNumbers2) 실행되는 것을 막기
			CCNumbers2 :=
				
				
			; 일단 디클라인 났으니 Next Transaction 버튼 클릭해서 새로운 입력창 열고
			; 위의 값으로 결제 다시 시도				
			Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
			OnlinePayment(loc_of_MostRecentPo)
		}

		
		; CCNumbers3 에 값이 있으면 if 문 실행
		; 여기서 얻어온 CCNumbers2 값은 패션고나 LA쇼룸 에서 얻은 cc 값이 한 번 디클라인 난 뒤 
		; 위의 
		; IfNotEqual, iCountForOnlinePayment, 1 를 실행해서 얻어온 LAMBS 의 CC 값에서
		; CCNumbers 값과 CCNumbers2 값이 또 디클라인 난 뒤
		; CCNumbers3 값이 있다면 그것을 이용해서 결제를 시도하기 위한 것

		if(CCNumbers3){

			CCNumbers := % CCNumbers3
			;ExpDate := % ExpDate3
			CVV := % CVV3
			Month := % Month3
			Year := % Year3
			
				
			; 이 if 문은 CCNumbers3 안에 값이 있어야만 실행된다
			; CCNumbers3 값을 이용했는데도 또 디클라인 났을 때 if(CCNumbers3) 가 또 실행되는 것을 막기 위해
			; CCNumbers3 값을 없애서 if(CCNumbers3) 실행되는 것을 막기
			CCNumbers3 :=
				
				
			; 일단 디클라인 났으니 Next Transaction 버튼 클릭해서 새로운 입력창 열고
			; 위의 값으로 결제 다시 시도				
			Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
			OnlinePayment(loc_of_MostRecentPo)
		}

		
		; CCNumbers4 에 값이 있으면 if 문 실행
		; 여기서 얻어온 CCNumbers4 값은 패션고나 LA쇼룸 에서 얻은 cc 값이 한 번 디클라인 난 뒤 
		; 위의 
		; IfNotEqual, iCountForOnlinePayment, 1 를 실행해서 얻어온 LAMBS 의 CC 값에서
		; CCNumbers 값, CCNumbers2 값, CCNumbers3 값이 또 디클라인 난 뒤
		; CCNumbers4 값이 있다면 그것을 이용해서 결제를 시도하기 위한 것

		if(CCNumbers4){

			CCNumbers := % CCNumbers4
			;ExpDate := % ExpDate4
			CVV := % CVV4
			Month := % Month4
			Year := % Year4
			
				
			; 이 if 문은 CCNumbers4 안에 값이 있어야만 실행된다
			; CCNumbers4 값을 이용했는데도 또 디클라인 났을 때 if(CCNumbers4) 가 또 실행되는 것을 막기 위해
			; CCNumbers4 값을 없애서 if(CCNumbers3) 실행되는 것을 막기
			CCNumbers4 :=
				
				
			; 일단 디클라인 났으니 Next Transaction 버튼 클릭해서 새로운 입력창 열고
			; 위의 값으로 결제 다시 시도				
			Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
			OnlinePayment(loc_of_MostRecentPo)
		}

			
		; 위에서 LAMBS 예 잇는 모든 CC 값을 이용했지만 모두 디클라인 났으므로
		; Next Transaction 버튼 클릭해서 새로운 입력창 열고 디클라인 함수 호출
		Paymentwb.document.getElementsByTagName("INPUT")[58].Click()			
		DeclineProcessing()
		
	}







	; Show Receipt 클릭 버튼
	Paymentwb.document.getElementsByTagName("INPUT")[57].Click()


	; 버튼 클릭 후 다음 화면 뜰때까지 로딩 기다리기
	While Paymentwb.readyState != 4 || Paymentwb.document.readyState != "complete" || Paymentwb.busy ; wait for the page to load
		Sleep, 10


	WinActivate, Transaction Receipt - Internet Explorer

	WinWaitActive, Transaction Receipt - Internet Explorer


	; Transaction Receipt - Internet Explorer 창의 PRINT 버튼 클릭하기
;	Paymentwb.document.getElementsByTagName("INPUT")[5].Click()



	; 영수증 인쇄 하기 위해 프린터 창 열기
	Send, ^p



	; 프린터 창 활성화 될때까지 기다리기
	windowtitle = Print
	WinWaitActive, % windowtitle


	; 2장 인쇄하기 위해 프린터 창의 Number of copies 값을 2로 바꾸기
	ControlClick, Edit5, %windowtitle%
	Control_InputText("Edit5", 2, windowtitle)
	;Control_InputText("Edit5", 2, windowtitle)
	
	; Print 버튼 눌러서 인쇄하기	
	ControlSend, Button13, {Enter}, %windowtitle%
	
	
	; Print 누르면 모래시계 뜬 뒤 신용카드 영수증 화면으로 돌아간다
	; 그 창 닫기
	WinActivate, Transaction Receipt - Internet Explorer
	WinClose, Transaction Receipt - Internet Explorer


	; Next Transaction 버튼 클릭해서 새로운 입력창 열면서 마무리
	Paymentwb.document.getElementsByTagName("INPUT")[58].Click()
	








;	ControlClick, Edit5, Print
;	Control_InputText("Edit5", "2", Print)
;	ControlSend, Edit5, {Enter}, Print

;	ControlSetText, %TextLoc%, %Contents%, %windowtitle%
	
/*	
	; 영수증 화면의 PRINT 버튼을 클릭하면 Print 설정 확인 창(익스플로러 창이 아님)이 뜬다.
	; Print 창의 Print 버튼
	ControlClick, Button13, Print
	
	; Print 누르면 모래시계 오래 뜬 뒤 신용카드 영수증 화면으로 돌아간다
	; 그 창 닫기
	WinActivate, Transaction Receipt - Internet Explorer
	WinClose, Transaction Receipt - Internet Explorer
	
	
	; Show Receipt 버튼 클릭 한 뒤 신용카드 영수증 화면의 CLOSE 버튼
	Paymentwb.document.getElementsByTagName("INPUT")[6].Click()
*/


	
	
	
	
	
	return
}