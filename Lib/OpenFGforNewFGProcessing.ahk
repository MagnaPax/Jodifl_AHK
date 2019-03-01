/*
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
*/

#Include function.ahk
#Include FindTextFunctionONLY.ahk


;~ global CurrentOrderIdNumber, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, CustomerNoteOnWebVal, StaffOnlyNoteVal, CompanyName, PendingOrderStatus, AlreadyProcessedItem, CurrentPONumber, ClickLOC, AdditionalInfo, driver, sURL, PO_Number, Order_ID_Only



; 오토핫키 urldownloadtofile

;해외 배송인지 확인하는 것 먼저

; CONSOLIDATION으로 들어온 주문을 배송하려 했더니 경고창 나왔음 (ERROR MESSAGE 폴더에 저장했음)

;21. Gui를 여러개 만들기 - Gui
; http://jogamja.tistory.com/m/21




	OpenFGforNewFGProcessing(CurrentPONumber, FGServer){
		
		; 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기
		;WinClose, http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber% - Internet Explorer
		WinClose, http://vendoradmin.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber% - Internet Explorer

		Clipboard :=
		
		
		Loginname = customer3
		Password = Jo123456789
		;URL = http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber%
		URL = http://vendoradmin.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber%

		WB := ComObjCreate("InternetExplorer.Application")
		WB.Visible := true
		WB.Navigate(URL)
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		   Sleep, 100



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
		if(RegExMatch(sURL, "imU)login")){
			wb.document.getElementById("tbUserID").value := Loginname  ;ID 입력
			wb.document.getElementById("tbPassword").value := Password ; 비밀번호 입력
			wb.document.getElementsByTagName("A")[0].Click() ; 로그인 버튼 누르기

			While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			   Sleep, 100
		}



		; Customer notes 읽어오기
		;CustomerNoteOnWeb := wb.document.getElementsById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderItemDetail_tbCustomerNote").value
		CustomerNoteOnWebVal := wb.document.getElementsByTagName("TEXTAREA")[2].innerText
		StringUpper, CustomerNoteOnWebVal, CustomerNoteOnWebVal ; 고객 메모 대문자로 바꾸기
;		MsgBox, % CustomerNoteOnWebVal
		
		
		; Staff only notes 읽어오기
		;StaffOnlyNoteVal := wb.document.getElementsById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderItemDetail_tbStaffOnlyNoteVal").value
		StaffOnlyNoteVal := wb.document.getElementsByTagName("TEXTAREA")[3].innerText
		StringUpper, StaffOnlyNoteVal, StaffOnlyNoteVal ; Staff only notes 대문자로 바꾸기
		
		
/*		
		Start_Style_# = 66
		Start_Color_# = 68
		Start_BO_# = 69
		
		MsgBox, % Start_Style_#
		
		loop, 10{
			
			;%A_Index%_Style_# := wb.document.getElementsByTagName("SPAN")[Start_Style_#].innerText
			%A_Index%_Style_# := wb.document.getElementsByTagName("SPAN")[66].innerText
			MsgBox, % 1_Style_#
			
			;%A_Index%_Color_# := wb.document.getElementsByTagName("SPAN")[Start_Color_#].innerText
			%A_Index%_Color_# := wb.document.getElementsByTagName("SPAN")[68].innerText
			MsgBox, % 1_Color_#
			
			;%A_Index%_BO_# := wb.document.getElementsByTagName("SPAN")[Start_BO_#].innerText
			%A_Index%_BO_# := wb.document.getElementsByTagName("SPAN")[69].innerText
			MsgBox, % 1_BO_#			
			
		}
*/
	
	
	; Order Status 의 상태(값) 얻기
	CurrentStatusONWeb := wb.document.getElementsByTagName("SELECT")[1].value
;	MsgBox, CurrentStatusONWeb value is`n`n%CurrentStatusONWeb%
	
	
	; Order Status 가 New Orders 가 아니면 
	; Ship Today 등으로 이미 처리한 주문이므로 현재 창 닫고 함수 빠져나가기
	IfNotEqual, CurrentStatusONWeb, 1
	{
		WinClose, http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber% - Internet Explorer 
		;MsgBox, , , THIS ORDER ALREADY HAS BEEN PROCESSED.`n`n`nTHIS WINDOW WILL BE CLOSED IN 5 SECONDS, 5
		MsgBox, , , The Number is %CurrentOrderIdNumber%`nTHIS ORDER ALREADY HAS BEEN PROCESSED.
		AlreadyProcessedItem = 1
		return
	}

	; Order Status DropdownList
	wb.document.getElementsByTagName("SELECT")[1].focus()
	SendInput, {Down}
	SendInput, {Enter}


	; Move and E-mail 버튼이 생길때까지 기다리기
	While wb.document.getElementsByTagName("INPUT")[2].value = ""
		Sleep, 100




	; Click Move and E-mail Button
	wb.document.getElementsByTagName("INPUT")[2].Click()
/*
	; SEND MESSAGE 버튼이 생길때까지 기다리기
	While wb.document.getElementsByTagName("INPUT")[19].value = ""
		Sleep, 100
*/


	Sleep 1000

	; Click SEND MESSAGE Button
	wb.document.getElementsByTagName("INPUT")[19].Click()

/*
	;;;;;;;;;;;;;;;;;;;;;;;;; 이거 있어도 되고 없어도 됨. 이거 없애면 처리 속도가 조금 더 빨라짐 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 100
*/

	return
	}
	










OpenFGforNewFGProcessing_UsingIE(){
	

	; ### 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기 ###
	;WinClose, http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber% - Internet Explorer
	
	; ### 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기 ###
	; 이전 IE 창 열려있으면 자꾸 이런저런 에러 나서 아예 IE 창들을 모두 그룹으로 묶고 한꺼번에 닫고 시작
	;~ GroupAdd,ExplorerGroup, ahk_class CabinetWClass
	GroupAdd,ExplorerGroup, ahk_class IEFrame
	WinClose,ahk_group ExplorerGroup


;~ /* 이거 작동하긴 했는데 이 코드 추가한 뒤 에러가 나는 것 같음
	; Do you want to close all tabs 경고창 나오면 Close all tabs 버튼 눌러서 IE창 모두 닫기
	Sleep 200
	IfWinExist, Internet Explorer
	{
		WinActivate, Internet Explorer
		ControlSend, Button1, {Enter}, Internet Explorer
	}
*/
	; 0x800704a6 - A system shut down has already been scheduled. 에러 나는 것 처리 위해
	; #################################################################################
	; https://autohotkey.com/board/topic/97731-unknown-error/page-2
	; https://autohotkey.com/board/topic/89355-suspend-script-while-process-is-running/
	; #################################################################################
	Process , Exist, iexplore.exe ;
	while (ErrorLevel){
	suspend, on
	Process, Exist, iexplore.exe
	sleep, 1000
	}
	suspend, off

	Clipboard :=



	; New Orders 검색화면 열기
	Loginname = customer3
	Password = Jo123456789
	URL = https://vendoradmin.fashiongo.net/#/order/orders/new
	WinMaximize, ahk_class IEFrame

	WB := ComObjCreate("InternetExplorer.Application")
	WB.Visible := true
	WB.Navigate(URL)
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 1000



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
	if(RegExMatch(sURL, "imU)login")){
			
		;~ wb.document.getElementById("tbUserID").value := Loginname  ;ID 입력
		;~ wb.document.getElementsByTagName("INPUT")[0].innerText := Loginname ;ID 입력
		
		;~ wb.document.getElementById("tbPassword").value := Password ; 비밀번호 입력
		;~ wb.document.getElementsByTagName("INPUT")[1].innerText := Password ;비밀번호 입력
			
			
		; 얘가 기계로 임력하는 걸 아는건지 위와 같은 일반적인 방법으로 로그인 하려고 하면 자꾸 에러가 나서 다음과 같이 이동 후 입력하는 방법 사용
		wb.document.getElementsByTagName("INPUT")[0].focus() ;ID 입력
		SendInput, % Loginname
		Sleep 100
			
		wb.document.getElementsByTagName("INPUT")[1].focus() ;비밀번호 입력
		SendInput, % Password
		Sleep 100
			
		wb.document.getElementsByTagName("BUTTON")[0].Click() ; 로그인 버튼 누르기
			
			
		; 로그인 후 New Orders 검색화면 열기
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100

		Sleep 500
		wb.document.getElementsByTagName("I")[4].Click() ; 메뉴 바의 All Orders 의 ˅ 버튼 누르기
		wb.document.getElementsByTagName("A")[13].Click() ; All Orders 안에 있는 New Orders 버튼 누르기
			

		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
		
	}
	
	Sleep, 500

	; Serch 의 Company Name 드롭다운리스트 값을 PO Number로 바꾸기
	wb.document.getElementsByTagName("SELECT")[1].focus()
	Send, {Down}
	Send, {Enter}
	Sleep 100


	; 검색창에 PO 번호 입력하기
	wb.document.getElementsByTagName("INPUT")[0].focus()
	SendInput, % CurrentPONumber
	Send, {Enter}


	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 100


	;~ if wb.document.getElementsByTagName("A")[52].innerText
	;~ {
		;~ MsgBox, it's exist
	;~ }
	;~ text := wb.document.documentElement.outerHtml
	
	;~ FileAppend, %text%, MTR2E511F454.txt
	
	;~ MsgBox, % text
	

	Sleep 1000
	
	; PO Number 는 안에 OuterHTML 값을 포함하고 있는데 만약 그 값이 현재 PO 값과 같으면 클릭해서 들어가기		
	ValOFOuterHTML := wb.document.getElementsByTagName("A")[53].outerHtml
	if ValOFOuterHTML contains %CurrentPONumber%
	{
		wb.document.getElementsByTagName("A")[53].Click()
		
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
		Sleep 500
		ReadingAndProcessingOnFGForNewOrders() ; 클릭해서 들어간 페이지에서 처리하기 위해
		return ; 이미 위 함수에서 처리했기 때문에 밑으로 진행할 필요 없이 그냥 함수 끝내고 리턴
	}
	
	; 이유는 모르겠지만 PO Number 링크 태그 값이 ("A")[53] 이 아니라 ("A")[52] 일 때도 있기 때문에 ("A")[53] 에 값이 없으면 ("A")[52] 안의 값을 한 번 더 비교해 보기
	else{
		ValOFOuterHTML := wb.document.getElementsByTagName("A")[52].outerHtml
		if ValOFOuterHTML contains %CurrentPONumber%
		{
			wb.document.getElementsByTagName("A")[52].Click()
		
			While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
				Sleep, 100
			
			Sleep 500
			ReadingAndProcessingOnFGForNewOrders() ; 클릭해서 들어간 페이지에서 처리하기 위해
			return ; 이미 위 함수에서 처리했기 때문에 밑으로 진행할 필요 없이 그냥 함수 끝내고 리턴
		}
	}

	Sleep 1000
	
	; 만약 ("A")[53] 와 ("A")[52] 를 검색해 봤는데 안에 OuterHTML 값이 없으면 PO 번호로 검색했을 때 아예 검색이 안 된 것으로 이미 처리된 주문이라는 뜻
	; Ship Today 등으로 이미 처리한 주문이므로 현재 창 닫고 함수 빠져나가기
	ValOFOuterHTML := wb.document.getElementsByTagName("A")[53].outerHtml
	if ValOFOuterHTML not contains %CurrentPONumber%
	{
		; Account Summary 열어서 Pending Orders 있는지 확인 후 
		; 있으면 Create Invoice 탭으로 넘어가기		
		AccountSummayrProcessingONCreateSalesOrdersSmallTab()


		; 펜딩값이 없었다면 AccountSummayrProcessingONCreateSalesOrdersSmallTab 함수에서 Create Invoce 탭으로 넘어가지 않았을테니 아래 절차 실행하기
		; 만약 펜딩값이 있었다면 이미 위의 함수에서 Create Invoce 탭으로 넘어간 상태임
		if(!PendingOrderStatus){		
			
			; BO 목록표 열기 (Customer Order + 버튼 클릭)
			ControlClick, WindowsForms10.BUTTON.app.0.378734a4, Accounts Summary
			WinWaitActive, Customer Order +Zoom In
			WinMaximize


			; QOH 더블클릭해서 오름차순으로 정열하기
			Text:="|<QOH>*161$24.T3sVlaAVVYAVUY4VUY4zUY4VVYAVlaAVT3sV20003U00U"
			if ok:=FindText(639,99,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Click
				Click
			}
		}


		;~ MsgBox, 262144, Already Processed Order, THIS ORDER ALREADY HAS BEEN PROCESSED.`n`n`nTHIS WINDOW WILL BE CLOSED IN 3 SECONDS, 3
		MsgBox, 4100, Already Processed Order, The Number is %CurrentOrderIdNumber%`nTHIS ORDER MIGHT BE ALREADY PROCESSED.`n(CHECKED ("A")[53])`n`n`nWOULD YOU LIKE TO MAKE AN INVOICE?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%

		; Yes 눌렀으면 인쇄하기
		IfMsgBox, Yes
		{
			PrintInvoiceInTheCaseOfAlreeadyProcessedOrder()
		}

		WinClose, Customer Order +Zoom In		
		WinClose, Accounts Summary
		WinClose, FashionGo Vendor Admin - Internet Explorer
		
		; 이미 처리된 주문이라는 표시 해주기 위해 AlreadyProcessedItem 값을 1로 만듬
		AlreadyProcessedItem = 1
		
		return		


/* ############################# 이건 나중에 만들어보자. ##############################################
		MsgBox, 4100, Already Processed Order, The Number is %CurrentOrderIdNumber%`nTHIS ORDER MIGHT BE ALREADY PROCESSED.`n(CHECKED ("A")[53])`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
		
		; No 눌렀으면 다음 주문으로 이동
		IfMsgBox, No
		{
			; 펜딩 오더가 있어서 Create Invoice 탭을 열고 Account Summary 창을 열어봤을 경우엔 일단 그 창들 닫고 처리해야 되니까
			IfWinExist, Customer Order +Zoom In
				WinClose
			
			IfWinExist, Accounts Summary
				WinClose
			
			WinClose, FashionGo Vendor Admin - Internet Explorer
			
			; 이미 처리된 주문이라는 표시 해주기 위해 AlreadyProcessedItem 값을 1로 만듬
			AlreadyProcessedItem = 1
			
			return
		}
*/

		
	}
	else{
			
			ValOFOuterHTML := wb.document.getElementsByTagName("A")[52].outerHtml
			if ValOFOuterHTML not contains %CurrentPONumber%
			{
		
			; Create Sales Orders Small 열기
			OpenCreateSalesOrdersSmallTab()
			
			;Hide All 클릭해서 메뉴 바 없애기
			ClickAtThePoint(213, 65)

			;Account Summary 열기
			MouseClick, l, 920, 155
			WinWaitActive, Accounts Summary
			WinMaximize
			
			;~ MsgBox, 262144, Already Processed Order, THIS ORDER ALREADY HAS BEEN PROCESSED.`n`n`nTHIS WINDOW WILL BE CLOSED IN 3 SECONDS, 3
			MsgBox, 262144, Already Processed Order, The Number is %CurrentOrderIdNumber%`nTHIS ORDER MIGHT BE ALREADY PROCESSED.`n(CHECKED ("A")[52])`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%
			WinClose, FashionGo Vendor Admin - Internet Explorer		
			WinClose, Accounts Summary
			
			AlreadyProcessedItem = 1
			return
		}			
	}	


	return
}



; 단독 PO 번호로 띄워진 뉴오더 페이지에서 Status 바꾸고 정보 읽는 등 처리하기
ReadingAndProcessingOnFGForNewOrders(){
	Sleep 100

	; 이미 열려있는 Virtual Terminal 의 IE 창 사용하기
	; 윈도우 타이틀 복잡할 줄 알았는데 아래처름 드릅게 단순함
	wb := IEGet("FashionGo Vendor Admin - Internet Explorer")


	Sleep 200	
	; Customer notes 읽어오기
	;CustomerNoteOnWeb := wb.document.getElementsById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderItemDetail_tbCustomerNote").value
	CustomerNoteOnWebVal := wb.document.getElementsByTagName("TEXTAREA")[0].innerText
	;~ CustomerNoteOnWebVal := wb.document.getElementsByTagName("TEXTAREA")[0].value
	StringUpper, CustomerNoteOnWebVal, CustomerNoteOnWebVal ; 고객 메모 대문자로 바꾸기

	; CustomerNoteOnWebVal 안에 있는 메모 내용 CustomerNoteOnWebVal.txt 파일에 저장하기
	FileAppend, %CustomerNoteOnWebVal%, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt
	

	Sleep 100
	; Staff only notes 읽어오기
	StaffOnlyNoteVal := wb.document.getElementsByTagName("TEXTAREA")[2].innerText
	StringUpper, StaffOnlyNoteVal, StaffOnlyNoteVal ; Staff only notes 대문자로 바꾸기
	
	; StaffOnlyNoteVal 안에 있는 메모 내용 StaffOnlyNoteVal.txt 파일에 저장하기
	FileAppend, %StaffOnlyNoteVal%, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt


	Sleep 100
	; Additional Info 읽어오기
	AdditionalInfo := wb.document.getElementsByTagName("TEXTAREA")[1].innerText
	StringUpper, AdditionalInfo, AdditionalInfo ; Staff only notes 대문자로 바꾸기
	
	; AdditionalInfo 안에 있는 메모 내용 AdditionalInfo.txt 파일에 저장하기
	FileAppend, %AdditionalInfo%, %A_ScriptDir%\CreatedFiles\AdditionalInfo.txt
	

	
	;~ MsgBox, %CustomerNoteOnWebVal%`n`n%StaffOnlyNoteVal%




	
	; Order Status 드롭다운리스트를 기본값인 New Orders 에서 Confirmed Order 로 바꾸기
	Sleep 100
	wb.document.getElementsByTagName("SELECT")[0].focus()
	Send, {Down}
	Send, {Enter}


	Sleep 100
	; Order Status 의 상태(값) 얻기
	CurrentStatusONWeb := wb.document.getElementsByTagName("SELECT")[0].value
	;~ MsgBox, CurrentStatusONWeb value is`n`n%CurrentStatusONWeb%
	
	
	Sleep 100
	; 혹시 Order Status 가 Confirmed Order 로 안 바뀌었으면( New Orders 값은 1, Confirmed Orders 는 2)
	; 재귀함수호출로 처음부터 다시 시작해서 값 바꾸기
	IfEqual, CurrentStatusONWeb, 1
		ReadingAndProcessingOnFGForNewOrders()
	

	; Update & Notify Buyer 버튼 클릭
	wb.document.getElementsByTagName("BUTTON")[2].Click()


	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 100


	; Send Mail 버튼 클릭
	;~ wb.document.getElementsByTagName("BUTTON")[11].Click()
	wb.document.getElementsByTagName("BUTTON")[11].focus()
	Sleep 100
	Send, {Enter}
	



/*
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 100

	; 커서 상태가 작업처리중이면 끝날때까지 기다리기
	while (A_cursor = "Wait")
		sleep 200
*/

	return
}









	
;~ GetActiveBrowserURL.ahk 파일에 들어있는 함수 가져온 것
/*
GetActiveBrowserURL() {
;	WinActivate, ahk_exe iexplore.exe
	
	global ModernBrowsers, LegacyBrowsers
	WinGetClass, sClass, A
	If sClass In % ModernBrowsers
		Return GetBrowserURL_ACC(sClass)
	Else If sClass In % LegacyBrowsers
		Return GetBrowserURL_DDE(sClass) ; empty string if DDE not supported (or not a browser)
	Else
		Return ""
}


; "GetBrowserURL_DDE" adapted from DDE code by Sean, (AHK_L version by maraskan_user)
; Found at http://autohotkey.com/board/topic/17633-/?p=434518

GetBrowserURL_DDE(sClass) {
	WinGet, sServer, ProcessName, % "ahk_class " sClass
	StringTrimRight, sServer, sServer, 4
	iCodePage := A_IsUnicode ? 0x04B0 : 0x03EC ; 0x04B0 = CP_WINUNICODE, 0x03EC = CP_WINANSI
	DllCall("DdeInitialize", "UPtrP", idInst, "Uint", 0, "Uint", 0, "Uint", 0)
	hServer := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", sServer, "int", iCodePage)
	hTopic := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "WWW_GetWindowInfo", "int", iCodePage)
	hItem := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "0xFFFFFFFF", "int", iCodePage)
	hConv := DllCall("DdeConnect", "UPtr", idInst, "UPtr", hServer, "UPtr", hTopic, "Uint", 0)
	hData := DllCall("DdeClientTransaction", "Uint", 0, "Uint", 0, "UPtr", hConv, "UPtr", hItem, "UInt", 1, "Uint", 0x20B0, "Uint", 10000, "UPtrP", nResult) ; 0x20B0 = XTYP_REQUEST, 10000 = 10s timeout
	sData := DllCall("DdeAccessData", "Uint", hData, "Uint", 0, "Str")
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hServer)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hTopic)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hItem)
	DllCall("DdeUnaccessData", "UPtr", hData)
	DllCall("DdeFreeDataHandle", "UPtr", hData)
	DllCall("DdeDisconnect", "UPtr", hConv)
	DllCall("DdeUninitialize", "UPtr", idInst)
	csvWindowInfo := StrGet(&sData, "CP0")
	StringSplit, sWindowInfo, csvWindowInfo, `" ;"; comment to avoid a syntax highlighting issue in autohotkey.com/boards
	Return sWindowInfo2
}

GetBrowserURL_ACC(sClass) {
	global nWindow, accAddressBar
	If (nWindow != WinExist("ahk_class " sClass)) ; reuses accAddressBar if it's the same window
	{
		nWindow := WinExist("ahk_class " sClass)
		accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindow))
	}
	Try sURL := accAddressBar.accValue(0)
	If (sURL == "") {
		WinGet, nWindows, List, % "ahk_class " sClass ; In case of a nested browser window as in the old CoolNovo (TO DO: check if still needed)
		If (nWindows > 1) {
			accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindows2))
			Try sURL := accAddressBar.accValue(0)
		}
	}
	If ((sURL != "") and (SubStr(sURL, 1, 4) != "http")) ; Modern browsers omit "http://"
		sURL := "http://" sURL
	If (sURL == "")
		nWindow := -1 ; Don't remember the window if there is no URL
	Return sURL
}

; "GetAddressBar" based in code by uname
; Found at http://autohotkey.com/board/topic/103178-/?p=637687

GetAddressBar(accObj) {
	Try If ((accObj.accRole(0) == 42) and IsURL(accObj.accValue(0)))
		Return accObj
	Try If ((accObj.accRole(0) == 42) and IsURL("http://" accObj.accValue(0))) ; Modern browsers omit "http://"
		Return accObj
	For nChild, accChild in Acc_Children(accObj)
		If IsObject(accAddressBar := GetAddressBar(accChild))
			Return accAddressBar
}

IsURL(sURL) {
	Return RegExMatch(sURL, "^(?<Protocol>https?|ftp)://(?<Domain>(?:[\w-]+\.)+\w\w+)(?::(?<Port>\d+))?/?(?<Path>(?:[^:/?# ]*/?)+)(?:\?(?<Query>[^#]+)?)?(?:\#(?<Hash>.+)?)?$")
}

; The code below is part of the Acc.ahk Standard Library by Sean (updated by jethrow)
; Found at http://autohotkey.com/board/topic/77303-/?p=491516

Acc_Init()
{
	static h
	If Not h
		h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}
Acc_ObjectFromWindow(hWnd, idObject = 0)
{
	Acc_Init()
	If DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
	Return ComObjEnwrap(9,pacc,1)
}
Acc_Query(Acc) {
	Try Return ComObj(9, ComObjQuery(Acc,"{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
}
Acc_Children(Acc) {
	If ComObjType(Acc,"Name") != "IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	Else {
		Acc_Init(), cChildren:=Acc.accChildCount, Children:=[]
		If DllCall("oleacc\AccessibleChildren", "Ptr",ComObjValue(Acc), "Int",0, "Int",cChildren, "Ptr",VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*",cChildren)=0 {
			Loop %cChildren%
				i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i), Children.Insert(NumGet(varChildren,i-8)=9?Acc_Query(child):child), NumGet(varChildren,i-8)=9?ObjRelease(child):
			Return Children.MaxIndex()?Children:
		} Else
			ErrorLevel := "AccessibleChildren DllCall Failed"
	}
}
*/
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
OpenFGforNewFGProcessing_NEW(){
		
	Clipboard :=
		
		
	Loginname = customer3
	Password = Jo123456789
	
	
	;~ Run, https://vendoradmin.fashiongo.net/#/order/orders/new;pn1=1;ps1=50;sb1=checkoutDateDesc;df=08%2F19%2F2017;dt=09%2F18%2F2017;so=po;sq=%CurrentPONumber%
	
	;~ https://vendoradmin.fashiongo.net/#/order/orders/new;pn1=1;ps1=50;sb1=checkoutDateDesc;df=08%2F19%2F2017;dt=09%2F18%2F2017;so=po;sq=MTR2DCF78043
	;~ https://vendoradmin.fashiongo.net/#/order/orders/new;pn1=1;ps1=50;sb1=checkoutDateDesc;df=082F2017;dt=092F2017;so=po;sq=MTR2DCF78043
	
	Run, https://vendoradmin.fashiongo.net/#/order/12297142

	
	ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber.png
	FoundX += 5
	FoundY += 5
	MouseClick, l, %FoundX%, %FoundY%
	
	MsgBox

	return
	
	
	}
	
	
OpenFGforNewFGProcessing_CHROME(){
	
	; 만약 PO옆의 TAB버튼 못 찾으면 다시 시작하기 위해
	RESTART:

	FoundTheTabButton :=

	URL = https://vendoradmin.fashiongo.net/#/order/orders/new
	run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL
	WinMaximize, ahk_class Chrome_WidgetWin_1


	; 커서 상태가 Wait(로딩중)인 동안 기다리기
	while (A_cursor = "Wait")
		sleep 200
	
	WinActivate, FashionGo Vendor Admin - Google Chrome

	; Company Name 다운드롭리스트 찾아서 PO Number 로 바꾸기
	loop{
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder00]CompanyName.png
		if(ErrorLevel = 0){
			FoundX += 2
			FoundY += 2
			MouseClick, l, %FoundX%, %FoundY%
			Send, {Down}
			Send, {Enter}
			sleep 200

			
			break
		}
	}
	
	
	while (A_cursor = "Wait")
		sleep 200

	WinActivate, FashionGo Vendor Admin - Google Chrome
	
	
	; 상태가 PO NUMBER로 바뀌었는지 찾기. 찾으면 루프 끝내고 나가기
	loop, 20{		
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder00]PO_Number.png
		if(ErrorLevel = 0){			
			break
		}
	}
	
	Sleep 1000
	; 만약 상태가 컨펌으로 바뀌지 않았으면(그림 못 찾았으면)
	; 창 닫고 그냥 처음부터 다시 시작
	if(ErrorLevel = 1){
		MsgBox, , Memo, The Number is %CurrentOrderIdNumber%`nIT'S AN ERROR OCCURRED. PLEASE CLICK Ok TO RESTART
		Reload
		
		;~ WinClose, ahk_class Chrome_WidgetWin_1
		;~ TheError = 1
		;~ return

	}
	
	




	while (A_cursor = "Wait")
		sleep 200

	; 검색란에 PO 번호 입력하기
	loop{
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder00]SerchBlank.png
		if(ErrorLevel = 0){
			FoundX += 5
			MouseClick, l, %FoundX%, %FoundY%
			send, ^a
			SendInput, %CurrentPONumber%
			SendInput, {Enter}
			sleep 200

			
			break
		}
	}





	while (A_cursor = "Wait")
		sleep 200


/*	
	; PO 새탭으로 열기 이미 ship today등으로 처리된 주문은 PO가 뜨지 않으니 무한대가 아닌 30번 정도만 찾아보기
	loop, 5{
		
		Text:="|<Tab Button>*144$12.TTU7U7U9UFUUV1W1U1U1U1zzU"
		;~ Text:="|<>*159$12.yTU7UDUPUnVUX3a3Y3U3U3zyU"
		
		if ok:=FindText(889,449,150000,150000,0,0,Text)
		;~ if ok:=FindText(2810,447,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			
			; PO 를 찾았으면 변수값을 1로 설정하기
			FoundTheTabButton = 1
			  
			break
		}
	}

	; 만약 TAB버튼 못 찾았으면 다른 코드로 버튼을 찾고 그래도 없으면
	; 메세지 띄워서 다시 검색해볼건지 다음으로 넘어갈 지 묻고 No 누르면 다음 주문으로 넘어가기
	outer:
	IfNotEqual, FoundTheTabButton, 1
	;~ if(ErrorLevel = 1)
	{
		loop, 5{
			
			Text:="|<Tab Button 2>*159$12.yTU7UDUPUnVUX3a3Y3U3U3zyU"
			;~ Text:="|<Tab Button>*144$12.TTU7U7U9UFUUV1W1U1U1U1zzU"

			if ok:=FindText(2810,447,150000,150000,0,0,Text)
			;~ if ok:=FindText(889,449,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Click
				
				; PO 를 찾았으면 변수값을 1로 설정하기
				FoundTheTabButton = 1
				  
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}
		
		loop, 5{
						
			Text:="|<>*152$12.TDk7U7UBUNUkVVX1W1U1U1zzU"

			if ok:=FindText(890,446,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Click
				
				; PO 를 찾았으면 변수값을 1로 설정하기
				FoundTheTabButton = 1
				  
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
		
		MsgBox, 4100, Memo, The Number is %CurrentOrderIdNumber%`nIT MIGHT BE AN ORDER ALREADY PROCESSED`nWOULD YOU LIKE TO TRY WEB PROCESS AGAIN?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER
		IfMsgBox, No
		{
			; 이 함수 밖에서 AlreadyProcessedItem 값이 1이면 다음 주문으로 넘어가는 처리 해주기 위해 값을 1로 설정하기
			AlreadyProcessedItem = 1
			return
		}
		IfMsgBox, Yes
		{
			WinClose, ahk_class Chrome_WidgetWin_1
			GOTO, RESTART
		}
	}

break_outer: ; Tab 버튼 찾았으면 이리로 나옴
*/

/*
		MsgBox, 4100, Memo, The Number is %CurrentOrderIdNumber%`n`nIF YOU CLICK No, MOVE TO THE NEXT ORDER
		IfMsgBox, No
		{
			AlreadyProcessedItem = 1
			return
		}
*/



	; PO 새탭으로 열기
	Sleep 200
	loop, 10{
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber.png
		;~ ImageSearch, FoundX, FoundY, 655, 403, 1046, 503, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber.png
		if(ErrorLevel = 0){
			FoundX += 2
			FoundY += 2
			MouseClick, l, %FoundX%, %FoundY%
			sleep 200
			
			break
		}
	}
	; 만약 TAB버튼 못 찾았으면 흰색 배경의 버튼을 찾고 그래도 없으면 계속 찾다가 없으면
	; 메세지 띄워서 Yes 누르면 그냥 진행하고 No 누르면 다음 주문으로 넘어가기
	if(ErrorLevel = 1){
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_WH.png
			;~ ImageSearch, FoundX, FoundY, 655, 403, 1046, 503, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_WH.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_03.png
			;~ ImageSearch, FoundX, FoundY, 655, 403, 1046, 503, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_03.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_04.png
			;~ ImageSearch, FoundX, FoundY, 655, 403, 1046, 503, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_04.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_05.png
			;~ ImageSearch, FoundX, FoundY, 655, 403, 1046, 503, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_04.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_06.png
			;~ ImageSearch, FoundX, FoundY, 655, 403, 1046, 503, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_04.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_07.png
			;~ ImageSearch, FoundX, FoundY, 655, 403, 1046, 503, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_04.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_08.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		
			
		loop, 2{
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber_09.png
			if(ErrorLevel = 0){
				FoundX += 2
				FoundY += 2
				MouseClick, l, %FoundX%, %FoundY%
				sleep 200
				
				goto break_outer ; 찾았으면 아예 밖으로 나오기 위해
			}
		}		


		MsgBox, 4100, Memo, The Number is %CurrentOrderIdNumber%`n`nTHIS ORDER MIGHT HAS BEEN PROCESSED ALREADY AS CONSOLIDATION OR SHIP TODAY. IF NOT, IT'S AN ERROR OCCURRED`n`nIF YOU'D LIKE TO RELOAD THE PROGRAM, PLEASE CLICK Yes BUTTON.`nOR CLICK No BUTTON TO MOVE TO THE NEXT ORDER
		WinActivate, ahk_class Chrome_WidgetWin_1
		IfMsgBox, No
		{
			AlreadyProcessedItem = 1
			return
		}
		
		IfMsgBox, Yes
			Reload
		;~ {
			;~ WinClose, ahk_class Chrome_WidgetWin_1
			;~ GOTO, RESTART
		;~ }
	}
	break_outer: ; Tab 버튼 찾았으면 이리로 나옴
*/





	while (A_cursor = "Wait")
		sleep 200
/*
	; Order Status 바꿔서 선택하기
	loop{
		Text:="|<New Orders Donwdrop>*196$70.UU0007U0800320001X00U00A8QV444wyCDScW9AEEP4N4nCGEYm10cEcGA99yGc42V2z8QWY1GUEO4+0UC6E3C118Ec20MMWAE6AV6F8lUVkV07W3ssVu"

		if ok:=FindText(2385,297,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Send, {Down}
			Send, {Enter}
			
			sleep 200
			break
		}
	}
*/		

	; Order Status 바꿔서 선택하기
	loop{
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder02]NEW_ORDERS_DOWNDROP_LIST.png
		if(ErrorLevel = 0){
			FoundX += 2
			FoundY += 2
			MouseClick, l, %FoundX%, %FoundY%
			Send, {Down}
			Send, {Enter}
			sleep 200
			
			break
		}
	}
*/





	while (A_cursor = "Wait")
		sleep 200
/*
	; Update & Notify Buyer 버튼 클릭하기
	loop{
		Text:="|<Update Notify Buyer>*191$137.00060000000001XU0000001V00A0A00w0kU0nA01y00003200M0M01g1V01UM03C000064y7ltts6M3W77hvA6AXNXXw9qNalaM7U7YP6NaMAN6nBbMH4n0nAkC0/9XAnAkTWBalAka96DaTUQUHH6NaB0lYNdyFVAGAnAU1h0XaAnAS1X8nn0X2MaN6NU2C17ANaMw36Fba17QnAnQng6Q26BXAkk6Qr66m3lwDXtnk7w4AC7NVUDkyA7403000000000000030000M00060000000000000A0001U000A0000000000000k0006002"

		if ok:=FindText(2721,298,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			
			sleep 200
			break
		}
	}
*/

	; Update & Notify Buyer 버튼 클릭하기
	loop{
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder03]UPDATE_NOTIFY_BUYER_BUTTON.png
		if(ErrorLevel = 0){
			FoundX += 2
			FoundY += 2
			MouseClick, l, %FoundX%, %FoundY%
			sleep 200
			
			break
		}
	}
*/





	while (A_cursor = "Wait")
		sleep 200
/*
	; SEND EMAIL 버튼 클릭하기
	loop{
		Text:="|<Send Mail>*181$63.zzzzzzzzzzzzzzzlzzzyNkTzzyD7lznA1zzzlsSDztb661kD3VUlATUE41sQ8W9kMmMaD11wFD10H4ls8A29z0yMaD094FAs7n4lsV8W9U42MUD4N0FC1kn61sXA29zzzzzzzzzzw"

		if ok:=FindText(3259,798,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			
			sleep 200
			break
		}
	}
*/

	; SEND EMAIL 버튼 클릭하기
	Sleep 500
	loop{
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[FGNewOrder04]SEND_EMAIL.png
		if(ErrorLevel = 0){
			FoundX += 2
			FoundY += 2
			MouseClick, l, %FoundX%, %FoundY%
			
			break
		}
	}
*/	





	;~ WinActivate, ahk_class Chrome_WidgetWin_1
	
	
	return	
}






OpenFGforNewFGProcessing_UsingChromeBySelenium(){
	
	
	Clipboard :=
	sURL :=	
		
	; New Orders 검색화면 열기
	Loginname = customer3
	Password = Jo123456789
	URL = https://vendoradmin.fashiongo.net/#/order/orders/new


	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.Get("https://www.daum.net/")
	driver.findElementsByTagName("INPUT")[2].SendKeys("1234")
	;~ wb.document.getElementsByTagName("INPUT")[0].focus() ;ID 입력
	
	
	
	
/*	
	driver.Get("https://vendoradmin.fashiongo.net/#/order/orders/new")
	WinMaximize, ahk_class Chrome_WidgetWin_1


	; 현재 url 얻는 부분
	Send, !d
	Sleep 100
	Send, ^c
	
	Sleep 700
	sURL := % Clipboard
	

	; 얻은 현재 url이 로그인 화면이면 로그인 하기
	if(RegExMatch(sURL, "imU)login")){
		driver.findElementsByName("INPUT")[0].SendKeys(Loginname)
		;~ wb.document.getElementsByTagName("INPUT")[0].focus() ;ID 입력
	}
			
	
	MsgBox, %sURL%
*/	
	
	return
}


00OpenFGforNewFGProcessing_UsingChromeBySelenium(){
	

	; ### 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기 ###
	;WinClose, http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber% - Internet Explorer
	
	; ### 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기 ###
	; 이전 IE 창 열려있으면 자꾸 이런저런 에러 나서 아예 IE 창들을 모두 그룹으로 묶고 한꺼번에 닫고 시작
	;~ GroupAdd,ExplorerGroup, ahk_class CabinetWClass
;	GroupAdd,ExplorerGroup, ahk_class IEFrame
;	WinClose,ahk_group ExplorerGroup


/*
	; 0x800704a6 - A system shut down has already been scheduled. 에러 나는 것 처리 위해
	; #################################################################################
	; https://autohotkey.com/board/topic/97731-unknown-error/page-2
	; https://autohotkey.com/board/topic/89355-suspend-script-while-process-is-running/
	; #################################################################################
	Process , Exist, iexplore.exe ;
	while (ErrorLevel){
	suspend, on
	Process, Exist, iexplore.exe
	sleep, 1000
	}
	suspend, off

	Clipboard :=
*/

		
	; New Orders 검색화면 열기
	Loginname = customer3
	Password = Jo123456789
	URL = https://vendoradmin.fashiongo.net/#/order/orders/new


	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.Get("https://vendoradmin.fashiongo.net/#/order/orders/new")
	WinMaximize, ahk_class Chrome_WidgetWin_1


	; 현재 url 얻는 부분
	Send, !d
	Send, ^c
	
	Sleep 700
	sURL := % Clipboard

	; 얻은 현재 url이 로그인 화면이면 로그인 하기
	if(RegExMatch(sURL, "imU)login")){
		
		driver.findElementsByName("INPUT").item[0].SendKeys(Loginname)
			
		;~ wb.document.getElementById("tbUserID").value := Loginname  ;ID 입력
		;~ wb.document.getElementsByTagName("INPUT")[0].innerText := Loginname ;ID 입력
		
		;~ wb.document.getElementById("tbPassword").value := Password ; 비밀번호 입력
		;~ wb.document.getElementsByTagName("INPUT")[1].innerText := Password ;비밀번호 입력
			
			
		; 얘가 기계로 임력하는 걸 아는건지 위와 같은 일반적인 방법으로 로그인 하려고 하면 자꾸 에러가 나서 다음과 같이 이동 후 입력하는 방법 사용
		wb.document.getElementsByTagName("INPUT")[0].focus() ;ID 입력
		SendInput, % Loginname
		Sleep 100
			
		wb.document.getElementsByTagName("INPUT")[1].focus() ;비밀번호 입력
		SendInput, % Password
		Sleep 100
			
		wb.document.getElementsByTagName("BUTTON")[0].Click() ; 로그인 버튼 누르기
			
			
		; 로그인 후 New Orders 검색화면 열기
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100

		Sleep 500
		wb.document.getElementsByTagName("I")[4].Click() ; 메뉴 바의 All Orders 의 ˅ 버튼 누르기
		wb.document.getElementsByTagName("A")[13].Click() ; All Orders 안에 있는 New Orders 버튼 누르기
			

		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
		
	}


	; Serch 의 Company Name 드롭다운리스트 값을 PO Number로 바꾸기
	wb.document.getElementsByTagName("SELECT")[1].focus()
	Send, {Down}
	Send, {Enter}
	Sleep 100


	; 검색창에 PO 번호 입력하기
	wb.document.getElementsByTagName("INPUT")[0].focus()
	SendInput, % CurrentPONumber
	Send, {Enter}


	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 100


	;~ if wb.document.getElementsByTagName("A")[52].innerText
	;~ {
		;~ MsgBox, it's exist
	;~ }
	;~ text := wb.document.documentElement.outerHtml
	
	;~ FileAppend, %text%, MTR2E511F454.txt
	
	;~ MsgBox, % text
	

	Sleep 1000
	
	; PO Number 는 안에 OuterHTML 값을 포함하고 있는데 만약 그 값이 현재 PO 값과 같으면 클릭해서 들어가기		
	ValOFOuterHTML := wb.document.getElementsByTagName("A")[53].outerHtml
	if ValOFOuterHTML contains %CurrentPONumber%
	{
		wb.document.getElementsByTagName("A")[53].Click()
		
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
		Sleep 500
		ReadingAndProcessingOnFGForNewOrders() ; 클릭해서 들어간 페이지에서 처리하기 위해
		return ; 이미 위 함수에서 처리했기 때문에 밑으로 진행할 필요 없이 그냥 함수 끝내고 리턴
	}
	
	; 이유는 모르겠지만 PO Number 링크 태그 값이 ("A")[53] 이 아니라 ("A")[52] 일 때도 있기 때문에 ("A")[53] 에 값이 없으면 ("A")[52] 안의 값을 한 번 더 비교해 보기
	else{
		ValOFOuterHTML := wb.document.getElementsByTagName("A")[52].outerHtml
		if ValOFOuterHTML contains %CurrentPONumber%
		{
			wb.document.getElementsByTagName("A")[52].Click()
		
			While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
				Sleep, 100
			
			Sleep 500
			ReadingAndProcessingOnFGForNewOrders() ; 클릭해서 들어간 페이지에서 처리하기 위해
			return ; 이미 위 함수에서 처리했기 때문에 밑으로 진행할 필요 없이 그냥 함수 끝내고 리턴
		}
	}

	Sleep 1000
	
	; 만약 ("A")[53] 와 ("A")[52] 를 검색해 봤는데 안에 OuterHTML 값이 없으면 PO 번호로 검색했을 때 아예 검색이 안 된 것으로 이미 처리된 주문이라는 뜻
	; Ship Today 등으로 이미 처리한 주문이므로 현재 창 닫고 함수 빠져나가기
	ValOFOuterHTML := wb.document.getElementsByTagName("A")[53].outerHtml
	if ValOFOuterHTML not contains %CurrentPONumber%
	{
	
		; Create Sales Orders Small 열기
		OpenCreateSalesOrdersSmallTab()
		
		;Hide All 클릭해서 메뉴 바 없애기
		ClickAtThePoint(213, 65)

		;Account Summary 열기
		MouseClick, l, 920, 155
		WinWaitActive, Accounts Summary
		WinMaximize
		
		; BO 목록표 열기 (Customer Order + 버튼 클릭)
		ControlClick, WindowsForms10.BUTTON.app.0.378734a4, Accounts Summary
		WinWaitActive, Customer Order +Zoom In
		WinMaximize
			
		
		;~ MsgBox, 262144, Already Processed Order, THIS ORDER ALREADY HAS BEEN PROCESSED.`n`n`nTHIS WINDOW WILL BE CLOSED IN 3 SECONDS, 3
		MsgBox, 262144, Already Processed Order, The Number is %CurrentOrderIdNumber%`nTHIS ORDER MIGHT BE ALREADY PROCESSED.`n(CHECKED ("A")[53])`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%
		
		WinClose, Customer Order +Zoom In		
		WinClose, Accounts Summary
		WinClose, FashionGo Vendor Admin - Internet Explorer
		
		; 이미 처리된 주문이라는 표시 해주기 위해 AlreadyProcessedItem 값을 1로 만듬
		AlreadyProcessedItem = 1
		
		return		


/* ############################# 이건 나중에 만들어보자. ##############################################
		MsgBox, 4100, Already Processed Order, The Number is %CurrentOrderIdNumber%`nTHIS ORDER MIGHT BE ALREADY PROCESSED.`n(CHECKED ("A")[53])`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
		
		; No 눌렀으면 다음 주문으로 이동
		IfMsgBox, No
		{
			; 펜딩 오더가 있어서 Create Invoice 탭을 열고 Account Summary 창을 열어봤을 경우엔 일단 그 창들 닫고 처리해야 되니까
			IfWinExist, Customer Order +Zoom In
				WinClose
			
			IfWinExist, Accounts Summary
				WinClose
			
			WinClose, FashionGo Vendor Admin - Internet Explorer
			
			; 이미 처리된 주문이라는 표시 해주기 위해 AlreadyProcessedItem 값을 1로 만듬
			AlreadyProcessedItem = 1
			
			return
		}
*/

		
	}
	else{
			
			ValOFOuterHTML := wb.document.getElementsByTagName("A")[52].outerHtml
			if ValOFOuterHTML not contains %CurrentPONumber%
			{
		
			; Create Sales Orders Small 열기
			OpenCreateSalesOrdersSmallTab()
			
			;Hide All 클릭해서 메뉴 바 없애기
			ClickAtThePoint(213, 65)

			;Account Summary 열기
			MouseClick, l, 920, 155
			WinWaitActive, Accounts Summary
			WinMaximize
			
			;~ MsgBox, 262144, Already Processed Order, THIS ORDER ALREADY HAS BEEN PROCESSED.`n`n`nTHIS WINDOW WILL BE CLOSED IN 3 SECONDS, 3
			MsgBox, 262144, Already Processed Order, The Number is %CurrentOrderIdNumber%`nTHIS ORDER MIGHT BE ALREADY PROCESSED.`n(CHECKED ("A")[52])`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%
			WinClose, FashionGo Vendor Admin - Internet Explorer		
			WinClose, Accounts Summary
			
			AlreadyProcessedItem = 1
			return
		}			
	}	


	return
}



; 단독 PO 번호로 띄워진 뉴오더 페이지에서 Status 바꾸고 정보 읽는 등 처리하기
00ReadingAndProcessingOnFGForNewOrders(){
	Sleep 100

	; 이미 열려있는 Virtual Terminal 의 IE 창 사용하기
	; 윈도우 타이틀 복잡할 줄 알았는데 아래처름 드릅게 단순함
	wb := IEGet("FashionGo Vendor Admin - Internet Explorer")


	Sleep 200	
	; Customer notes 읽어오기
	;CustomerNoteOnWeb := wb.document.getElementsById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderItemDetail_tbCustomerNote").value
	CustomerNoteOnWebVal := wb.document.getElementsByTagName("TEXTAREA")[0].innerText
	;~ CustomerNoteOnWebVal := wb.document.getElementsByTagName("TEXTAREA")[0].value
	StringUpper, CustomerNoteOnWebVal, CustomerNoteOnWebVal ; 고객 메모 대문자로 바꾸기	

	Sleep 100
	; Staff only notes 읽어오기
	StaffOnlyNoteVal := wb.document.getElementsByTagName("TEXTAREA")[2].innerText
	StringUpper, StaffOnlyNoteVal, StaffOnlyNoteVal ; Staff only notes 대문자로 바꾸기
	

	Sleep 100
	; Additional Info 읽어오기
	StaffOnlyNoteVal := wb.document.getElementsByTagName("TEXTAREA")[1].innerText
	StringUpper, AdditionalInfo, AdditionalInfo ; Staff only notes 대문자로 바꾸기
		
	
	;~ MsgBox, %CustomerNoteOnWebVal%`n`n%StaffOnlyNoteVal%




	
	; Order Status 드롭다운리스트를 기본값인 New Orders 에서 Confirmed Order 로 바꾸기
	Sleep 100
	wb.document.getElementsByTagName("SELECT")[0].focus()
	Send, {Down}
	Send, {Enter}


	Sleep 100
	; Order Status 의 상태(값) 얻기
	CurrentStatusONWeb := wb.document.getElementsByTagName("SELECT")[0].value
	;~ MsgBox, CurrentStatusONWeb value is`n`n%CurrentStatusONWeb%
	
	
	Sleep 100
	; 혹시 Order Status 가 Confirmed Order 로 안 바뀌었으면( New Orders 값은 1, Confirmed Orders 는 2)
	; 재귀함수호출로 처음부터 다시 시작해서 값 바꾸기
	IfEqual, CurrentStatusONWeb, 1
		ReadingAndProcessingOnFGForNewOrders()
	

	; Update & Notify Buyer 버튼 클릭
	wb.document.getElementsByTagName("BUTTON")[2].Click()


	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 100


	; Send Mail 버튼 클릭
	;~ wb.document.getElementsByTagName("BUTTON")[11].Click()
	wb.document.getElementsByTagName("BUTTON")[11].focus()
	Sleep 100
	Send, {Enter}
	



/*
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 100

	; 커서 상태가 작업처리중이면 끝날때까지 기다리기
	while (A_cursor = "Wait")
		sleep 200
*/

	return
}









PrintInvoiceInTheCaseOfAlreeadyProcessedOrder(){
	

			; 펜딩 오더가 있어서 Create Invoice 탭을 열고 Account Summary 창을 열어봤을 경우엔 일단 그 창들 닫고 처리해야 되니까
			IfWinExist, Customer Order +Zoom In
				WinClose
			
			IfWinExist, Accounts Summary
				WinClose
				
			; BO 목록표(인보이스 종이) 인쇄하기
			PrintingBOList()			

			; 목록표 인쇄 후 Create Invoce 탭으로 넘어가서 인보이스 만들 준비
			OpenCreateInvoiceTab()
			Sleep 200
								
			;New & Clear 버튼 클릭
			MouseClick, l, 60, 124, 2, 
			sleep 1000			
											
			; CompanyName 변수값 넣기
			DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
			Sleep 100
			MouseGetPos, , , , control, 1
			ControlSetText, %control%, %CompanyName%, %wintitle%
			ControlClick %control%, %wintitle%
			SendInput, {Enter}


			; Warning 창이 뜨면 Credit 있다는 것
			Sleep 2000
			IfWinActive, Warning
			{					
				MsgBox, IT'S APPLY CREDIT
				WinClose, Warning
			}				

			; Sales Orders 버튼 클릭 하기
			MouseClick, l, 232, 388
			Sleep 100
				
				
			; Please enter Company Nmae first 에러 창이 뜨면 창 닫고 CompanyName 다시 입력하기
			Sleep 500
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[ERROR]COMPANY_NAME_FIRST.png
				
			if(ErrorLevel = 0){			
				ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[ERROR]COMPANY_NAME_FIRST_OK.png
				MouseClick, l, %FoundX%, %FoundY%
				Sleep 200			

				; CompanyName 변수값 넣기
				DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
				Sleep 100
				MouseGetPos, , , , control, 1
				ControlSetText, %control%, %CompanyName%, %wintitle%
				ControlClick %control%, %wintitle%
				SendInput, {Enter}
			}


			; The data no found 에러 떴을 때. 아마 같은 회사명 가진 다른 회사가 있을 것임
			IfWinExist, Confirm (No Data)
				MsgBox, 262144, Having Same Name, IT'S AN ERROR OCCURED, PROBABLY THERE IS AN ANOTHER COMPANY HAVING SAME COMPANY NAME.`n`nPLEASE CHECK IT FIRST AND THEN CLICK OK BUTTON TO CONTINUE
				
			; 만약 Please enter 'Company Name' first 경고창이 뜨면 
			; 다시 Company Name 입력하고 Sales Orders 버튼 클릭하기
			IfWinActive, ahk_class #32770
			{
				WinClose, ahk_class #32770
	
				; CompanyName 변수값 넣기
				DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
				Sleep 100
				MouseGetPos, , , , control, 1
				ControlSetText, %control%, %CompanyName%, %wintitle%
				ControlClick %control%, %wintitle%
				SendInput, {Enter}
				
				; Sales Orders 버튼 클릭 하기
				MouseClick, l, 232, 388	
			}


			; 다음 주문으로 넘어가기 전 메세지 띄우기
			MsgBox, 262144, Memo, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nThe Shipping State is %State%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%


			WinClose, Transfer from Sales Order
				
			WinClose, ahk_exe iexplore.exe
			
	
	return
}
























































