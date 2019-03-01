
;~ #Include %A_ScriptDir%\lib\
;~ #Include FindText.ahk

	;LAMBS활성화 후 화면 초기화 하기
	Start(){
		
		;LAMBS Window 활성화 하기
		WinActivate, LAMBS -  Garment Manufacturer & Wholesale Software
		windowtitle = LAMBS -  Garment Manufacturer & Wholesale Software
		CheckTheWindowPresentAndActiveIt(windowtitle)

		;Hide All 클릭해서 메뉴 바 없애기
		ClickAtThePoint(213, 65)
		
		return
	}
	
	;LAMBS Invoice 1 탭 클릭
	Start_Invoice_1(){
		; Invoice 1 탭에서 시작하기
		MouseClick, l, 41, 265, 1
		
		return
	}
	
	
	;위치 받아서 클릭하기
	ClickAtThePoint(XPoint, YPoint){
		MouseClick, l, XPoint, YPoint, 1
		Sleep 1000
		return
	}
	
	
	;윈도우 창이 존재하는지 확인 후 활성화
	ActivateWindows(windowtitle){
		Loop,
		IfWinExist, % windowtitle
		break

		WinActivate, % windowtitle

	}
	
	;지정된 창이 존재하는지 무한정 확인 후 그 창 활성화
	CheckTheWindowPresentAndActiveIt(windowtitle){
		WinWait, % windowtitle
		WinActivate, % windowtitle
		return
	}
	
	;지정된 창이 존재할 때까지 무한정 기다리다가 창이 존재하면 그 창이 닫힐 때까지 기다리기
	CheckTheWindowPresentAndWaitUntillItClose(windowtitle){
		Sleep 500
		
		WinWait, % windowtitle
;		MsgBox, Opened`n%windowtitle%
		WinWaitClose, % windowtitle  ; WinWait가 발견한 창이 닫히기를 기다립니다.
;		MsgBox, Closed`n%windowtitle%
		;MsgBox, Transfer from Sales Order 노트패드가 이제 닫혔습니다.
		return

		/*
		Run, notepad.exe
		WinWait, Untitled - Notepad
		WinWaitClose  ; WinWait가 발견한 창이 닫히기를 기다립니다.
		MsgBox, 노트패드가 이제 닫혔습니다.
		return
		*/
	}
	
	
	; 자동 입력 방지 찾아내기
	DetectHumanVarification(jpgLocation, FoundX, FoundY, FromX, FromY, ToX, ToY){
		
		;jpgLocation = %A_ScripDir%PICTURES\HumanVarificationArrowSmallButton.png
		ImageSearch, , , FromX, FromY, ToX, ToY, %jpgLocation%
		

		if(ErrorLevel = 0){
			MsgBox, Found it
			;MsgBox, FoundX in function : %FoundX%`nFoundY in function : %FoundY%
			return 0 
		}
		else if(ErrorLevel = 1){
			MsgBox, Couldn't find it!
			return 1
			;break
		}
		else if(ErrorLevel = 2){
			MsgBox, Unexpected Error Occur!
			return 2
			;break
		}
	}

	

	;이미지 서치
	;pX, pY 값은 아무거나 입력해도 됨. 어차피 ByRef로 전역변수로 선언되어 값을 받기 위함이 아닌 이미지를 찾은 위치를 돌려주기 위한 변수
	PicSearch(jpgLocation, ByRef FoundX, ByRef FoundY, FromX, FromY, ToX, ToY){

;			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %jpgLocation%  ;%A_ScripDir%PICTURES\00New_Clear.png
			ImageSearch, FoundX, FoundY, %FromX%, %FromY%, %ToX%, %ToY%, %jpgLocation%  ;%A_ScripDir%PICTURES\00New_Clear.png

			
			if(ErrorLevel = 0){
				;MsgBox, Found it
				;MsgBox, FoundX in function : %FoundX%`nFoundY in function : %FoundY%
				return
			}
			else if(ErrorLevel = 1){
				;MsgBox, Couldn't find it!
				return
				;break
			}
			else if(ErrorLevel = 2){
				MsgBox, Unexpected Error Occur!
				return
				;break
			}
		

		;MsgBox, End of function
		return
	}

	;이미지 서치 후 값 돌려주기
	PicSearchAndReturnItsVal(jpgLocation){
		Loop, 10{
			ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %jpgLocation%  ;%A_ScripDir%PICTURES\00New_Clear.png
			
			if(ErrorLevel = 0){
				;MsgBox, Found it
				pX += 70
				pY += 5

				MouseMove, pX, pY
				MouseGetPos, , , , control
				ControlGetText, QtyofItemsOnLAMBS, %control%, LAMBS
;				MsgBox, % QtyofItemsOnLAMBS
				
				return QtyofItemsOnLAMBS
			}
			else if(ErrorLevel = 1){
				;MsgBox, Couldn't find it!
				;break
			}
			else if(ErrorLevel = 2){
				MsgBox, Unexpected Error Occur!
				;break
			}
		
		}
		MsgBox, Out!
		return
	}
	
	;이미지 서치
	emailSearch(){
		jpgLocation = %A_ScripDir%PICTURES\email.png

		Loop, 5{
			ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %jpgLocation%  ;%A_ScripDir%PICTURES\00New_Clear.png
			
			if(ErrorLevel = 0){
				MsgBox, Found it
				pX += 40
				pY += 5
				Click, %pX%, %pY%
				return
			}
			else if(ErrorLevel = 1){
				;MsgBox, Couldn't find it!
				;break
			}
			else if(ErrorLevel = 2){
				MsgBox, Unexpected Error Occur!
				;break
			}
		
		}
		MsgBox, Out!
		return
	}	
	

	;그림이 찾아졌으면 클릭
	ClickTheImage(ResultOfSearching){
		if(ResultOfSearching=0){
		
		pX += 20
		pY += 10
		
		MouseClick, l, %pX%, %pY%, 2, 

		sleep 500
		return
		}
	}
	
	
	FindSecurityButtonAndClickItThenInputNumber1(){
		;StringCaseSense, On
		;jpgLocation = %A_ScripDir%PICTURES\02Security_Button.png
		jpgLocation = %A_ScripDir%PICTURES\02Security_Button1.png
		
		Loop, 3{
			;ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %jpgLocation%  ;%A_ScripDir%PICTURES\00New_Clear.png
			ImageSearch, pX, pY, 400, 100, 600, 200, %jpgLocation%  ;%A_ScripDir%PICTURES\00New_Clear.png
		
			;그림이 찾아졌으면 클릭
			if(ErrorLevel = 0){
				;MsgBox, Found it
				pX += 20
				pY += 10
		
				;MouseClick, l, %pX%, %pY%, 2
				MouseClick, l, %pX%, %pY%

				sleep 200
				
				;비밀번호 입력칸에 1 입력 후 엔터
;				SendInput, 1
;				SendInput, {Enter}
;				Sleep 1000
				
				return 0
			}
			else if(ErrorLevel = 1){
				;MsgBox, Couldn't find it!
				;break
			}
			else if(ErrorLevel = 2){
				MsgBox, Unexpected Error Occur!
				;break
			}
		
		}
;		MsgBox, Out!
	}
	
	;CC 버튼 찾고 클릭
	FindCCButtonAndClickIt(){
		jpgLocation = %A_ScripDir%PICTURES\CCButton.png
		jpgLocation2 = %A_ScripDir%PICTURES\CCButton2.png
		
		Loop, 10{
			ImageSearch, pX, pY, 0, 0, 2000, 300, %jpgLocation%
		
			;그림이 찾아졌으면 클릭
			if(ErrorLevel = 0){
				
				pX += 20
				pY += 10
		
				MouseClick, l, %pX%, %pY%, 2

				sleep 500
				
				return 0
			}
			;못 찾았으면 CCButton2 그림으로 다시 찾고 클릭
			else if(ErrorLevel = 1){
				;MsgBox, Couldn't find it!
				ImageSearch, pX, pY, 0, 0, 2000, 300, %jpgLocation2%
				pX += 20
				pY += 10
				MouseClick, l, %pX%, %pY%, 2
				
				return 0
				
				;break
			}
			else if(ErrorLevel = 2){
				MsgBox, Unexpected Error Occur!
				;break
			}
		
		}
		MsgBox, Out!
	}	
	

	
	
	
	;인보이스 내용이 담긴 Invoice_Memo 읽어서 내용 중 PO번호만 얻는 함수
	FindingPOs(Invoice_Memo, ByRef FoundPos, ByRef F_arr, ByRef L_arr, ByRef W_arr, ByRef i, ByRef j, ByRef k, ByRef lv_F, ByRef lv_W, ByRef lv_L){
		
		;MTR로 시작하는 FASHIONGO, 웹 주문 PO찾기
		;인보이스 맨 끝에 '공란TRACKING#' 먼저 입력하고 PO번호 찾아야 됨. 문장 맨 뒤까지 검색하는 방법을 모르겠음
		while(FoundPos := RegExMatch(Invoice_Memo, "imU)(Sales #\d.*/PO\s#(MTR.*)[\,|\s])", SubPat, FoundPos + strLen(SubPat)))
		{
;			MsgBox, FoundPos = %FoundPos%`nSubPat = %SubPat%`nSubPat1 = %SubPat1%`nSubPat2 = %SubPat2%
			
			;MsgBox, % strLen(SubPat)
			;SubPat2 길이가 10보다 크다는건 패션고 PO이다. 따라서 F_arr배열에 집어넣는다.
			if(strLen(SubPat2) >= 10){
				
				F_arr[i] := SubPat2
				i += 1
				lv_F := % FoundPos
;				MsgBox, lv_F`n%lv_F%
;				MsgBox, % F_arr[%a_index%]
;				MsgBox, % AA[0]

			}
			else{
					;MsgBox, It's a WEB
					W_arr[j] := SubPat2   ;SubPat2 길이가 10보다 작은건 웹 PO니까 W_arr배열에 넣는다.
					j += j
					lv_W := % FoundPos
;					MsgBox, lv_W`n%lv_W%
				}
			
				
		;MsgBox, FoundPos = %FoundPos%`nSubPat = %SubPat%`nSubPat1 = %SubPat1%`nSubPat2 = %SubPat2%
		;MsgBox, It's MTR
		;MsgBox, % strLen(SubPat)
		
		}
		; 배열의 i와 j번째까지 값이 저장됨(배열에 i나 j번째 들어있는 PO값이 패션고나 웹의 PO중 가장 최근 것)
		i -= 1
		j -= 1
		
		;MsgBox, FoundPos value`n%FoundPos%
		

		;OP로 시작하는 LASHOROOM 주문 PO찾기
		;인보이스 맨 끝에 '공란TRACKING#' 먼저 입력하고 PO번호 찾아야 됨. 문장 맨 뒤까지 검색하는 방법을 모르겠음
		FoundPos = 1
		while(FoundPos := RegExMatch(Invoice_Memo, "imU)(Sales #\d.*/PO\s#(OP.*)[\,|\s])", SubPat, FoundPos + strLen(SubPat))) 
		{
			;MsgBox, % strLen(SubPat)
				L_arr[k] := SubPat2
				k += 1
				lv_L := % FoundPos
;				MsgBox, lv_L`n%lv_L%
				
				
;		MsgBox, FoundPos = %FoundPos%`nSubPat = %SubPat%`nSubPat1 = %SubPat1%`nSubPat2 = %SubPat2%
		;MsgBox, It's OP
		;LASHOROOM 처리하는 함수 호출(SubPat2값 넘겨줘야 됨)
		}
		; 배열의 k번째까지 값이 저장됨(배열에 k번째 들어있는 PO값이 LASHOROOM PO중 가장 최근 것)
		k -= 1
				
		return
	}
	
	
	;셋 중 가장 큰 수 찾는 함수
	get_max_among_3(a, b, c){
		if(a > b){
			if(a > c)
				return a ;
			else return c ;
		}
		else{
			if(b > c)
				return b ;
			else return c ;
		}
	}



	
	
	
	; 마우스 위치 받아서 전체선택 후 복사해서 클립보드 값 리턴
	ClickAndCtrlAll(ValX, ValY){
		;먼저 클립보드 내용 비워주기
		Clipboard :=
		
		MouseClick, l, ValX, ValY, 2
		SendInput, ^a^c
		sleep 1000 ;클립보드에 내용 복사 후 꼭 1초를 쉬어줘야 제대로 클립보드에 입력됨
		
		;MsgBox, % Clipboard
		return %Clipboard%
	}
	
	
	
	; 시작위치와 종료위치를 받아서 드래그 후 복사해서 클립보드 값을 리턴해줌
	DragAndCopy(fromX, fromY, toX, toY){
		;먼저 클립보드 내용 비워주기
		Clipboard :=
		
		MouseClickDrag, l, fromX, fromY, toX, toY
		SendInput, ^c
		sleep 1000 ;클립보드에 내용 복사 후 꼭 1초를 쉬어줘야 제대로 클립보드에 입력됨
		
		MsgBox, % Clipboard
		return %Clipboard%
		
		/*
		;MsgBox, %clipboard%
		InvoiceMemoOnLAMBS := % Clipboard
		MsgBox, InvoiceMemoOnLAMBS`n`n`n%InvoiceMemoOnLAMBS%
		SendInput, ^c
		*/
	}
	
	
	; 시작위치, 종료위치, 값을 받아서 드래그 후 붙여넣기
	DragAndPast(fromX, fromY, toX, toY, InputVal){
		
;		MsgBox, %InputVal%
		
		MouseClickDrag, l, fromX, fromY, toX, toY
		SendInput, %InputVal%
		sleep 500
		
		
		return	
	}



	; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
	FindAdd2InAdd1(Address1, ByRef Address2){
		while(FoundPos := RegExMatch(Address1, "im)\s((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)", FoundAdd2)){
			if(ErrorLevel = 0){
				Address2 := FoundAdd2
				
				; Address2 에서 모든 Space(공란)를 제거합니다
				;StringReplace, Address2, Address2, %A_Space%, , All
				
				; Address2의 왼쪽에서 공란을 제거합니다.
				Address2 := LTrim(Address2, " ")
				
;				MsgBox, Found iT`n%Address2%
			}
		break
		}
	}


	; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
	DeleteAdd2InAdd1(ByRef Address1){
;		MsgBox, %Address1%
		
		while(FoundPos := RegExMatch(Address1, "im)\s((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)", FoundAdd2)){
			if(ErrorLevel = 0){
				Address1 := SubStr(Address1, 1, FoundPos-1)
;				MsgBox, Found iT`n%Address1%
			}
		break
		}		
	}
	
	
	
	; 마우스 위치, 입력값 받아서 그 위치에서 클릭 후 값 입력	
	MouseClickAndPaste(ValX, ValY, InputVal){
		MsgBox, % InputVal
		MouseClick, l, ValX, ValY, 2
		Sleep 500
		SendInput, %InputVal%
		SendInput, {Enter}
		sleep 500
		return
	}


	;비활성 창에 건내받은 키를 누른다
	Control_SnedButton(Key, windowtitle){
;		ControlSend, , {%key%}, %windowtitle%
		ControlSend, , {%key%}, %windowtitle%, l, 2
		Sleep 1000
		
		return
	}


	; TextLoc - 텍스트 입력할 위치
	; Contents - 입력 내용
	; windowtitle - 윈도우 창 제목
	Control_InputText(TextLoc, Contents, windowtitle){
		ControlSetText, %TextLoc%, %Contents%, %windowtitle%
		Sleep 500
		
		return
	}
	
	;몇번째인지 숫자 입력 받아서 LAMBS의 Invoice Memo에 내용 적고 이메일 보내기
	InputDeclineAtInvoiceMemo(num, username){
		
		FormatTime, CurrentDateTime, ,MM/dd/yyyy
		
		;LAMBS Window 활성화 하기
		windowtitle = LAMBS -  Garment Manufacturer & Wholesale Software
		CheckTheWindowPresentAndActiveIt(windowtitle)
		
		
;		WinActivate, Untitled - Notepad
		
		; No of Box 클릭 후 탭 툴러서 Invoice Memo로 들어가기
		; 이렇게 이동하면 자동으로 커서가 Invoice Memo 내용 맨 마지막에 위치하니까
		MouseClick, l, 591, 500
		SendInput, {Tab}
		SendInput, {SPACE}DECLINE {#}%num% EMAIL SENT %CurrentDateTime%{SPACE}%username%{SPACE}
		
		return
	}
	
	;상태바에서 알트키, 방향키 등 눌러서 Create Invoice로 돌아가기
	OpenCreateInvoiceTab(){
		
		Start()

		; 혹시 위의 동작으로도 Create Invoice Tab 탭으로 넘어가지 않았을 때 백업 용도로 비활성화된 Create Invoice Tab 버튼 클릭하기
		; 이 코드는 메뉴를 클릭해서 Create Invoice Tab 을 새로 여는 게 아니라 이미 열려있는데 활성화 되지 않았을 때만 유효함			
		; 혹시 마우스 커서 등이 가려서 못 찾는 것을 방지하기 위해 일단 마우스를 옮기기
;		MouseMove, A_ScreenWidth/2, A_ScreenHeight/2

		Text:="|<Create Invoice Tab>*174$73.C00000U00100800000E0000080000080000042bDCQ4gUC8ss14EUF2N98YUX0XtkDV8Y42ETUF1040YE218888UY20G8l4Y43YCSAs94EQFlo"

		if ok:=FindText(380,90,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
				
			; 탭 버튼 클릭 후 화면의 정중앙으로 마우스 포인터 이동시키기
;			MouseMove, A_ScreenWidth/2, A_ScreenHeight/2
			;~ MsgBox
			
			; 찾아서 클릭했으면 그냥 함수 빠져나가기
			;~ MsgBox, BEING FOUND BY FindText()
			return
		}
*/		

		Send, {Alt}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Down}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Down}
		Sleep 10
		
		Send, {Down}
		Sleep 10
		
		Send, {Down}
		Sleep 10
		
		Send, {Enter}
		Sleep 100	
		
		return
	}

	
	;상태바에서 알트키, 방향키 등 눌러서 Create Sales Orders Small Tab열기
	OpenCreateSalesOrdersSmallTab(){

		Start()		

		; 혹시 위의 동작으로도 Create Sales Orders Small 탭으로 넘어가지 않았을 때 백업 용도로 비활성화된 Create Sales Orders Small 버튼 클릭하기
		; 이 코드는 메뉴를 클릭해서 Create Sales Orders Small 을 새로 여는 게 아니라 이미 열려있는데 활성화 되지 않았을 때만 유효함
		; 혹시 마우스 커서 등이 가려서 못 찾는 것을 방지하기 위해 일단 마우스를 옮기기
;		MouseMove, A_ScreenWidth/2, A_ScreenHeight/2


		Text:="|<>*176$26.Q002M000a0009UhnmKAY4YG8792W22EcUYbm8D9U"
		
		if ok:=FindText(296,90,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
				
			; 탭 버튼 클릭 후 화면의 정중앙으로 마우스 포인터 이동시키기
;			MouseMove, A_ScreenWidth/2, A_ScreenHeight/2
			;~ MsgBox
				
			; 버튼 찾아서 클릭했으면 그냥 함수를 빠져나가기
			;~ MsgBox, BEING FOUND BY FindText()
			return
		}
*/		

		Send, {Alt}
		Sleep 10
	
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10

		Send, {Down}
		Sleep 10
		
		Send, {Right}
		Sleep 10

		Send, {Down}
		Sleep 10
		
		Send, {Enter}
		Sleep 100		

		return
	}


	;상태바에서 알트키, 방향키 등 눌러서 Customer Info Tab 열기
	OpenCustomerInfoTab(){
		
		Start()
	
		Send, {Alt}
		Sleep 100
		
		Send, {Right}
		Sleep 100
		
		Send, {Right}
		Sleep 100
		
		Send, {Down}
		Sleep 100
		
		Send, {Right}
		Sleep 100
		
		Send, {Enter}
		Sleep 1000
		
		return
	}

	
	;상태바에서 알트키, 방향키 등 눌러서 Style Master Tab열기
	OpenStyleMasterTab(){
		
		Start()		
	
		Send, {Alt}
		Sleep 100
		
		Send, {Right}
		Sleep 100
		
		Send, {Right}
		Sleep 100
		
		Send, {Down}
		Sleep 100
			
		Send, {Down}
		Sleep 100
			
		Send, {Down}
		Sleep 100
		
		Send, {Right}
		Sleep 100
		
		Send, {Enter}
		Sleep 1000
		
		return
	}


	;상태바에서 알트키, 방향키 등 눌러서 My Own Customer Order Report Tab 열기
	MyOwnCustomerOrderReportTab(){
		
		Start()
	
		Send, {Alt}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Down}
		Sleep 10
				
		Send, {Down}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Enter}
		Sleep 100
		
		Send, {Down}
		Sleep 100
				
		return
	}


	InputInfoAtUPS(loc, UPSVal, windowtitle){
		
		ControlClick, %loc%, %windowtitle%
		Sleep 100
		Send, ^a
		SendInput, %UPSVal%
;		Control_InputText("Edit5", Attention, windowtitle)		
		SendInput, {Enter}
		
		Sleep 1000
		
		return
	}
	
	
	
	ignoreUPSWorldShip(){
		
		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip
	
		;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	;	ControlSend, Button1, {Enter}, UPS WorldShip
	;	Control_SnedButton(Enter, UPS WorldShip)


		;UPS WorldShip 경고창이 활성화로 뜨면 그냥 무시하고 엔터치기
;		IfWinActive, UPS WorldShip
;			Send, {Enter}
		
		return
	}
	
	
	; 배열에 들어있는 값 갯수 구하는 함수
	Obj_Length(obj) { 
		length := 0 
		for idx in obj 
			length++ 
		return length 
	}
	
	
	ExtractWtsofBoxes(Invoice_Wts){

		; Invoice_Wts값에 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 Wts_of_Boxes에 저장
		StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.

		; Wts_of_Boxes를 루프 돌려서 들어있는 값 개수만큼 No_of_Boxes에 저장
		Loop, %Wts_of_Boxes0%
		{
			
		;	this_Wts := Wts_of_Boxes%a_index%
		;	MsgBox, The Weight Box number %a_index% is %this_Wts%
			No_of_Boxes := % A_Index
		}
		
		MsgBox, The number of value of ColorArray is `n`n%No_of_Boxes%
	
		return
	}
	
	
	; CConLAMBS 함수 호출 후 DeclineProcessing 호출하는 
	CallCConLAMBSAndAskWhetherCCwentOrNot(Consolidation, loc_of_MostRecentPo, Invoice_Memo, Invoice_No){
		
		CConLAMBS(loc_of_MostRecentPo)

		; 카드 결제가 제대로 됐는지 묻고 Yes 눌렀으면 UPS Label 인쇄 함수 호출
		MsgBox, 4100, UPS Label Print out, Does Credit Card Go Passed?
		IfMsgBox, Yes
		{
			UPSLabelPrintOut()
			
			WrapUp1st(Consolidation, loc_of_MostRecentPo, Invoice_Memo, Invoice_No, Delivery)
		}
		else IfMsgBox, No
		{
			;Decline 이메일 보내는 함수 호출
			;MsgBox, NO~~~
			DeclineProcessing()
		}

		return
	}
	
	
	
	

	
	; 고객에게 이메일 발송하는 함수
	SendingEmailCustomerInvoice(Invoice_No){
		
		Start()
		
		CoordMode, mouse, Relative

		;security 버튼 풀기
		FindSecurityButtonAndClickItThenInputNumber1()
		Sleep 500


		; 아웃룩 없으면 실행하도록 메세지로 알림
		IfWinNotExist, ahk_exe OUTLOOK.EXE
			MsgBox, Is Outlook Activated?


		;Email 버튼이 체크 안 되었는 지 찾아보고 체크 안 되어 있으면 체크하기
		;이메일 보낼 때 사용
		jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked.png
		CheckEmailButtonOrReleaseIt(jpgLocation)
		jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked_Activated.png
		CheckEmailButtonOrReleaseIt(jpgLocation)
		jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked_Activated2.png
		CheckEmailButtonOrReleaseIt(jpgLocation)


/*
;Email 버튼이 체크 되었는 지 찾아보고 체크 되어 있으면 체크를 풀기
;일반 인쇄 시 혹은 이메일 보낸 후 체크버튼 풀 때 사용
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked.png
CheckEmailButtonOrReleaseIt(jpgLocation)
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked_Activated.png
CheckEmailButtonOrReleaseIt(jpgLocation)
*/



		; Invoice(M) 버튼 클릭하기
		WinActivate, LAMBS
		Send, !m
		Sleep 100

		;아웃룩 창 이름 windowtitle 변수에 넣기
		windowtitle = Untitled - Message (HTML)	

		;아웃룩 창이 뜰 때까지 기다림
		WinWaitActive, %windowtitle%
		
		Sleep 1000
		
		;도착주소에 전역변수인 Email값 넣기
		ControlClick RichEdit20WPT1, %windowtitle%, , l, 3

		SendInput, %Email%
		SendInput, {Tab}


		
		;날짜를 입력하기 위한 처리
		FormatTime, CurrentDateTime, ,MM/dd/yyyy
		title = INVOICE JODIFL %Invoice_No% %CurrentDateTime%
		
				
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
	
	
		;Email 버튼이 체크 되었는 지 찾아보고 체크 되어 있으면 체크를 풀기
		;일반 인쇄 시 혹은 이메일 보낸 후 체크버튼 풀 때 사용
		jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked.png
		CheckEmailButtonOrReleaseIt(jpgLocation)
		jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked_Activated.png
		CheckEmailButtonOrReleaseIt(jpgLocation)
			
		
		return
	}
	
	
	; Invoice_Memo 읽어서 내용 중 쇼에서 온 오더나 티파니가 받은 오더 등을 찾는 함수
	FindingPOsOfShowPhoneEmail(Invoice_Memo, Invoice_No){

		;, Sales #숫자/PO{공백문자}#으로 시작하고 모든숫자나 모든문자만 \d.*|[a-zA-Z]* 에 맞고 콤마나 공백문자 [\,|\s] 앞의 문자열을 찾아라
		FoundPos := RegExMatch(Invoice_Memo, "imU)(Sales #\d.*/PO\s#(\d.*|[a-zA-Z]*)[\,|\s])", SubPat)
		
		;만약 쇼,전화,이메일 주문을 찾았으면 고객에게 이메일 발송
		if(FoundPos)
			SendingEmailCustomerInvoice(Invoice_No)

		
;		MsgBox, % SubPat2
;		;MsgBox, % FoundPos
			
		return
	}	
	
	
	; 패션고 페이지에서 정보 읽어오기
	FindInfoInFASHIONGO(HTMLSource){
		
		;값이 제대로 왔나 확인
;		MsgBox, % HTMLSource
		
;		FoundPos = 1

		; CompanyName 찾기
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipCompanyName.>(.*)</span>)", SubPat)
		
		; CompanyName 찾았으면 전역변수인 CompanyName에 값 넣기
		if(FoundPos){
			CompanyName := SubPat2
			;StringUpper, CompanyName, CompanyName ; 회사 이름 대문자로 바꾸기 회사 이름은 그냥 대소문자 섞인 게 나은 것 같음
			
			;MsgBox, 함수 안에서 찾은 것 CompanyName`n%CompanyName%
			;MsgBox, % FoundPos
		}
		
;		SubPat :=
	
		; Attention 찾기. CompanyName 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAttention.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; Attention 찾았으면 전역변수인 Attention 에 값 넣기
		if(FoundPos){
			Attention := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Attention`n%Attention%
			;MsgBox, % FoundPos
		}
		
		; Address1 찾기. Attention 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; Address1 찾았으면 전역변수인 Address1 에 값 넣기
		if(FoundPos){
			Address1 := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Address1`n%Address1%
			;MsgBox, % FoundPos
		}		
		
		; City 찾기. Address1 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress2.>(.*)[\,])", SubPat, FoundPos + strLen(SubPat))
		
		; City 찾았으면 전역변수인 City 에 값 넣기
		if(FoundPos){
			City := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 City`n%City%
			;MsgBox, % FoundPos
		}
		

		; State 찾기. 처음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress2.>.*(\w\w)\s\d.*</span>)", SubPat) ;, FoundPos + strLen(SubPat))
		
		; State 찾았으면 전역변수인 State 에 값 넣기
		if(FoundPos){
			State := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 State`n%State%
			;MsgBox, % FoundPos
		}

		
		; ZipCode 찾기. 처음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress2.>.*(\d.*)</span>)", SubPat) ;, FoundPos + strLen(SubPat))
		
		; ZipCode 찾았으면 전역변수인 ZipCode 에 값 넣기
		if(FoundPos){
			ZipCode := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 ZipCode`n%ZipCode%
			;MsgBox, % FoundPos
		}
		

		; Country 찾기. ZipCode 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipCountry.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; Country 찾았으면 전역변수인 City 에 값 넣기
		if(FoundPos){
			Country := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Country`n%Country%
			;MsgBox, % FoundPos
		}		
		
		
		; Phone 찾기. Country 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblPhoneShipping.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; Phone 찾았으면 전역변수인 Phone 에 값 넣기
		if(FoundPos){
			Phone := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Phone`n%Phone%
			;MsgBox, % FoundPos
		}
		


		; Customer Note 찾기. Phone 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(font-weight: bold.>`r(.*)</textarea>)", SubPat, FoundPos + strLen(SubPat))
		
		
		; Customer Note 찾았으면 전역변수인 CustomerNoteOnWeb 에 값 넣기
		if(FoundPos){
			CustomerNoteOnWeb := SubPat2
			StringUpper, CustomerNoteOnWeb, CustomerNoteOnWeb
			
			;MsgBox, 함수 안에서 찾은 것 `n`n%CustomerNoteOnWeb%
			;MsgBox, 함수 안에서 찾은 것 `n`n%SubPat2%
			;MsgBox, % FoundPos
		}				
		

		; Staff Only Note 찾기. Customer Note 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(font-weight: bold.>`r(.*)</textarea>)", SubPat, FoundPos + strLen(SubPat))
		
		
		; Staff Only Note 찾았으면 전역변수인 CustomerNoteOnWeb 에 값 넣기
		if(FoundPos){
			StaffOnlyNote := SubPat2
			StringUpper, StaffOnlyNote, StaffOnlyNote
			
			;MsgBox, 함수 안에서 찾은 것 `n`n%StaffOnlyNote%
			;MsgBox, 함수 안에서 찾은 것 `n`n%SubPat2%
			;MsgBox, % FoundPos
		}	

		; BillingAdd1 찾기
		FoundPos := RegExMatch(HTMLSource, "mU)(lblBillAddress1.>(.*)\s*</span>)", SubPat)
		
		; BillingAdd1 찾았으면 전역변수인 BillingAdd1 에 값 넣기
		if(FoundPos){
			BillingAdd1 := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 CompanyName`n%CompanyName%
			;MsgBox, % FoundPos
		}
		
		
		; BillingZip 찾기. BillingZip 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblBillAddress2.>.*\s(\d*)</span>)", SubPat) ;, FoundPos + strLen(SubPat))
		
		; BillingZip 찾았으면 전역변수인 BillingZip 에 값 넣기
		if(FoundPos){
			BillingZip := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 City`n%City%
			;MsgBox, % FoundPos
		}
		


;		WinActivate, LAMBS
		
		Start()
		Start_Invoice_1()		
		
		MouseMove, 857, 313
		MouseGetPos, , , , control
		ControlGetText, SubTotal, %control%, LAMBS
		

		; Invoice 2 탭 클릭하기
		MouseClick, l, 101, 266, LAMBS
		Sleep 100
		
		;MouseMove, 260, 738
		MouseMove, 310, 742
		MouseGetPos, , , , control
		ControlGetText, Email, %control%, LAMBS

		
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
		DeleteAdd2InAdd1(Address1)
		
		
		; Invoice 1 탭 클릭하기
		MouseClick,l , 41, 265		
		
		
;		MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nState : %State%`n`nPhone : %Phone%`n`nEmail : %Email%%`n`CustomerNoteOnWeb : %CustomerNoteOnWeb%
		
		
		return

	}
	
	


	
	
	
	; OP072125585는 consolidation 주문.
	; 이건 배송 주소가 아예 la쇼룸이다.
	; LA쇼룸 페이지에서 정보 읽어오기
	; 여기서는 좀 꼼수가 통했음. 자세히 보면 정규식이 Country 빼고 다 똑같음. 처음에 CompanyName 찾고 그 다음을 찾는 방식으로 정보 얻음
	FindInfoInLASHOWROOM(HTMLSource){
		
		; CompanyName 찾기
		;FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat)
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)<input type)", SubPat)
		
		; CompanyName 찾았으면 전역변수인 CompanyName에 값 넣기
		if(FoundPos){
;			CompanyName1 := SubPat1
			CompanyName := SubPat2
;			StringUpper, CompanyName, CompanyName1
			StringUpper, CompanyName, CompanyName
			
			MsgBox, 함수 안에서 찾은 것`n`nCompanyName1 : %CompanyName1%`n`n`n`nCompanyName :%CompanyName%
			;MsgBox, % FoundPos
		}


		; Address1 찾기. CompanyName 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; Address1 찾았으면 전역변수인 Address1 에 값 넣기
		if(FoundPos){
			;Address11 := SubPat1			
			Address1 := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`nAddress11 : %Address11%`n`n`n`nAddress1 :%Address1%
			;MsgBox, % FoundPos
		}


		; Address2 찾기. Address1 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; Address2 찾았으면 전역변수인 Address2 에 값 넣기
		if(FoundPos){
			;Address22 := SubPat1			
			Address2 := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`n Address22 : %Address22%`n`n`n`n Address2 :%Address2%
			;MsgBox, % FoundPos
		}
		
		
		; City 찾기. Address2 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; City 찾았으면 전역변수인 City 에 값 넣기
		if(FoundPos){
			;City1 := SubPat1			
			City := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`n City1 : %City1%`n`n`n`n City :%City%
			;MsgBox, % FoundPos
		}		
		

		; State 찾기. City 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; State 찾았으면 전역변수인 State 에 값 넣기
		if(FoundPos){
			;State1 := SubPat1			
			State := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`n State1 : %State1%`n`n`n`n State :%State%
			;MsgBox, % FoundPos
		}		
		
		
		; ZipCode 찾기. State 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; ZipCode 찾았으면 전역변수인 ZipCode 에 값 넣기
		if(FoundPos){
			;ZipCode1 := SubPat1			
			ZipCode := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`n ZipCode1 : %ZipCode1%`n`n`n`n ZipCode :%ZipCode%
			;MsgBox, % FoundPos
		}
		
		
		; Country 찾기. ZipCode 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray. id=.sh_ship_country.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; Country 찾았으면 전역변수인 Country 에 값 넣기
		if(FoundPos){
			;Country1 := SubPat1
			Country := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`n Country1 : %Country1%`n`n`n`n Country :%Country%
			;MsgBox, % FoundPos
		}


		; Attention 찾기. Country 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; Attention 찾았으면 전역변수인 Attention 에 값 넣기
		if(FoundPos){
			;Attention1 := SubPat1			
			Attention := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`n Attention1 : %Attention1%`n`n`n`n Attention :%Attention%
			;MsgBox, % FoundPos
		}


		; Phone 찾기. Attention 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(<td class=.pad_lt f f11 td_gray.>(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; Phone 찾았으면 전역변수인 Phone 에 값 넣기
		if(FoundPos){
			;Phone1 := SubPat1			
			Phone := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것`n`n Phone1 : %Phone1%`n`n`n`n Phone :%Phone%
			;MsgBox, % FoundPos
		}
		
		
		; Customer Note 찾기. Phone 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(Comment:(.*)</td>)", SubPat, FoundPos + strLen(SubPat))
		
		; Customer Note 찾았으면 전역변수인 CustomerNoteOnWeb 에 값 넣기
		if(FoundPos){
			CustomerNoteOnWeb := SubPat2
			StringUpper, CustomerNoteOnWeb, CustomerNoteOnWeb
			
			;MsgBox, 함수 안에서 찾은 것 `n`n%CustomerNoteOnWeb%
			;MsgBox, 함수 안에서 찾은 것 `n`n%SubPat2%
			;MsgBox, % FoundPos
		}
		

		; Staff Only Note 찾기. Customer Note 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(cancel_comment.>(.*)</textarea>)", SubPat, FoundPos + strLen(SubPat))
		
		
		; Staff Only Note 찾았으면 전역변수인 CustomerNoteOnWeb 에 값 넣기
		if(FoundPos){
			StaffOnlyNote := SubPat2
			StringUpper, StaffOnlyNote, StaffOnlyNote
			
			;MsgBox, 함수 안에서 찾은 것 `n`n%StaffOnlyNote%
			;MsgBox, 함수 안에서 찾은 것 `n`n%SubPat2%
			;MsgBox, % FoundPos
		}
		
		





		Start()
		Start_Invoice_1()		
		
		MouseMove, 857, 313
		MouseGetPos, , , , control
		ControlGetText, SubTotal, %control%, LAMBS
		
		
		; Invoice 2 탭 클릭하기
		MouseClick, l, 101, 266, LAMBS
		Sleep 100
		
		
		MouseMove, 260, 738
		MouseGetPos, , , , control
		ControlGetText, Email, %control%, LAMBS		
		
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
		DeleteAdd2InAdd1(Address1)
		
		
		; Invoice 1 탭 클릭하기
		MouseClick,l , 41, 265

		

		return
	}

	
	
	
	
	
	
	
	
	
	
	
	
		
	; 주문들의 마지막 기본 페이지 열기
	; i는 패션고 
	; j는 조디플
	; k는 LA쇼룸
	OpenAllBaseWebPageOfOrders(i, j, k, F_arr, W_arr, L_arr, FGServer){
		
		; 만약에 i에 값이 있으면
		; 즉, 패션고 PO가 인보이스 메모에 있으면 인보이스 메모에 있는 PO 모두 다 열기
		if(i){
			loop, %i%
			{
		;		MsgBox, A_index is `n%a_index%
		;		MsgBox, % F_arr[A_Index]
				PO_F := F_arr[A_Index]


				URL = http://vendoradmin%FGServer%.fashiongo.net/Orders.aspx#&&DateTo=&DateFrom=&SortByIndex=0&MainOrderStatusIndex=0&SearchFieldIndex=1&SearchText=%PO_F%&PeriodIndex=0
				
				; 새창으로 여니까 더 불편한 것 같음
				Run, % URL
				
				; 크롬 새창으로 열기
;				run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL
				
			}
		}
		
		
		
		return
	}
	


	; CCNumbers 받아서 4개씩 읽어서 PartiallyCCNum1,2,3,4,5에 넣기 
	PutInSpacesAmongCCNumbers(CCNumbers, ByRef PartiallyCCNum1, ByRef PartiallyCCNum2, ByRef PartiallyCCNum3, ByRef PartiallyCCNum4, ByRef PartiallyCCNum5){
		MsgBox, % CCNumbers
		StartingPos = 1

		loop, 5{
			PartiallyCCNum%A_Index% := SubStr(CCNumbers, StartingPos, 4)
			;MsgBox, % PartiallyCCNum%A_Index%
			StartingPos := StartingPos + 4
		}
		MsgBox, %PartiallyCCNum1%   %PartiallyCCNum2%   %PartiallyCCNum3%   %PartiallyCCNum4%   %PartiallyCCNum5%
		;return %PartiallyCCNum%
	}
	
	
	

;이메일 버튼을 체크하거나 체크를 푸는 함수
CheckEmailButtonOrReleaseIt(jpgLocation){
	
	WinActivate, LAMBS

	Loop, 3
	{

		WinActivate, LAMBS
		
		;jpgLocation에 있는 이미지를 찾고 찾았으면 FoundX, FoundY에 찾은 이미지 마우스 위치 받기
		PicSearch(jpgLocation, FoundX, FoundY, 900, 230, 1000, 280)
		
		
		;MsgBox, FoundX : %FoundX%`nFoundY : %FoundY%
		;Email 클릭하기 위해 위치를 조정하기
		FoundX += 10
		FoundY += 10
		
		; 이유는 모르겠는데 이미지 검색하면 Audit창이 뜰 때가 있음. Audit 창 닫기
		IfWinExist, Audit
			WinClose
		IfWinActive, Audit
			WinClose
		
		;이미지를 찾았으면 ErrorLevel 값이 0이 됨
		;0이 아니라면 즉, 이미지를 찾았다면 Email 클릭하기
		if !(ErrorLevel){
			;MouseClick, l, 992, 19
			MouseClick, l, FoundX, FoundY
			
			;Email 버튼 클릭 후 이미지 찾는 데 방해되지 않게 마우스 치우기
			MouseMove, 100, 100
			;return
		}
			
;		MsgBox, if found image = 0`n %ErrorLevel%
	}

	return
}
	
	
; 원래 LAMBS_J02 메인 함수에 있었던 consolidation 처리 내용 원본	
/*

		if(Consolidation){
			
			Invoice_Memo := % InvoiceMemoOnLAMBS

			OrdersFrom := 
			FindingPOs(Invoice_Memo, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L)	


			F_PO = % F_arr[i] ;패션고의 PO중 가장 최근의 PO값을 F_PO에 넣었다.
			W_PO = % W_arr[j] ;웹의     PO중 가장 최근의 PO값을 W_PO에 넣었다
			L_PO = % L_arr[k] ;LA쇼룸의 PO중 가장 최근의 PO값을 L_PO에 넣었다
			
			loc_of_MostRecentPo := get_max_among_3(lv_F, lv_W, lv_L)
			


			; 만약 쇼, 전화, 이메일 주문이면 아무일도 안하고
			; 패션고 주문이면 가장 최근의 PO번호 창을 열고
			; ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;조디플 홈피는 만들어야 되고;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
			; LA쇼룸 주문이면 가장 최근의 po번호 창을 열고
			if(loc_of_MostRecentPo = 1){
			}
			else if(loc_of_MostRecentPo = lv_F){
				URL = http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=%F_PO%
			
				; 크롬 새창으로 열기
				run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL			
			}
			else if(loc_of_MostRecentPo = lv_W){
				MsgBox, You should code it
			}
			else{
				; 변수 PO_L의 7번째 부터 PO_L의 길이만큼(PO_L길이 끝까지) 읽어오는 함수
				PartiallyPO := substr(L_PO,7,StrLen(L_PO))
				URL = https://admin.lashowroom.com/order_edit_v1.php?order_id=%PartiallyPO%&list_option=new
				RUN, % URL
			}
			
		
			; 각 주문별 CC함수 호출
			if(loc_of_MostRecentPo = 1)
				CConLAMBS(loc_of_MostRecentPo)
			else if(loc_of_MostRecentPo = lv_F)
				CConFashiongo()
			else if(loc_of_MostRecentPo = lv_W)
				CConJODIFL_WEB()
			else
				CConLASHOROOM()


		; 카드 결제가 제대로 됐는지 묻고 Yes 눌렀으면 UPS Label 인쇄 함수 호출
		MsgBox, 4100, UPS Label Print out, Does Credit Card Go Passed?
		IfMsgBox, Yes
		{
			WrapUp1st(1, loc_of_MostRecentPo, Invoice_Memo, Invoice_No)
			
			; 웹 주문이면 마지막 기본 페이지 열고 끝내기
			if(loc_of_MostRecentPo = 1){
				; 웹 주문 아니면 할 게 없음
			}
			else{
				; 웹 주문들의 마지막 기본 페이지 열기
				OpenAllBaseWebPageOfOrders(i, j, k, F_arr, W_arr, L_arr)
			}		
			
			Reload

		}
		else IfMsgBox, No
		{
			;Decline 이메일 보내는 함수 호출
;			MsgBox, NO~~~
			DeclineProcessing()
			Reload
			
		}	
	}
	
	
*/

;Consolidatin 처리 함수
ConsolidationProcessing(ApplyCredit, Invoice_Memo, Invoice_No, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L, Delivery, FGServer){
	
			Invoice_Memo := % InvoiceMemoOnLAMBS

			OrdersFrom := 
			FindingPOs(Invoice_Memo, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L)	


			F_PO = % F_arr[i] ;패션고의 PO중 가장 최근의 PO값을 F_PO에 넣었다.
			W_PO = % W_arr[j] ;웹의     PO중 가장 최근의 PO값을 W_PO에 넣었다
			L_PO = % L_arr[k] ;LA쇼룸의 PO중 가장 최근의 PO값을 L_PO에 넣었다
			
			loc_of_MostRecentPo := get_max_among_3(lv_F, lv_W, lv_L)
			


			; 만약 쇼, 전화, 이메일 주문이면 아무일도 안하고
			; 패션고 주문이면 가장 최근의 PO번호 창을 열고
			; ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;조디플 홈피는 만들어야 되고;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
			; LA쇼룸 주문이면 가장 최근의 po번호 창을 열고
			if(loc_of_MostRecentPo = 1){
			}
			else if(loc_of_MostRecentPo = lv_F){
				
				; 패션고 consolidation이면 customer notes를 읽기 위해 함수 호출				
				GetInfoFromFashiongo(F_PO, FGServer)
							
;				URL = http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=%F_PO%
			
				; 크롬 새창으로 열기
;				run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL


			}
			else if(loc_of_MostRecentPo = lv_W){
				GetInfoFromLAMBS()
				;MsgBox, You should code it			
			}
			else{
				; LASHOWROOM의 consolidation이면 customer notes를 읽기 위해 함수 호출
				GetInfoFromLASHOROOM(L_PO)
/*
				; 변수 PO_L의 7번째 부터 PO_L의 길이만큼(PO_L길이 끝까지) 읽어오는 함수
				PartiallyPO := substr(L_PO,7,StrLen(L_PO))
				URL = https://admin.lashowroom.com/order_edit_v1.php?order_id=%PartiallyPO%&list_option=new
				RUN, % URL
*/				
			}


			; 만약 ApplyCredit 가 체크 됐으면 ApplyCredit함수 호출
			if(ApplyCredit){
			
				Start()
				ApplyCreditFunction(Invoice_No)
				Sleep 500
				
				; Credit 적용된 이후에는 값이 바뀌었을 테니 Invoice Balance 값 다시 저장하기
				InvoiceBalance := ClickAndCtrlAll(867, 591)
				SendInput, {F8}
			}			
			
			
			; LAMBS의 인보이스메모, 커스토머메모, 웹의 커스토머메모 띄우기
			MsgBox, -Invoice Memo On LAMBS-`n`n`n%InvoiceMemoOnLAMBS%`n`n`n`n`n`n`n-Customer Memo On LAMBS-`n`n`n%CustomerMemoOnLAMBS%`n`n`n`n`n`n`n-Customer Notes on Web-`n`n`n%CustomerNoteOnWeb%`n`n`n`n`n`n`n`n%StaffOnlyNote%
			
			
		
			; 각 주문별 CC함수 호출
			if(loc_of_MostRecentPo = 1)
				CConLAMBS(loc_of_MostRecentPo)
			else if(loc_of_MostRecentPo = lv_F)
				CConFashiongo(F_PO, FGServer)
			else if(loc_of_MostRecentPo = lv_W)
				CConJODIFL_WEB()
			else
				CConLASHOROOM()


		; 카드 결제가 제대로 됐는지 묻고 Yes 눌렀으면 UPS Label 인쇄 함수 호출
		MsgBox, 4100, UPS Label Print out, Does Credit Card Go Passed?
		IfMsgBox, Yes
		{
			; WrapUp1st 함수 인자 첫번째 값이 1인 이유는 consolidation 값을 참 값인 1로 넘기기 위해
			; consolidation 값이 참 값인 1을 넘겨주면 아래 함수에서 Invoice Memo에 TRACKING# 대신
			; CONSOLIDATION 이나 DELIVERY TO 를 입력하게 됨
			WrapUp1st(1, loc_of_MostRecentPo, Invoice_Memo, Invoice_No, Delivery)
			
			; 웹 주문이면 마지막 기본 페이지 열고 끝내기
			if(loc_of_MostRecentPo = 1){
				; 웹 주문 아니면 할 게 없음
			}
			else{
				; 웹 주문들의 마지막 기본 페이지 열기
				OpenAllBaseWebPageOfOrders(i, j, k, F_arr, W_arr, L_arr, FGServer)
			}		
			
			Reload

		}
		else IfMsgBox, No
		{
		; 쇼,전화,이메일 주문이면 이미 CConLAMBS 열어본 상태이므로 그냥 DeclineProcessing 호출
		if(loc_of_MostRecentPo = 1){
			;Decline 이메일 보내는 함수 호출
			;MsgBox, NO~~~
			DeclineProcessing()
		}
		else ; CConLAMBS 함수 호출 후 DeclineProcessing 호출하는 ; CallCConLAMBSAndAskWhetherCCwentOrNot 함수 인자 첫번째 값이 1인 이유는 consolidation 값을 참 값인 1로 넘기기 위해
			CallCConLAMBSAndAskWhetherCCwentOrNot(1, loc_of_MostRecentPo, Invoice_Memo, Invoice_No)
			Reload
			
		}	

	
	return
}



; 현재 존재하는 IE창 접속 하기.
; WinTitle 과 Internet Explorer_Server Number 를 이용
; Access an IE object by WinTitle and Internet Explorer_Server Number:
WBGet(WinTitle="ahk_class IEFrame", Svr#=1) {               ;// based on ComObjQuery docs
   static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
        , IID := "{0002DF05-0000-0000-C000-000000000046}"   ;// IID_IWebBrowserApp
;//     , IID := "{332C4427-26CB-11D0-B483-00C04FD90119}"   ;// IID_IHTMLWindow2
   SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%
   if (ErrorLevel != "FAIL") {
      lResult:=ErrorLevel, VarSetCapacity(GUID,16,0)
      if DllCall("ole32\CLSIDFromString", "wstr","{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr",&GUID) >= 0 {
         DllCall("oleacc\ObjectFromLresult", "ptr",lResult, "ptr",&GUID, "ptr",0, "ptr*",pdoc)
         return ComObj(9,ComObjQuery(pdoc,IID,IID),1), ObjRelease(pdoc)
      }
   }
}

; 현재 존재하는 IE창 접속 하기.
; Window/Tab Name 를 이용
; Access an IE object by Window/Tab Name:
IEGet(name="") {
   IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame     ;// Get active window if no parameter
   Name := (Name="New Tab - Windows Internet Explorer")? "about:Tabs":RegExReplace(Name, " - (Windows|Microsoft)? ?Internet Explorer$")
   for wb in ComObjCreate("Shell.Application").Windows()
      if wb.LocationName=Name and InStr(wb.FullName, "iexplore.exe")
         return wb
}


OrganizingCCinfoOnFASHIONGO(){
	return
}




OrganizingFASHIONGOCCinfo(){
	
	;MsgBox, function starts
		
	; CCinfo.txt 내용을 RawCCinfo 변수에 저장하기
	FileRead, RawCCinfo, %A_ScriptDir%\CreatedFiles\CCinfo.txt
	
	;MsgBox, % RawCCinfo
	
	;CCNumbers := RegExMatch(RawCCinfo, "mU)(lblShipAttention.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
	FoundPos := RegExMatch(RawCCinfo, "mU)(CardNumber.\s(.*) CVV)", SubPat)

	; CCNumbers 찾았으면 전역변수인 CCNumbers 에 값 넣기
	if(FoundPos){
		CCNumbers := SubPat2
		
		; CCNumbers 에서 모든 - 제거하기.
		StringReplace, CCNumbers, CCNumbers, -, , All
	}
	
	
	; CVV 찾기. CCNumbers 다음부터 찾기 시작
	FoundPos := RegExMatch(RawCCinfo, "mU)(CVV.(.*)`r)", SubPat) ;, FoundPos + strLen(SubPat))

	; CVV 찾았으면 전역변수인 CVV 에 값 넣기
	if(FoundPos){
		CVV := SubPat2
	}
	
	
	; Month 찾기.
	FoundPos := RegExMatch(RawCCinfo, "mU)(Exp Date.\s(.*)/)", SubPat)

	; Month 찾았으면 전역변수인 Month 에 값 넣기
	if(FoundPos){
		Month := SubPat2
		
		;MsgBox, % StrLen(Month)
		
		; Month 문자 길이를 StringLength 에 넣기
		StringLength := StrLen(Month)
		
		;MsgBox, % StringLength
		
		; StringLength 길이가 1이면 숫자가 하나만 있다는 뜻이므로 왼쪽에 0 붙여주기 
		IfEqual, StringLength, 1
			Month = 0%Month%
		
		;MsgBox, % Month
	}
	
	
	; Year 찾기. Month 다음부터 찾기 시작
	FoundPos := RegExMatch(RawCCinfo, "mU)(Exp Date.*/(\d\d\d\d))", SubPat) ;, FoundPos + strLen(SubPat))

	; Year 찾았으면 전역변수인 Year 에 값 넣기
	if(FoundPos){
		Year := SubPat2
	}
	
		
	


	; CCinfo.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
	;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
	FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfo.txt, 1
	
	;C:\Users\JODIFL4\Desktop\000000000\LAMBS\CreatedFiles


	
		
	return
}

;OrganizingLAMBSCCinfo(ObtainedCCNumbers, ObtainedExpDate, ObtainedCVV){
OrganizingLAMBSCCinfo(ObtainedExpDate, __Num__){
	
	
	; LAMBS 에 있는 첫 번째 cc 정보로 카드 승인 취소 나면 CCNumbers2, CCNumbers3 등 변수를 넘겨주면서 함수를 호출할테니
	; 받은 신용카드번호, 보안번호를 전역변수인 CCNumbers 와 CVV에 
;	CCNumbers := ObtainedCCNumbers
;	CVV := ObtainedCVV
	

	
	; Month 찾기.
	FoundPos := RegExMatch(ObtainedExpDate, "mU)((\d\d)\s/)", SubPat)

	; Month 찾았으면 전역변수인 Month 에 값 넣기
	if(FoundPos){
		Month%__Num__% := SubPat2
		
		;MsgBox, % StrLen(Month)
		
		; Month 문자 길이를 StringLength 에 넣기
		;StringLength := StrLen(Month)
		
		;MsgBox, % StringLength
		
		; StringLength 길이가 1이면 숫자가 하나만 있다는 뜻이므로 왼쪽에 0 붙여주기 
		;IfEqual, StringLength, 1
			;Month = 0%Month%
		
		;MsgBox, % Month
	}
		
	
	; Year 찾기. 
	FoundPos := RegExMatch(ObtainedExpDate, "mU)(.*/\s(\d\d\d\d))", SubPat)

	; Year 찾았으면 전역변수인 Year 에 값 넣기
	if(FoundPos){
		Year%__Num__% := SubPat2
	}
	
			
	
	
	return
}

; LAMBS에 있는 인보이스 밸런스 값 읽어오기
GetInvoiceBalanceOnLAMBS()
{

	;LAMBS활성화 후 화면 초기화
	Start()
	
	;Invoice Balance 값 얻기
	MouseClick, l, 870, 594
	Send, ^a^c
	Sleep 700
	InvoiceBalance := Clipboard		

/*
	;Invoice Balance 값 얻기
	MouseMove, 870, 594
	Sleep 100		
	MouseGetPos, , , , control
	ControlGetText, InvoiceBalance, %control%, LAMBS
*/

;	MsgBox, % InvoiceBalance
	
	return
}


; Account Summary 열어서 Pending Orders 있는지 확인 후 
; 있으면 Create Invoice 탭으로 넘어가기
AccountSummayrProcessingONCreateSalesOrdersSmallTab(){
	
	wintitle = LAMBS -  Garment Manufacturer & Wholesale Software
	
	; Create Sales Orders Small 열기
	OpenCreateSalesOrdersSmallTab()
	
	;Hide All 클릭해서 메뉴 바 없애기
	ClickAtThePoint(213, 65)

	;Account Summary 열기
	MouseClick, l, 920, 155
	WinWaitActive, Accounts Summary
	WinMaximize
	
	;~ MsgBox, , , PLEASE CHECK THE ORDER HISTORY`n`nTHIS WINDOW WILL BE CLOSED IN 4 SECONDS, 4

		
		; 지정된 위치(현재 페이지)에서 Capture2Text 이용해서 값 얻기
		DllCall("SetCursorPos", int, 41-8, int, 484-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).		
		SendInput, #q
		Sleep 200
		DllCall("SetCursorPos", int, 89-8, int, 497-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 200
		SendInput, #q
		Sleep 200
					
		; Capture2Text 창 닫기		
		IfWinExist, Capture2Text - OCR Text
			WinClose


		CheckPendingIDExist := % Clipboard
		Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
		
		; 숫자만 저장
		CheckPendingIDExist := RegExReplace(CheckPendingIDExist, "[^0-9]", "")


		; CheckPendingIDExist 에서 공백문자 제거하기
		StringReplace , CheckPendingIDExist, CheckPendingIDExist, %A_Space%,,All


		; 만약 Pending 된 주문이 있으면 Create Invoice 이동해서 펜딩된 오더 열기
		if(CheckPendingIDExist){
			
			; Yes 버튼을 눌러서 Pending Order 처리를 해줬다는 것을 표시하기 위해
			; PendingOrderStatus 값을 1로 바꿈
			PendingOrderStatus = 1
;			MsgBox, PendingOrderStatus IS %PendingOrderStatus%
			
			WinClose, Accounts Summary
			
			OpenCreateInvoiceTab()
			Sleep 200
					
			;New & Clear 버튼 클릭
			MouseClick, l, 60, 124, 2, 
			sleep 1000	

			;Order ID 라디오 버튼 선택하기
			MouseClick, l, 37, 205, 2, 
			sleep 500
			
			;인보이스 번호 입력하기
			MouseClick, l, 209, 205, 2, 
			SendInput, %CheckPendingIDExist%
			SendInput, {Enter}
			sleep 2000
			
			; 인보이스 번호 입력 후 데이터 없다는(?) 에러가 나타나면 다시 입력하기
			IfWinExist, Confirm
			{
				WinClose, Confirm
					
				;New & Clear 버튼 클릭
				MouseClick, l, 60, 124, 2, 
				sleep 1000
				
				;인보이스 번호 입력하기
				MouseClick, l, 209, 205, 2, 
				SendInput, %CheckPendingIDExist%
				SendInput, {Enter}
				sleep 2000
			}
			
			
			; Accounts Summary Button 클릭하기
			ImageSearch, FoundX, FoundY, 1126, 131, 1278, 182, %A_ScripDir%PICTURES\AccountsSummaryButton.png	
			FoundX += 5
			FoundY += 5
			MouseClick, l, %FoundX%, %FoundY%
			
			WinWaitActive, Accounts Summary
			
			WinMove, Accounts Summary, , 752, 55
			
					
			; 인쇄할 지 말지 판단하기 위해 일단 BO 목록표 열기 (Customer Order + 버튼 클릭)
			ControlClick, WindowsForms10.BUTTON.app.0.378734a4, Accounts Summary
			WinWaitActive, Customer Order +Zoom In			
			
			WinMove, Customer Order +Zoom In, , 752, 224, 1150
								

/*
			; CompanyName 변수값 넣기
			DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
			Sleep 100
			MouseGetPos, , , , control, 1
			ControlSetText, %control%, %CompanyName%, %wintitle%
			ControlClick %control%, %wintitle%
			SendInput, {Enter}
*/

/*
			; QOH 클릭해서 재고 현황 잘 볼 수 있게
			;ImageSearch, FoundX, FoundY, 722, 68, 974, 206, %A_ScripDir%QOH.png
			;~ ImageSearch, FoundX, FoundY, 555, 62, 773, 244, %A_ScripDir%QOH.png
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%QOH.png
			FoundX += 5
			FoundY += 5
			MouseClick, l, %FoundX%, %FoundY%
			MsgBox, ErrorLevel is %ErrorLevel%`n`n Should be 0
*/			


			Sleep 100
			
			; QOH 클릭해서 재고 현황 잘 볼 수 있게
			Text:="|<QOH>*161$24.T3sVlaAVVYAVUY4VUY4zUY4VVYAVlaAVT3sV20003U00U"

			if ok:=FindText(639,99,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Click
				Click
			}



			; 이 리턴은 살려야 됨
			return			
			
		}
			
	
	
	
	
; 자동으로 Pending 찾기 전 사람에게 물어보는 옛날 방법
/*	
		; Pneding 아이템이 있는지 묻고 Yes 눌렀으면 OpenCreateInvoiceTab 함수 호출해서 Create Invoice Tab 로 이동 후 CompanyName 넣기
		MsgBox, 260, Pending Items Check, Does Pending Items Exist?
		IfMsgBox, Yes
		{
			; Yes 버튼을 눌러서 Pending Order 처리를 해줬다는 것을 표시하기 위해
			; PendingOrderStatus 값을 1로 바꿈
			PendingOrderStatus = 1
;			MsgBox, PendingOrderStatus IS %PendingOrderStatus%
			
			WinClose, Accounts Summary
			
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


			return
		}
*/


	return
}


; Account Summary 열어서 BO 목록 인쇄하기
PrintingBOList(){
	
	wintitle = LAMBS -  Garment Manufacturer & Wholesale Software
	
	; Create Sales Orders Small 
	OpenCreateSalesOrdersSmallTab()
	
	;Hide All 클릭해서 메뉴 바 없애기
	ClickAtThePoint(213, 65)

	;Account Summary 의 BO 목록 열기
	;Account Summary 열기
	DllCall("SetCursorPos", int, 920-8, int, 155-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
	Sleep 100
	MouseGetPos, , , , control, 1
	ControlClick, %control%, %wintitle%
;	MouseClick, l, 920, 155
	Sleep 100
	WinActivate, Accounts Summary
	WinWaitActive, Accounts Summary
	
	
	; BO 목록표 열기 (Customer Order + 버튼 클릭)
	ControlClick, WindowsForms10.BUTTON.app.0.378734a4, Accounts Summary
	WinWaitActive, Customer Order +Zoom In
	WinMaximize


	; LOC 1 클릭하기
	Text:="|<LOC1 Button>*132$31.UC1s2E8V4788900Y44U0G22E0911804UUY02E8V41DXUS3w"

	if ok:=FindText(836,98,150000,150000,0,0,Text)
	{
		CoordMode, Mouse
		X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		MouseMove, X+W//2, Y+H//2
		Click
		
		ClickLOC = Ok
		;~ MsgBox, BEING FOUND BY FindText()
		
		
	}

/*
	; LOC 1 클릮하기
	ImageSearch, FoundX, FoundY, 722, 68, 974, 206, %A_ScripDir%PICTURES\LOC1.png
	FoundX += 5
	FoundY += 5
	MouseClick, l, %FoundX%, %FoundY%
	
	; 만약 LOC1 버튼 그림을 못 찾았으면 Color 버튼을 먼저 찾아서 나중에 클릭하기
	if(ErrorLevel = 1){
		; Color 버튼의 115만큼 오른쪽으로 옆을 클릭하기
		ImageSearch, FoundX, FoundY, 416, 60, 626, 176, %A_ScripDir%PICTURES\ColorButtonONCustomerOrderZoonIn.png	
		FoundX += 260
		FoundY += 5
		MouseClick, l, %FoundX%, %FoundY%
	}
*/	


	; Print(F12) 버튼 더블 클릭하기
	MouseClick, l, 47, 45, 2

	; 페이지 뜰때까지 기다리기
	WinWaitActive, Page Setup


	; Print preview 화면이 나올때까지 OK 버튼 클릭해서 인쇄하기
	Loop{
		;~ ControlClick Button6, Page Setup

		Text:="|<>*150$15.C4O8aUZY4cUbY4YUYm8WC4Q"
		if ok:=FindText(1013,602,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		}
		
		IfWinExist, Print preview
			break
	}

	
	; 클릭 포인트 얻기 쉽게 창 최대화 하기
	WinMaximize, Print preview
	
	; 프린터 모양 아이콘 클릭해서 인쇄하기
	MouseClick, l, 16, 40, 2
	
	
	WinClose, Print preview
	WinWaitActive, Page Setup
	WinClose, Page Setup


	;~ ##########################################################################################################################################
	;~ 빠른 진행을 위해 이 기능 잠깐 꺼두자
	;~ ##########################################################################################################################################
	; Order Date 위치 찾은 후 Capture2Text 이용해서 주문 날짜 값 얻는 함수 호출
	FindingOrderMonthAndCompareWithNow()


	WinClose, Customer Order +Zoom In
	WinClose, Accounts Summary	
	
	
	return
}


; Order Date 위치 찾은 후 Capture2Text 이용해서 주문 날짜 값 얻는 함수 호출
FindingOrderMonthAndCompareWithNow(){
	

		; Order Date 위치 찾은 후 Capture2Text 이용해서 주문 날짜 값 얻기
		Text:="|<Order Date>*125$59.0010000000Q02003k0E1404004E0U44tsss8Hnnc9aG9UEUG8kG8YG0V0YFUYFDY12D8z18WE824WF14F4kE4F4X1kVswUD3tns"

		if ok:=FindText(249,98,150000,150000,0,0,Text)
		{
			Current_X :=
			Current_Y :=
			OrderDate :=
			OrderMonth :=
			
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
;			MouseMove, X+W//2, Y+H//2
;			; 주문 날짜를 오래된 것부터 정열하기 위해
;			Click
			MouseMove, X-10, Y+50
			
			;~ MouseGetPos, Current_X, Current_Y
			
			; 지정된 위치(현재 페이지)에서 Capture2Text 이용해서 값 얻기
			SendInput, #q
			Sleep 200

			MouseMove, X+75, Y+940
			Sleep 200
			SendInput, #q
			Sleep 200
						
			; Capture2Text 창 닫기		
			IfWinExist, Capture2Text - OCR Text
				WinClose

			; 현재 보다 3달 전이면 지워야 되는 백오더이므로
			ThisMonth := A_MM - 3


			; Clipboard 에는 현재 Capture2Text 이용해서 얻은 날짜들 데이터가 들어가 있음
			; Clipboard 에서 새 줄 문자 나올때마다 나누고 양쪽 끝의(중간 아님) 스페이스와 탭키를 제거하기
			Loop, parse, Clipboard, `n, %A_Space%%A_Tab%
			{
				; 루프를 돌아서 현재 값이 저장된 A_LoopField 변수에서 왼쪽부터 2개만 읽어서 OrderMonth 변수에 저장
				StringLeft, OrderMonth, A_LoopField, 2


				; 값을 제대로 읽었는지 확인하기 위해 메세지 띄우기. No 버튼 누르면 루프 빠져나오기
;				MsgBox, 4, , OrderMonth is %OrderMonth%`n`nThis month is %ThisMonth%`n`n`n`nContinue?
;				IfMsgBox, No, break
				
				; 오더한 달이 현재 달보다 3달 전이거나 그보다 더 이전이면 메세지 출력하기
				IfLessOrEqual, OrderMonth, %ThisMonth%		; IfLessOrEqual, var, value (same: if var <= value)
				{
					MsgBox, 4100, OLD BO WARNNING, There is at least 1 old order in the list.`n`nWOULD YOU LIKE TO ACTIVE 'Create Sales Orders Small' TAB?`n`nIF YOU CLICK No, MOVE TO Create Invoice Tab to MAKE INVOICE
					
					; Yes 눌렀으면 
					IfMsgBox, Yes
					{
						WinClose, Customer Order +Zoom In
						WinClose, Accounts Summary
						OpenCreateSalesOrdersSmallTab()
						MsgBox, 4100, , CLICK OK IF READY TO MAKE A INVOICE
					}
					break
				}				
					
			}
			
			;~ MsgBox, % OrderDate
		}			
	
	return
}


State_Abbreviations(){
	
	IfEqual, State, AL
		State = Alabama
	
	IfEqual, State, AK
		State = ALASKA
	
	IfEqual, State, AZ
		State = ARIZONA_(WESTERN)
	
	IfEqual, State, AR
		State = ARKANSAS
	
	IfEqual, State, CA
		State = CALIFORNIA_(WESTERN)
	
	IfEqual, State, CO
		State = COLORADO_(WESTERN)
	
	IfEqual, State, CT
		State = CONNECTICUT
	
	IfEqual, State, DE
		State = DELAWARE
	
	IfEqual, State, FL
		State = FLORIDA
	
	IfEqual, State, GA
		State = GEORGIA
	
	IfEqual, State, HI
		State = HAWAII
	
	IfEqual, State, ID
		State = IDAHO_(NORTHWESTERN)
	
	IfEqual, State, IL
		State = ILLINOIS
	
	IfEqual, State, IN
		State = Indiana	
	
	IfEqual, State, IA
		State = Iowa	
	
	IfEqual, State, KS
		State = Kansas	
	
	IfEqual, State, KY
		State = Kentucky	
	
	IfEqual, State, LA
		State = Louisiana	
	
	IfEqual, State, ME
		State = Maine
	
	IfEqual, State, MD
		State = Maryland
	
	IfEqual, State, MA
		State = Massachusetts
	
	IfEqual, State, MI
		State = Michigan
	
	IfEqual, State, MN
		State = Minnesota
	
	IfEqual, State, MS
		State = Mississippi	
	
	IfEqual, State, MO
		State = Missouri	
	
	IfEqual, State, MT
		State = Montana	
	
	IfEqual, State, NE
		State = Nebraska	
	
	IfEqual, State, NV
		State = Nevada_(WESTERN)
	
	IfEqual, State, NH
		State = New Hampshire
	
	IfEqual, State, NJ
		State = New Jersey	
	
	IfEqual, State, NM
		State = New Mexico_(WESTERN)
	
	IfEqual, State, NY
		State = New York	
	
	IfEqual, State, NC
		State = North Carolina	
	
	IfEqual, State, ND
		State = North Dakota	
	
	IfEqual, State, OH
		State = Ohio	
	
	IfEqual, State, OK
		State = Oklahoma	
	
	IfEqual, State, OR
		State = Oregon_(WESTERN)
	
	IfEqual, State, PA
		State = Pennsylvania	
	
	IfEqual, State, RI
		State = Rhode Island	
	
	IfEqual, State, SC
		State = South Carolina	
	
	IfEqual, State, SD
		State = South Dakota
	
	IfEqual, State, TN
		State = Tennessee	
	
	IfEqual, State, TX
		State = Texas	
	
	IfEqual, State, UT
		State = Utah_(WESTERN)
	
	IfEqual, State, VT
		State = Vermont	
	
	IfEqual, State, VA
		State = West Virginia	
	
	IfEqual, State, WA
		State = Washington_(WESTERN)
	
	IfEqual, State, WV
		State = West Virginia	
	
	IfEqual, State, WI
		State = Wisconsin	
	
	IfEqual, State, WY
		State = Wyoming	
	
	IfEqual, State, AS
		State = American Samoa	
	
	IfEqual, State, DC
		State = District of Columbia	
	
	IfEqual, State, FM
		State = Federated States of Micronesia	
	
	IfEqual, State, GU
		State = Guam	
	
	IfEqual, State, MH
		State = Marshall Islands	
	
	IfEqual, State, MP
		State = Northern Mariana Islands	
	
	IfEqual, State, PW
		State = Palau	
	
	IfEqual, State, PR
		State = Puerto Rico	
	
	IfEqual, State, VI
		State = Virgin Islands	
	
	IfEqual, State, AE
		State = Armed Forces Africa	
	
	IfEqual, State, AA
		State = Armed Forces Americas	
	
	IfEqual, State, AE
		State = Armed Forces Canada	
	
	IfEqual, State, AP
		State = Armed Forces Pacific	


	return
}

ReadingINFOofCustomerOrderWindow(){
	
	Clipboard :=
	
	Loop, 50
		Send, {down}
	

	
	; 지정된 위치(현재 페이지)에서 Capture2Text 이용해서 값 얻기
	;~ DllCall("SetCursorPos", int, 41-8, int, 155-8)  ; 전체화면	
	;~ DllCall("SetCursorPos", int, 456-8, int, 155-8)  ; 아이템번호, 갯수만
	DllCall("SetCursorPos", int, 628-8, int, 150-8)  ; 갯수만
	SendInput, #q
	Sleep 200
	;~ DllCall("SetCursorPos", int, 1228-8, int, 1041-8)  ; 전체화면
	;~ DllCall("SetCursorPos", int, 706-8, int, 1043-8)  ; 아이템번호, 갯수만
	DllCall("SetCursorPos", int, 710-8, int, 700-8)  ; 갯수만
	Sleep 200
	SendInput, #q
	Sleep 200
/*
	; Capture2Text 창 닫기		
	IfWinExist, Capture2Text - OCR Text
		WinClose
*/

	ItemInfo := % Clipboard
	Sleep 1000 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
	
	;~ MsgBox, % ItemInfo
	
/*	
	;~ array := ["one", "two", "three"]

; 1부터 배열의 끝까지 반복합니다:
Loop % ItemInfo.Length()
    MsgBox % ItemInfo[A_Index]

; 배열의 내용을 열거합니다:
For index, value in ItemInfo
    MsgBox % "Item " index " is '" value "'"
*/	
	
	return
}