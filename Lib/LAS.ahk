#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



class LA extends CommWeb{


	; ########################
	; 공통된 Order 처리 위한 메소드
	; 원하는 검색 url 에서 PONumber 검색하기
	; ########################	
	LA_ProcessingCommonStepOfOrderProcessing(URL, PONumber){

		; 새로운 LA SHOWROOM 창을 열기
		LA.OpenNewBrowser(URL)			

		driver := ChromeGet()
		;~ MsgBox, % driver.Window.Title "`n" driver.Url
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		; 현재 페이지가 LA SHOWROOM 로그인 페이지면 Login 메소드 실행해서 로그인 후 New Orders 메뉴로 들어가기
		if(RegExMatch(driver.Url, "imU)login")){
			;~ MsgBox, It's LA SHOWROOM login page
			CommWeb.Login(URL)
			
			; 로그인 뒤 다시 해당 페이지를 열기(뉴오더면 뉴오더 페이지, 전체 검색이면 전체 오더 페이지)
			driver := ChromeGet()
			driver.Get(URL)
		}


		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 3000 ; 이상하게 로그인 후 잠시 쉬었다가 검색할 수 있는 페이지로 자동으로 바꾸기 때문에
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다


		; 검색란에 PO Number 입력하기
		Xpath = //*[@id="search_po"]
		;~ driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(PONumber).sendKeys(driver.Keys.ENTER)
		driver.FindElementByXPath(Xpath).SendKeys(PONumber).sendKeys(driver.Keys.ENTER)
		
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다


		return		
	}	
	
	
	; ########################
	; 페이지에 아이템 업데이트 하기
	; ########################
	;~ LA_UpdateItems(Array_StyleNo, Array_StylyColor, ShippingFee, InvoiceNo, TrackingNo, PONumber){
	LA_UpdateItems(PONumber){
				
		
		URL = https://admin.lashowroom.com/orders_cur_month.php ; 오더 검색창 주소		
		; LA SHOWROOM 에서 메뉴바의 ORDERS 클릭할때까지의 과정을 담은 아래 메소드 시작하기
		LA.LA_ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		
		; 다시 현재 창 컨트롤 할 수 있게 얻기
		driver := ChromeGet()
		
		

		; PO Number 입력
		Xpath = //*[@id="search_po"]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(PONumber).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 PONumber 입력 후 엔터쳐서 검색하기
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		
		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*		
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")
		
		MsgBox, % HTML_Source

		

		; OP 로 시작되는 패션고 PO Number 뽑아서 배열에 저장
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "popup'>(OP.*)</span>", SubPat, FoundPos + strLen(SubPat)))
		{
			MsgBox, % "SubPat1 : " . SubPat1
			Array_PONumber.Insert(SubPat1)
		}



		/*
		
		; MTR 로 시작되는 패션고 PO Number 뽑아서 배열에 저장
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "imU)>(MTR.*)</a>", SubPat, FoundPos + strLen(SubPat)))
			Array_PONumber.Insert(SubPat1)
		
		
		*/
		
		
		/*
		<span title="OP112436997-A" class="popup">OP112436997-A</span>
		*/
		
		
		
	}
	
	
	; ########################
	; 현재 페이지의 정보들 얻기
	; ########################		
	GetInfoFromLASPage(PONumber){
		
		
		Arr_Memo := object()
		Arr_CC := object()
		Arr_BillingADD := object()
		Arr_ShippingADD := object()
		
		
		URL = https://admin.lashowroom.com/orders_cur_month.php ; 오더 검색창 주소		
		; LA SHOWROOM 에서 메뉴바의 ORDERS 클릭할때까지의 과정을 담은 아래 메소드 시작하기
		LA.LA_ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		driver := ChromeGet()
		
	
		; PO Number 링크 나타날때까지 기다림
		;~ Xpath = //*[text() = '%PONumber%']
		Xpath = //*[contains(text(), '%PONumber%')]
		Sleep 500
		while(!driver.FindElementByXPath(Xpath).Attribute("innerText"))
			Sleep 100
		
		Sleep 1000
		
		

		; 검색된 PO 결과 창에서 가장 처음 값을 클릭하기 (토씨 하나 안 틀린 PO 값을 클릭하는게 아니라 PO 번호가 포함된 가장 처음의 링크 클릭)
		; 이게 희안한게 브라우저 글자 크기를 확대하면 작동을 안 한다
		if(driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]")){
			driver.FindElementByXPath("(//*[contains(text(), '" PONumber "')])[1]").click()
		}
		
		
		driver := ChromeGet() ; PO 클릮하면 새 창에서 열리니까
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		
		
		; Buyer Notes
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[3]/td[1]
		Arr_Memo[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_Memo[1]
		StringUpper, Data, Data ; 대문자로 바꾸기
		Arr_Memo[1] := Data
		
		
		
		
		; Contact Name on Billing Add
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td[2]
		Arr_CC[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")				
		
		

		; Billing Add
		; ADD1
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[2]
		Arr_BillingADD[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_BillingADD[1]
		StringUpper, Data, Data ; 대문자로 바꾸기
		Arr_BillingADD[1] := Data
		
		; ADD2
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]
		Arr_BillingADD[2] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_BillingADD[2]
		StringUpper, Data, Data ; 대문자로 바꾸기
		Arr_BillingADD[2] := Data
		
		; CITY
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[2]
		Arr_BillingADD[3] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_BillingADD[3]
		StringUpper, Data, Data ; 대문자로 바꾸기
		Arr_BillingADD[3] := Data
				
		; STATE
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[2]
		Arr_BillingADD[4] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_BillingADD[4]
		StringUpper, Data, Data ; 대문자로 바꾸기
		Arr_BillingADD[4] := Data
		
		; ZIP
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[6]/td[2]
		Arr_BillingADD[5] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_BillingADD[5]
		StringUpper, Data, Data ; 대문자로 바꾸기
		Arr_BillingADD[5] := Data
				
		; COUNTRY
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[2]
		Arr_BillingADD[6] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_BillingADD[6]
		StringUpper, Data, Data ; 대문자로 바꾸기
		Arr_BillingADD[6] := Data
				
		; PHONE
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td[2]
		Arr_BillingADD[7] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_BillingADD[7]
		Data := RegExReplace(Data, "[^0-9]", "") ;숫자만 저장
		Arr_BillingADD[7] := Data
		
		
		



		; Shipping Add
		; ADD1
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[2]/td[2]
		Arr_ShippingADD[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_ShippingADD[1]
		StringUpper, Data, Data ; ???? ???
		Arr_ShippingADD[1] := Data
		
		; ADD2
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[3]/td[2]
		Arr_ShippingADD[2] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_ShippingADD[2]
		StringUpper, Data, Data ; ???? ???
		Arr_ShippingADD[2] := Data
		
		; CITY
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[4]/td[2]
		Arr_ShippingADD[3] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_ShippingADD[3]
		StringUpper, Data, Data ; ???? ???
		Arr_ShippingADD[3] := Data
				
		; STATE
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[5]/td[2]
		Arr_ShippingADD[4] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_ShippingADD[4]
		StringUpper, Data, Data ; ???? ???
		Arr_ShippingADD[4] := Data
		
		; ZIP
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]
		Arr_ShippingADD[5] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_ShippingADD[5]
		StringUpper, Data, Data ; ???? ???
		Arr_ShippingADD[5] := Data
				
		; COUNTRY
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[7]/td[2]
		Arr_ShippingADD[6] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
		Data := Arr_ShippingADD[6]
		StringUpper, Data, Data ; ???? ???
		Arr_ShippingADD[6] := Data
				
		; PHONE
		Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[10]/td[2]
		Arr_ShippingADD[7] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_ShippingADD[7]
		Data := RegExReplace(Data, "[^0-9]", "") ;??? ??
		Arr_ShippingADD[7] := Data

		



;		MsgBox, % Arr_Memo[1] . "`n`n" . Arr_CC[1]
				
		/* 배열로부터 읽기 첫 번째 방법
		Loop % Arr_ShippingADD.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Arr_ShippingADD[A_Index]
		}
		*/
		
		
		
		; Update 버튼 클릭하기
		
		;~ Update_Xpath = //*[@id='update_order']
		Update_Xpath = //*[@id="update_order"]
		driver.FindElementByXPath(Update_Xpath).click()
		
		
		
		return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo]
	}











	; ########################
	; 현재 화면에 있는 모든 Customer PO 얻기
	; 여기서 얻어진 Customer PO 들은 아래 UpdateItems 메소드에서 업데이트 할 때 활용됨
	; ########################
	OpenOrderPageForItemUpdate(PONumber, Array_PONumber){
		

		; 검색되어 나온 고객 PO 모두 저정하기 위한 배열 선언
		Array_PONumber := object()

		; 프로그램 처음 시작할 때 LAS 의 PO 검색 페이지를 열고 시작했기 때문에 그 창 다시 사용하기 위해
		driver := ChromeGet()
		

		; 검색란에 PO Number 입력하기
		Xpath = //*[@id="search_po"]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(PONumber).sendKeys(driver.Keys.ENTER)
		
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 1000

		
		
		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*		
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")
		Sleep 700


		
		; OP 로 시작되는 패션고 PO Number 뽑아서 배열에 저장
		UnquotedOutputVar = imU)class="popup">(OP.*)</span></a>		
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, UnquotedOutputVar, SubPat, FoundPos + strLen(SubPat)))
			Array_PONumber.Insert(SubPat1)



/*      ; Array_PONumber 에 값이 제대로 들어갔는지 확인하기 위해
		Loop % Array_PONumber.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Array_PONumber[A_Index]
		}
		
		MsgBox PAUSE
*/


/*
		if(driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]"))
		{
			MsgBox, % driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("innerText") ; 이것만 됨
			MsgBox, % driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("innerText").Attribute("innerText")
			MsgBox, % driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("Xpath")
		}
*/


		
		; 브라우저 닫기
;		driver.close() ; closing just one tab of the browser
		
		return, Array_PONumber
		
		
		
		
	} ; OpenOrderPageForItemUpdate 메소드 끝














	
	; ########################
	; 페이지에 아이템 업데이트 하기
	; ########################
	;~ UpdateItems(Array_StyleNo, Array_StylyColor, ShippingFee, InvoiceNo, TrackingNo, PONumber){
	UpdateItems(Array_StyleNo, Array_StylyColor, Array_StylyQty, ShippingFee, InvoiceNo, TrackingNo, PONumber){



		; 체크된 아이템이 있는지 여부 확인하는 변수 초기화하기
		CheckedItemExistsOrNot :=


/*


		URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
		
		; 패션고에서 PO Number 얻기 위해 공통으로 처리하는 과정들을 담은 메소드 실행해서 PO Number 검색결과 얻기
		FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		; 다시 현재 창 컨트롤 할 수 있게 얻기
		driver := ChromeGet()

		
		; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾아서 클릭하기
		; 예를들어 MTR1D39747D26 로 검색했으면 검색된 화면에는 MTR1D39747D26 뿐만 아니라 MTR1D39747D26-BO1 등도 같이 표시되기 때문에 딱 원하는 PO Number만 클릭하기 위해
		if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
			driver.FindElementByXPath("//*[text() = '" PONumber "']").click()

*/


		; OpenOrderPageForItemUpdate 메소드에서 PO를 검색한 뒤 창을 닫지 않았기 때문에 현재 고객의 PO를 검색한 창이 열려있다.
		; 새로 창을 열지 않고 그 창을 다시 사용하기 위해
		driver := ChromeGet()


;		MsgBox, % "PONumber`n`n" . PONumber


		; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾은 뒤 해당 PO의 링크 클릭하기. 그러면 새창으로 열림
		; 예를들어 OP113037497 로 검색했으면 검색된 화면에는 OP113037497 뿐만 아니라 OP113037497-A 나 OP113037497-B 등도 같이 표시되기 때문에 딱 원하는 PO Number에 해당되는 탭버튼만 클릭하기 위해
		if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
			driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
				
;MsgBox, OPEN -A

		; PO 하이퍼 링크를 클릭하면 새창이 뜨니 새로 생긴 창을 컨트롤 할 수 있도록 함수호출
		driver := ChromeGet()
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 2000

;		MsgBox, ready to update items
		
		
		; 열린 창의 아이템 체크박스 숫자가 몇개인지 확인 한 뒤 그 갯수만큼 루프 돌려서 체크박스에 체크하기
		; 일단 아이템의 모든 체크박스를 클릭 한 뒤 업데이트 할 아이템만 체크를 풀고 백오더 버튼을 누르면
		; 업데이트 하지 않을 아이템만 백오더로 넘어간다
		MaxItem_# = 0

		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*		
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")
		Sleep 700



		; MTR 로 시작되는 패션고 PO Number 뽑아서 배열에 저장
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "imU)>(MTR.*)</a>", SubPat, FoundPos + strLen(SubPat)))
			Array_PONumber.Insert(SubPat1)



		;~ <input type="checkbox" name="mark_order_item[]" value="119491|0|6567" class="mark_order_items">

		; 소스 코드에서 체크 박스 html 뽑아서 체크박스 html 나올 때마다 MaxItem_# 의 값을 1씩 올려준다
		UnquotedOutputVar = imU)<input type="checkbox" name="mark_order_item[]".*class="mark_order_items">
		UnquotedOutputVar = class="mark_order_items">
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, UnquotedOutputVar, SubPat, FoundPos + strLen(SubPat)))
		{
			MaxItem_#++
;			MsgBox, found it
		}
			

;		MsgBox, % "MaxItem_#`n`n" . MaxItem_#


		; MaxItem_# 에 있는 값이 현재 열린 창의 최대 체크박스 갯수니까 그 갯수만큼만 루프 돌려서 모든 체크 박스에 체크하기
		Loop % MaxItem_#{
				
			; 아이템 옆에 있는 체크박스 체크해주기
			Xpath = (//input[@type='checkbox' and @name='mark_order_item[]'])[%A_Index%]
			driver.FindElementByXPath(Xpath).click()
;			MsgBox, checked
		}
		
		
		
		Sleep 500


;		MsgBox, all itmes check boxes checked?



		; 아이템 업데이트 하기
		; 우선 아이템 번호와 색깔에 맞는 Amount 값을 확인 한 뒤 Amount 값이 0이 아니면
		; 해당하는 아이템을 체크한 뒤 (실제로는 위에서 모든 아이템을 체크 해줬기 때문에 여기서 체크하면 체크가 취소되는 것)
		; 배열에서 해당 아이템의 값을 지우고 다시 루프의 첫번째로 돌아가기
		Loop % Array_StyleNo.Maxindex(){
			Loop % Array_StylyColor.Maxindex(){
				

				StyleNo := Array_StyleNo[A_Index]
				StyleColor := Array_StylyColor[A_Index]
				StyleQty := Array_StylyQty[A_Index]
				
				MsgBox % "Style No. " . A_Index . " is " . Array_StyleNo[A_Index]  "`n"  "Style Color" . A_Index . " is " . Array_StylyColor[A_Index] . "`n" . "Style Qty" . A_Index . " is " . Array_StylyQty[A_Index] ; 업데이트 할 스타일 번호와 색깔 확인하기 위해
				

				;~ Xpath = //*[text() = '%StyleNo%']


				; 빈 배열이 아니라면 즉, 배열에 아이템이 한개롣 들어있다면 시작
				;~ if(StyleNo != ""){
				if(Array_StyleNo[A_Index] != ""){
					
;					MsgBox, There is item value in Array_StyleNo
					
					; 원래 대로라면 현재 페이제에 찾는 StyleNo 가 있는지 확인 후 있다면 if 문을 실행시켜야 되는데
					; LAS 는 StyleNo 가 따로 분리되어 있지 않고 사진과 함께 한 구역에 있음
					; 그래서 불명확하지만 색깔로 일단 찾고 Amount 값을 찾는 곳에서 조건에 맞는 동작을 하도록 했음
					; 만약에 현재 페이지에 찾는 StyleColor 가 있다면
					;~ if(driver.FindElementByXPath(Xpath)){
					if(driver.FindElementByXPath("//*[text() = '" StyleColor "']")){
						
						Amount := ""
						
;						MsgBox, found matched item
						
						; StyleNo , StyleColor , StyleQty 셋 다 맞는 아이템의 Amount 값의 위치
						; StyleNo , StyleColor 두 개만 맞는 아이템의 Amount 값의 위치 해서 Amount에 넣으면 값이 0.00이든 아니든 있고 셋 다 맞는 것 찾으면 셋 다 맞지 않으면 값이 없다. 셋 다 맞는 것 해도 되고 두 개만 맞는 것 해도 될 것 같은데
						; 왜냐면 백오더로 넘어간 아이템은 Amount 값이 0이기 때문에 Amount 값이 0인 아이템은 아래 if 문에서 체크하지 않게 한다
						Xpath2 = //td[contains(text(), '%StyleNo%')]//parent::tr//child::td[contains(text(),'%StyleColor%')]//parent::tr//child::td[@class='stnd_rtg f f12 td_gray pad_rt']
						Xpath2 = //td[contains(text(), '%StyleNo%')]//parent::tr//child::td[contains(text(),'%StyleColor%')]//parent::tr//child::input[@value='%StyleQty%']//parent::td//following-sibling::td[@class='stnd_rtg f f12 td_gray pad_rt']
						
						
						Amount := driver.FindElementByXPath(Xpath2).Attribute("innerText")
						

						;~ /* 알파벳과 숫자만 저장 (알페벳과 숫자 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
						Amount := RegExReplace(Amount, "[^a-zA-Z0-9\.]", "")
						*/

						
						MsgBox, % "Amount is " . Amount

							
						; 만약 Amount 값이 있다면 아래 if문 실행해서 현재 배열값 초기화 시키고 체크박스 클릭하기
						;~ if(Amount != 0.00){
						if(Amount){
							
							; 체크박스 클릭한 아이템은 배열에서 지우기
							Array_StyleNo[A_Index] := ""
							Array_StylyColor[A_Index] := ""


							; StyleNo , StyleColor , StyleQty 셋 다 맞는 맞는 체크박스 체크하기
							Xpath = //td[contains(text(), '%StyleNo%')]//parent::tr//child::td[contains(text(),'%StyleColor%')]//parent::tr//child::td//child::input[@type='checkbox']
							Xpath = //td[contains(text(), '%StyleNo%')]//parent::tr//child::td[contains(text(),'%StyleColor%')]//parent::tr//child::input[@value='%StyleQty%']//parent::td//following-sibling::td//child::input[@type='checkbox']
							driver.FindElementByXPath(Xpath).click()


							MsgBox % "Style No. " . A_Index . " is " . StyleNo  "`n"  "Style Color" . A_Index . " is " . StyleColor "`n" . "StyleQty" . A_Index . " is " . StyleQty . "`n"  "IS CHECKED THEN DELETED." ; 체크되고 지워진 배열값 확인하기 위해
							
							CheckedItemExistsOrNot = 1
							
						}
						; 만약 Amount 값이 없다면 이미 백오더로 넘어간 아이템이라는 뜻
						else{

							; Amount 값이 0.00 이라면 이미 백오더로 넘어간 것. 체크를 해제해주지 않으면 아래에서 백오더 버튼 누를 때 또 백오더로 넘어가게 됨.
							; 그런 중복을 방지하기 위해서 StyleNo 와 StyleColor 모두 맞는 체크박스 체크해서 이 루프문 들어오기 전 위에서 모든 아이템을 체크했을 때 체크된 것을 풀어주기
							Xpath = //td[contains(text(), '%StyleNo%')]//parent::tr//child::td[contains(text(),'%StyleColor%')]//parent::tr//child::td//child::input[@type='checkbox']
							driver.FindElementByXPath(Xpath).click()
							
							MsgBox % "ITEM " . Array_StyleNo[A_Index] . " " . Array_StylyColor[A_Index] . " IS ON THIS PAGE, BUT IT'S ALREADY PROCESSED AS BACK ORDER, SO THIS APPLICATION IS GOING TO OPEN ANOTHER ORDER DETAILS THEN UPDATE THIS ITEM ON THERE"

						}
					}
					
				}
				
			} ; Loop % Array_StylyColor.Maxindex() 루프 끝
			
		break ; 이걸 해줘야 Style No 와 Style Color 값이 서로 맞게 돌아간다

		} ; Loop % Array_StyleNo.Maxindex() 루프 끝
		



		; Amount 값이 0.00인 체크 박스들은 위에서 체크된 체크 풀어주기
		; 이렇게 풀어주지 않으면 Amount 값이 0.00인 아이템들이 그대로 체크된채로 아래 if 문을 통해 백오더 버튼이 눌러져서 새로운 PO로 만들어질 수 있음
		; 이건 쓸데없는 동작이니 여기서 Amount 값이 0.00인 아이템들은 체크를 풀어주기
		; MaxItem_# 에 있는 값이 현재 열린 창의 최대 체크박스 갯수(아이템 갯수)니까 그 갯수만큼만 루프 돌려서 Amount 값을 확인한 뒤 만약 값이 0.00이면 체크 풀어주기
		Loop % MaxItem_#{
				
			; Amount 값이 들어있는 Xpath 읽어서 만약 값이 0.00이면 그 Amount 에 해당하는 체크박스 풀어주기
			Xpath = (//td[@class='stnd_rtg f f12 td_gray pad_rt'])[%A_Index%]
			Amount := driver.FindElementByXPath(Xpath).Attribute("innerText")
			driver.FindElementByXPath(Xpath).click()
			
			; 알파벳과 숫자만 저장 (알페벳과 숫자 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
			Amount := RegExReplace(Amount, "[^a-zA-Z0-9\.]", "")
			
			MsgBox, % "Amount is " . Amount
			
			; Amount 값이 0.00인 체크 박스 체크 된 것 체크 풀어주기
			if(Amount == 0.00){
				
				Xpath_2 := Xpath . "//following-sibling::td//child::input"
				
				; 해당 체크 박스의 상태값을 isSelected 에 저장
				isSelected := driver.FindElementByXPath(Xpath_2).isSelected()
				
				; isSelected 값이 -1이면 체크 박스에 체크가 됐다는 뜻이므로 체크를 풀어주기
				if(isSelected == -1){
					driver.FindElementByXPath(Xpath_2).click()
				}			
				
				; Amount 값이 0.00인 체크 박스가 있을때마다 값을 줄여주기
				; 만약 MaxItem_# 값이 0이 되면 아래 if 문에서 백오더 버튼 누르지 말고 그냥 크롬 창 닫기
				MaxItem_#--
			}

		}
		


		
		MsgBox, % "MaxItem_#`n" . MaxItem_#
		
		; 아이템이 한개라도 체크 됐으면 다른 아이템들 백오더로 넘기고 Order Status 상태 Fully Shipped 로 바꾸기
		if(CheckedItemExistsOrNot == 1){
			; MaxItem_# 값이 0이면 현재 열린 창에 있는 모든 아이템은 Amount 값이 0.00이라는 뜻. 그렇다면 백오더 버튼을 누르면 안되고 그냥 창을 닫아야 된다
			if(MaxItem_# != 0){

MsgBox, Ready to Click Back Order Button

				; 백오더 버튼 클릭해서 체크되지 않은(지금 다루고 있는 주문이 아닌) 아이템들 백오더로 넘기기
				Xpath = //*[@id="back_order"]
;				driver.FindElementByXPath(Xpath).click()
	
				; 백오더 버튼 클릭 후 나오는 확인 메세지창 확인 누르기
				Sleep 500
				Send, {Enter}


				; Shipping & Handling Fee 입력
				Xpath = //*[@id="ship_charge"]
				driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
				driver.FindElementByXPath(Xpath).SendKeys(ShippingFee)

				; InvoiceNo 입력
				Xpath = //*[@id="inv_no"]
				driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
				driver.FindElementByXPath(Xpath).SendKeys(InvoiceNo)
						
				; TrackingNo 입력
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[13]/td[2]/input
				driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
				driver.FindElementByXPath(Xpath).SendKeys(TrackingNo)

MsgBox, Ready to Click Update Button

				; Update 버튼 클릭
				Xpath = //*[@id="update_order"]
;				driver.FindElementByXPath(Xpath).click()

			}
		}




;		MsgBox, Close tab
		
		; 브라우저 닫기
		driver.close() ; closing just one tab of the browser
		
		
		
	} ; UpdateItems 메소드 끝
	
	
	


































}