#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include CommonWebProcessing.ahk
#Include MethodsForAll.ahk


;~ #Include ChromeGet.ahk ; 이미 열려있는 창을 사용할 수 있게 해주는 함수. 
;~ global driver



class FG extends CommonWebProcessing{
	
	;~ driver :=


	; ########################
	; 공통된 Order 처리 위한 메소드
	; 원하는 검색 url 에서 PONumber 검색하기
	; ########################	
	ProcessingCommonStepOfOrderProcessing(URL, PONumber){
		
		; 새로운 FG 창을 열기
		FG.OpenNewBrowser(URL)	

		driver := ChromeGet()
		;~ MsgBox, % driver.Window.Title "`n" driver.Url
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		
		; 현재 페이지가 FG 로그인 페이지면 Login 메소드 실행해서 로그인 후 New Orders 메뉴로 들어가기
		if(RegExMatch(driver.Url, "imU)login")){
			;~ MsgBox, It's FG login page
			FG.Login()
			
			; 로그인 뒤 다시 해당 페이지를 열기(뉴오더면 뉴오더 페이지, 전체 검색이면 전체 오더 페이지)
			driver := ChromeGet()
			driver.Get(URL)
		}


		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다


		; 검색란에 PO Number 입력하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[3]/input
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(PONumber)


		; PO Number 로 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[2]/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, po, "PO Number")  ; ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기
		

		; Input Period 상태를 Last 365 Days 로 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[1]/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 33, "Last 365 Days")  ; ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기		

		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다

		
		;~ MsgBox
		
	} ; ProcessingCommonStepOfOrderProcessing 메소드 끝
	



	; ########################
	; 아이템 업데이트 위해 오더 페이지 열기
	; ########################
	OpenOrderPageForItemUpdate(PONumber, Array_PONumber){
		

		URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소		
		
		; 패션고에서 PO Number 얻기 위해 공통으로 처리하는 과정들을 담은 메소드 실행해서 PO Number 검색결과 얻기
		FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		
		; 다시 현재 창 컨트롤 할 수 있게 얻기
		driver := ChromeGet()
		
		
		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*		
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")
		


		; MTR 로 시작되는 패션고 PO Number 뽑아서 배열에 저장
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "imU)>(MTR.*)</a>", SubPat, FoundPos + strLen(SubPat)))
			Array_PONumber.Insert(SubPat1)

/*
		Loop % Array_PONumber.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Array_PONumber[A_Index]
		}
*/


/*
		if(driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]"))
		{
			MsgBox, % driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("innerText") ; 이것만 됨
			MsgBox, % driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("innerText").Attribute("innerText")
			MsgBox, % driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("Xpath")
		}
*/


		MsgBox, Close tab
		; 브라우저 닫기
		driver.close() ; closing just one tab of the browser
		
		return, Array_PONumber
		

		
		
	}





	
	
	; ########################
	; 페이지에 아이템 업데이트 하기
	; ########################
	UpdateItems(Array_StyleNo, Array_StylyColor, ShippingFee, InvoiceNo, TrackingNo, PONumber){
		
		; 체크된 아이템이 있는지 여부 확인하는 변수 초기화하기
		CheckedItemExistsOrNot :=
		
		;~ https://vendoradmin.fashiongo.net/#/order/12551798
		
		

		URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
		
		; 패션고에서 PO Number 얻기 위해 공통으로 처리하는 과정들을 담은 메소드 실행해서 PO Number 검색결과 얻기
		FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		; 다시 현재 창 컨트롤 할 수 있게 얻기
		driver := ChromeGet()

		
		; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾아서 클릭하기
		; 예를들어 MTR1D39747D26 로 검색했으면 검색된 화면에는 MTR1D39747D26 뿐만 아니라 MTR1D39747D26-BO1 등도 같이 표시되기 때문에 딱 원하는 PO Number만 클릭하기 위해
		if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
			driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
		
		
		MsgBox, ready to update items
		
		
		
		
;		driver := ChromeGet()
;		MsgBox, % driver.Window.Title "`n" driver.Url
		
		
		
		; Order Status 상태를 읽은 후 Confirmed Orders 가 아니면 Confirmed Orders로 상태를 바꾼 후 Update 버튼 클릭
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select
		CurrentOrderStatus := driver.FindElementByXPath(Xpath).Attribute("value") ; Order Status 상태 값을 얻기
		if(CurrentOrderStatus != 2){ ; Order Status 값이 2가 아니면 if문으로 들어가서 상태 바꾸기
			driver.FindElementByXPath(Xpath).SendKeys("Confirmed Orders") ; Order Status 상태를 Confirmed Orders 로 바꾸기
			
			; Update 버튼 클릭
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[1]			
			while(driver.FindElementByXPath(Xpath).Attribute("value"))
				Sleep 100
			driver.FindElementByXPath(Xpath).click()


			
			; Successfully updated order 메세지가 없어질때까지 기다리기
			; ComObjError(false) 명령어가 없으면 Successfully updated order 메세지가 없어진 뒤에 값을 찾을 수 없다고 에러메세지 나타남
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[1]/div/div[2]/fg-toast/div/div/div/span[1]
			Sleep 500
			while(driver.FindElementByXPath(Xpath).Attribute("innerText"))
				Sleep 100

		}
		
		
		; Total 옆에 있는 전체 선택 체크박스 체크해주기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/thead/tr/th[10]/div/label/div
		driver.FindElementByXPath(Xpath).click()



		Loop % Array_StyleNo.Maxindex(){
			Loop % Array_StylyColor.Maxindex(){

				StyleNo := Array_StyleNo[A_Index]
				StyleColor := Array_StylyColor[A_Index]
				
				MsgBox % "Style No. " . A_Index . " is " . Array_StyleNo[A_Index]  "`n"  "Style Color" . A_Index . " is " . Array_StylyColor[A_Index] ; 업데이트 할 스타일 번호와 색깔 확인하기 위해


				;~ Xpath = //*[text() = '%StyleNo%']


				; 빈 배열이 아니라면 시작
				;~ if(StyleNo != ""){
				if(Array_StyleNo[A_Index] != ""){
					
					; 만약에 현재 페이지에 찾는 StyleNo 가 있다면
					;~ if(driver.FindElementByXPath(Xpath)){
					if(driver.FindElementByXPath("//*[text() = '" StyleNo "']")){
						
						; StyleNo 와 StylyColor 둘 다 맞는 아이템의 백오더 값의 위치
						Xpath2 = //a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%StyleColor%')]//following-sibling::div
						BackOrderStatus := driver.FindElementByXPath(Xpath2).Attribute("innerText")
						
						;~ MsgBox, % "Back Order is " . driver.FindElementByXPath(Xpath).Attribute("innerText")
							
						; 만약 백오더 값이 없다면 if문 실행해서 현재 배열값 초기화 시키고 체크박스 클릭하기
						if(!BackOrderStatus){
							
							; 체크박스 클릭한 아이템은 배열에서 지우기
							Array_StyleNo[A_Index] := ""
							Array_StylyColor[A_Index] := ""

							; StyleNo 와 StyleColor 모두 맞는 체크박스 위치
							Xpath = //a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%StyleColor%')]//parent::td//parent::tr//child::td[@class='align-mid text-center']//child::div[@class='fg-checkbox']
							
							driver.FindElementByXPath(Xpath).click()
							
							MsgBox % "Style No. " . A_Index . " is " . StyleNo  "`n"  "Style Color" . A_Index . " is " . StyleColor "`n"  "IS CHECKED THEN DELETED." ; 체크되고 지워진 배열값 확인하기 위해
							
							CheckedItemExistsOrNot = 1
						}
						else{
							MsgBox % "ITEM " . Array_StyleNo[A_Index] . " " . Array_StylyColor[A_Index] . " IS ON THIS PAGE, BUT IT'S PROCESSED AS BACK ORDER"
						}
					}
					
				}
				
			} ; Loop % Array_StylyColor.Maxindex() 루프 끝
			
		break ; 이걸 해줘야 Style No 와 Style Color 값이 서로 맞게 돌아간다

		} ; Loop % Array_StyleNo.Maxindex() 루프 끝
		







		
		; 아이템이 한개라도 체크 됐으면 다른 아이템들 백오더로 넘기고 Order Status 상태 Fully Shipped 로 바꾸기
		if(CheckedItemExistsOrNot == 1){			
			
			; 백오더 버튼 클릭해서 체크되지 않은(지금 다루고 있는 주문이 아닌) 아이템들 백오더로 넘기기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[1]/div[2]/button[3]
	;		driver.FindElementByXPath(Xpath).click()

			; Shipping & Handling Fee 입력
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[3]/div/div[6]/div[2]/div/input
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
			driver.FindElementByXPath(Xpath).SendKeys(ShippingFee)

			; InvoiceNo 입력
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[2]/ul/li[1]/span/div/input
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
			driver.FindElementByXPath(Xpath).SendKeys(InvoiceNo)
					
			; TrackingNo 입력
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[2]/ul/li[2]/span/div/input
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
			driver.FindElementByXPath(Xpath).SendKeys(TrackingNo)

			; Save 버튼 클릭
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[1]/div/div[3]/button
	;		driver.FindElementByXPath(Xpath).click()


			; Order Status 바꾸기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select
			FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 4, "Fully Shipped Orders")
			
			; Update & Notify Buyter 버튼 클릭하기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[2]
	;		driver.FindElementByXPath(Xpath).click()			
		}


		MsgBox, Close tab
		; 브라우저 닫기
		driver.close() ; closing just one tab of the browser



		
	} ; UpdateItems 메소드 끝
	
	
	




	


	; ########################
	; 현재 페이지의 정보들 얻기 (Style No, Style Color, Pre Order 여부, cc번호, Billing Add, Shipping Add)
	; ########################		
	;~ GettingInfoFromCurrentPage(PONumber, Arr_StyleNo, Arr_#ofColors){
	GettingInfoFromCurrentPage(PONumber){
		
		M_driver := new MethodsForAll
		
		Arr_CC := object()
		Arr_Memo := object()
		
		
		; 원하는 검색 url 에서 PONumber 검색하기
		URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
		FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		
		; 다시 현재 창 컨트롤 할 수 있게 얻기
		driver := ChromeGet()

		
		; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾아서 클릭하기
		; 예를들어 MTR1D39747D26 로 검색했으면 검색된 화면에는 MTR1D39747D26 뿐만 아니라 MTR1D39747D26-BO1 등도 같이 표시되기 때문에 딱 원하는 PO Number만 클릭하기 위해
		if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
			driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 1000

		; 카드 정보 얻기 위해 CC 링크 클릭
		Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/a)
		driver.FindElementByXPath(Xpath).click()
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 1000		
		
		; 비번 입력
		Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/fg-card-info-modal/div/div[2]/div/div/div[2]/fieldset/div/span/div/input)
		driver.FindElementByXPath(Xpath).SendKeys("1234").sendKeys(driver.Keys.ENTER)
		Sleep 500
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 1000		

		; CC Type
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/ul/li[1]/span[2]
		CCType := driver.FindElementByXPath(Xpath).Attribute("innerText")
		
		if CCType contains visa
			CCType = v
		else if CCType contains master
			CCType = m
		else if CCType contains american
			CCType = a
		else if CCType contains discover
			CCType = d
		
		Arr_CC[3] := CCType
		
		
		; Name
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[1]/ul/li[2]/span[2]
		Name := driver.FindElementByXPath(Xpath).Attribute("innerText")
		StringUpper, Name, Name ; 대문자로 바꾸기
		Arr_CC[1] := Name
		
		; CC #
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/ul/li[2]/span[2]
		CC# := driver.FindElementByXPath(Xpath).Attribute("innerText")
		StringReplace, CC#, CC#, -, , All ; 대쉬(-) 없애기
		Arr_CC[2] := CC#
		
		; CVV
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/ul/li[2]/span[4]
		Arr_CC[5] := driver.FindElementByXPath(Xpath).Attribute("innerText")		
		
		; Exp
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/ul/li[3]/span[2]		
		ExpDate := driver.FindElementByXPath(Xpath).Attribute("innerText")
		
		ExpDate := M_driver.Refine_ExpDate(ExpDate)
		Arr_CC[4] := ExpDate
		
		
		
		; Buyer Notes	
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[1]
		Arr_Memo[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_Memo[1]
		StringUpper, Data, Data
		Arr_Memo[1] := Data
		
		; Additional Info	
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[2]
		Arr_Memo[2] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_Memo[2]
		StringUpper, Data, Data
		Arr_Memo[2] := Data
		
		; Staff Notes (Internal Use Only)
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[2]/div/textarea
		Arr_Memo[3] := driver.FindElementByXPath(Xpath).Attribute("textContent")
		Data := Arr_Memo[3]
		StringUpper, Data, Data
		Arr_Memo[3] := Data
		
		
		

		
		; Billing Address		
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[1]/ul/li[3]/span[2]
		BillingAdd := driver.FindElementByXPath(Xpath).Attribute("innerText")

		; Shipping Address		
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[2]/ul/li[3]/span[2]
		ShippingAdd := driver.FindElementByXPath(Xpath).Attribute("innerText")		

		Arr_BillingADD := StrSplit(BillingAdd, ",")  ; 콤마 나올때마다 문자열 나눠서 Arr_BillingADD 배열에 넣기
		Arr_ShippingADD := StrSplit(ShippingAdd, ",")  ; 콤마 나올때마다 문자열 나눠서 Arr_BillingADD 배열에 넣기
		

		; Arr_BillingADD2 찾아서 배열 5번째에 넣고 1번째에는 ADD1만 남기기
		UnquotedOutputVar = im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)
		Arr_BillingADD[5] := M_driver.FindAdd2_In_Add1(Arr_BillingADD[1], UnquotedOutputVar) ; Arr_BillingADD[1] 에 들어있는 전체 주소를 넘겨서 ADD2 만 Arr_BillingADD[5] 에 넣기
		Arr_BillingADD[1] := M_driver.DeleteAdd2_In_Add1(Arr_BillingADD[1], UnquotedOutputVar) ; 전체 주소 중 ADD2를 지운뒤 Arr_BillingADD[1]에 넣기
		
;		MsgBox, % "add1 : " Arr_BillingADD[1] . "`n" . "add2 : " . Arr_BillingADD[5]


		; ZIP 찾아서 배열 6번째에 넣고 3번째에는 State(州)만 넣기
		UnquotedOutputVar = im)(\d.*)
		Arr_BillingADD[6] := M_driver.FindAdd2_In_Add1(Arr_BillingADD[3], UnquotedOutputVar) ; Arr_BillingADD[3] 에 있는 State(州) + Zip 을 Zip만 Arr_BillingADD[6] 에 넣기
		Arr_BillingADD[3] := M_driver.DeleteAdd2_In_Add1(Arr_BillingADD[3], UnquotedOutputVar) ; Arr_BillingADD[3] 값에서 Zip 지운 뒤 State(州) 만 Arr_BillingADD[3] 에 넣기
		
		
;		MsgBox, % "State(州) : " Arr_BillingADD[3] . "`n" . "Zip : " . Arr_BillingADD[6]




		; Arr_ShippingADD2 찾아서 배열 5번째에 넣고 1번째에는 ADD1만 남기기
		UnquotedOutputVar = im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)
		Arr_ShippingADD[5] := M_driver.FindAdd2_In_Add1(Arr_ShippingADD[1], UnquotedOutputVar) ; Arr_ShippingADD[1] 에 들어있는 전체 주소를 넘겨서 ADD2 만 Arr_ShippingADD[5] 에 넣기
		Arr_ShippingADD[1] := M_driver.DeleteAdd2_In_Add1(Arr_ShippingADD[1], UnquotedOutputVar) ; 전체 주소 중 ADD2를 지운뒤 Arr_BillingADD[1]에 넣기


		; ZIP 찾아서 배열 6번째에 넣고 3번째에는 State(州)만 넣기
		UnquotedOutputVar = im)(\d.*)
		Arr_ShippingADD[6] := M_driver.FindAdd2_In_Add1(Arr_ShippingADD[3], UnquotedOutputVar) ; Arr_ShippingADD[3] 에 있는 State(州) + Zip 을 Zip만 Arr_ShippingADD[6] 에 넣기
		Arr_ShippingADD[3] := M_driver.DeleteAdd2_In_Add1(Arr_ShippingADD[3], UnquotedOutputVar) ; Arr_ShippingADD[3] 값에서 Zip 지운 뒤 State(州) 만 Arr_ShippingADD[3] 에 넣기
		
		
		; Phone Number
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[1]/ul/li[4]/span[1]
		PhoneNumber := driver.FindElementByXPath(Xpath).Attribute("innerText")
		
		PhoneNumber := RegExReplace(PhoneNumber, "[^0-9]", "") ;숫자만 저장
		Arr_BillingADD[7] := PhoneNumber


		Arr_Temp := Arr_BillingADD.Clone()
		
		
		Arr_BillingADD[1] := Arr_Temp[1]
		Arr_BillingADD[2] := Arr_Temp[5]
		Arr_BillingADD[3] := Arr_Temp[2]
		Arr_BillingADD[4] := Arr_Temp[3]
		Arr_BillingADD[5] := Arr_Temp[6]
		Arr_BillingADD[6] := Arr_Temp[4]
		Arr_BillingADD[6] := Arr_Temp[4]
		Arr_BillingADD[7] := Arr_Temp[7]
		



		/*
			Arr_BillingADD[1] -> ADD1
			Arr_BillingADD[2] -> City
			Arr_BillingADD[3] -> State
			Arr_BillingADD[4] -> United States
			Arr_BillingADD[5] -> ADD2
			Arr_BillingADD[6] -> Zip
			Arr_BillingADD[7] -> PhoneNumber
		*/
		
		
		; Arr_BillingADD 배열의 모든 값을 대문자로 바꾸고 양쪽에 공란 없애기
		Loop % Arr_BillingADD.Maxindex(){
			Data := Arr_BillingADD[A_Index]
			StringUpper, Data, Data
			Arr_BillingADD[A_Index] := Trim(Data)
			;~ MsgBox % "Element number " . A_Index . " is " . Arr_BillingADD[A_Index]
		}


		; Arr_ShippingADD 배열의 모든 값을 대문자로 바꾸고 양쪽에 공란 없애기
		Loop % Arr_ShippingADD.Maxindex(){
			Data := Arr_ShippingADD[A_Index]
			StringUpper, Data, Data
			Arr_ShippingADD[A_Index] := Trim(Data)
			;~ MsgBox % "Element number " . A_Index . " is " . Arr_ShippingADD[A_Index]
		}




;		MsgBox, % "Arr_BillingADD[1] : " . Arr_BillingADD[1] . "`nArr_BillingADD[5] : " . Arr_BillingADD[5] . "`nArr_BillingADD[2] : " . Arr_BillingADD[2] . "`nArr_BillingADD[3] : " . Arr_BillingADD[3] . "`nArr_BillingADD[6] : " . Arr_BillingADD[6]
;		MsgBox, % "Arr_ShippingADD[1] : " . Arr_ShippingADD[1] . "`nArr_ShippingADD[5] : " . Arr_ShippingADD[5] . "`nArr_ShippingADD[2] : " . Arr_ShippingADD[2] . "`nArr_ShippingADD[3] : " . Arr_ShippingADD[3] . "`nArr_ShippingADD[6] : " . Arr_ShippingADD[6]





/*
		; Add1
		Arr_BillingADD[1]
		
		; Add2
		Arr_BillingADD[5]
		
		; City
		Arr_BillingADD[2]
		
		; State
		Arr_BillingADD[3]
		
		; Zip		
		Arr_BillingADD[6]
		
		; Nation (United States)
		Arr_BillingADD[4]
		
		;~ ; Phone Number
		;~ Xpath = 
		;~ Arr_CC[1] := driver.FindElementByXPath(Xpath).Attribute("innerText")
*/





;		MsgBox, % "Arr_CC[1] : " . Arr_CC[1] . "`nArr_CC[2] : " . Arr_CC[2] . "`nArr_CC[3] : " . Arr_CC[3] . "`nArr_CC[4] : " . Arr_CC[4] . "`nArr_CC[5] : " . Arr_CC[5] . "`nArr_CC[6] : " . Arr_CC[6] . "`nArr_CC[7] : " . Arr_CC[7] . "`nArr_CC[8] : " . Arr_CC[8] . "`nArr_CC[9] : " . Arr_CC[9] . "`nArr_CC[10] : " . Arr_CC[10] . "`nArr_CC[11] : " . Arr_CC[11]
		




		return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] ; 밑에 완성하고 나면 여러개의 배열을 리턴할테니. Arr_later은 나중에 더할 배열 










		
/*

; StyleNo 값 얻기
//a[@_ngcontent-c7='' and @target='_blank']

; Arr_PreOrder
; StyleNo 에 해당하는 pro order 값(상태) 얻기
//a[contains(text(),'%StyleNo%')]//parent::div//following-sibling::div


; Arr_Colors
; StyleNo 에 해당하는 색깔들 얻기
//a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[@class='text-s']


; Arr_BackOrder
; StyleNo 와 Color 에 해당하는 Back Ordered 값(상태) 얻기
//a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%Color%')]//parent::td//child::div[@class='order-table__state']


*/
		
		; 배열 생성
		;~ Array := object()
		Arr_FG := object()
		;~ Array_1 := []


		; 현재 페이지 제어하기 위한 메소드 호출
		driver := ChromeGet()
		MsgBox, % driver.Window.Title "`n" driver.Url



		; ### StyleNo 값 얻기 #####################################################################
		; Xpath로 찾은 제일 처음 값은 메뉴바의 무슨 이상한 것이 선택이 되서 그것을 제외한 값을 배열에 넣어주려고 A_Index 쓰지 않고 i를 사용해 2부터 시작
		; Xpath 값을 읽는 건 2 번째부터 시작하지만 배열에 저장될 때는 처음부터 차례대로 저장됨		
		
		i = 2	;처음에 메뉴바의 값 제거하기 위해 2로 시작
		ii = 1	;For Style No
;		j = 1	;For Pre Order	프리 오더가 됐나 안됐나만 살펴보기 위한 것이니까 값은 언제나 1로 고정됨. 1이상의 값을 사용할 필요가 없으니까. 변수 사용 대신 그냥 1로 바꿨음
		k = 1	;For Color 
		l = 1	;For Back Order	   백오더가 됐나 안됐나만 살펴보기 위한 것이니까 값은 언제나 1로 고정됨. 1이상의 값을 사용할 필요가 없으니까
		m = 1	;Xpath 가 같은 여러 값을 읽는 모든 Xpath를 1부터 시작시키기 위해. 그렇기 때문에 루프가 끝날 때마다 1로 초기화 됨
		loop{
			
			Xpath = (//a[@target='_blank'])[%i%]

			;~ MsgBox, % driver.FindElementByXPath(Xpath).Attribute("innerText")
			
			;~ Arr_FG.Insert(driver.FindElementByXPath(Xpath).Attribute("innerText"))
			Arr_FG[a_index]	:= (driver.FindElementByXPath(Xpath).Attribute("innerText"))
			;~ MsgBox, % a_index . " Arr_FG is " . Arr_FG[a_index]


			; StyleNo 에 해당하는 pro order 값(상태) 얻기
			StyleNo := % Arr_FG[a_index]			
			Xpath = //a[contains(text(),'%StyleNo%')]//parent::div//following-sibling::div			
			Arr_FG%a_index%_1 := driver.FindElementByXPath(Xpath).Attribute("innerText")			
			MsgBox, % "Pre Order InnerText of " . StyleNo . " is" . Arr_FG%a_index%_1


			; ### StyleNo 에 해당하는 색깔들 얻기 ############################################################
			ii := % a_index
			Color_LOC := % Arr_FG[ii]
			
			Loop{
				Xpath = (//a[contains(text(),'%Color_LOC%')]//parent::div//parent::td//following-sibling::td//child::div[@class='text-s'])[%m%]
				Arr_FG%ii%_1_%k% := driver.FindElementByXPath(Xpath).Attribute("innerText")
				MsgBox, % a_index . " Color of " . StyleNo . " is " . Arr_FG%ii%_1_%k%




				; ### StyleNo 와 Color 에 해당하는 Back Ordered 값(상태) 얻기 ###################################################
				RealColor := % Arr_FG%ii%_1_%k%
				
				Xpath_BO = (//a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%RealColor%')]//parent::td//child::div[@class='order-table__state'])
				Arr_FG%ii%_1_%k%_1 := driver.FindElementByXPath(Xpath_BO).Attribute("innerText")
				MsgBox, % "Style No : " . StyleNo . "`n" . RealColor . " of Back Order Status is " . Arr_FG%ii%_1_%k%_1
				; ### StyleNo 와 Color 에 해당하는 Back Ordered 값(상태) 얻기 ### 끝				




				; Color 값 찾는 Xpath 에서 더이상 Color 값이 없으면 루프 끝내고 나가기
				if( Arr_FG%ii%_1_%k% == ""){
					Arr_FG%ii%_1_%k%.remove(a_index) ; 마지막은 값이 없으니까 지워주기
					MsgBox, % "Empty value which placed at the end of Color " . Arr_FG[ii]_1_[k] . " is deleted"
					
					break
				}				
				
				m++
			}
			; ### StyleNo 에 해당하는 색깔들 얻기 ### 끝
			
			

			; Xpath 로 찾은 값에서 더이상 Style No 값이 없으면 루프 끝내고 나가기
			if(Arr_FG[a_index] == ""){
				Arr_FG.remove(a_index) ; 마지막은 값이 없으니까 지워주기
;				MsgBox, % "The last value of Style No " . Arr_FG[a_index] . " is deleted"
				MsgBox, % "There is no more Style No of This page. Last empty value placing at the end of Array of Style No is deleted"
				break
			}
			i++
			m = 1	;다음 Xpath를 위해 1로 초기화 됨
		}
		; ### StyleNo 값 얻기 ### 끝
		
		
		
		MsgBox, out
		
Loop,% Arr_FG.MaxIndex()
	Msgbox,% Arr_FG[a_index]

MsgBox, pause

		
		
/* 배열로부터 읽기
Array:=[1,3,"ㅋㅋ"]
MsgBox, % "The method value of Array.MaxIndex is " . Array.MaxIndex()
Loop,% Array.MaxIndex()
	Msgbox,% Array[a_index]
*/

/*
		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")
		

		; 현재 페이지의 Style No 모두 뽑아서 배열에 저장
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "imU)item/detail/.*>(.*)</a>", SubPat, FoundPos + strLen(SubPat)))
		{
			Array.Insert(SubPat1)
			;~ MsgBox, % SubPat1
			
			;~ Array%A_Index% := SubPat1
			
			MsgBox, % "A_Index is " . A_Index
			MsgBox, % "Value of SubPat1 is " . SubPat1
			MsgBox, % "Value in Array is : " . Array%A_Index%
			
		}
*/

MsgBox, % "Value in Array is : " . Array[1]

MsgBox, % Array.MaxIndex()
Loop,% Array.MaxIndex(){
	MsgBox, in
	Msgbox,% Array[a_index]
}


		;~ /* 배열로부터 읽기 두 번째 방법
		for index, element in Array
		{
			MsgBox % "Element number " . index . " is " . element
			

			; StyleNo 에 해당하는 pro order 값(상태) 얻기
			;~ Xpath = //a[contains(text(),'%StyleNo%')]//parent::div//following-sibling::div
			;~ Arr_PreOrder := driver.FindElementByXPath(Xpath2).Attribute("innerText")
			
			;~ MsgBox, % StyleColor
				
			
		}
		*/

		/* 배열로부터 읽기 첫 번째 방법
		Loop % Array.Maxindex(){
			
			MsgBox, reading
			MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
		}
		*/
				
		

		
		; 현재 페이지의 옷 색깔은 모두 추출해서 배열에 저장
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "imU)text-s.*>(.*)</div>", SubPat, FoundPos + strLen(SubPat)))
;			<div _ngcontent-c12="" class="text-s">OLIVE</div>
		{
			Arr_#ofColors.Insert(SubPat1)
			MsgBox, % SubPat1
		}		
		
		;~ /* 배열로부터 읽기 두 번째 방법
		for index, element in Arr_#ofColors
		{
			MsgBox % "Element number " . index . " is " . element
			

			; StyleNo 에 맞는 Color 값 얻기
			;~ Xpath2 = //a[contains(text(),'%element%')]//parent::div//parent::td//following-sibling::td//child::div[@class='text-s']
			;~ StyleColor := driver.FindElementByXPath(Xpath2).Attribute("innerText")
			
			;~ MsgBox, % StyleColor
				
			
		}
		*/
		

		
		
	}



















































	
	
	; ########################
	; 뉴오더 처리
	; ########################
	NewOrderProcessing(PONumber){


		URL = https://vendoradmin.fashiongo.net/#/order/orders/new ; 뉴오더 검색창 주소
		
		; 패션고에서 PO Number 얻기 위해 공통으로 처리하는 과정들을 담은 메소드 실행해서 PO Number 검색결과 얻기
		FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		; 다시 현재 창 컨트롤 할 수 있게 얻기
;		driver := ChromeGet()
		

		; 검색된 PO Number 클릭해서 페이지로 들어가기
		AlreadyProcessedPONumberOrNot := FG.FindTextAndClickIt(PONumber)


		; AlreadyProcessedPONumberOrNot 값이 1이면 PO 번호로 검색된 결과가 없다는 뜻, 즉 이미 처리된 주문이므로 메소드 빠져나가기
		if(AlreadyProcessedPONumberOrNot == 1){
			MsgBox no result
			return AlreadyProcessedPONumberOrNot
		}


	
		; Order Status 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 2, "Confirmed Orders")
		
		; Update & Notify Buyter 버튼 클릭하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[2]
		driver.FindElementByXPath(Xpath).click()
				
		; Send Mail 버튼 클릭하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/fg-notify-buyer-modal/div/div[2]/div/div/div[2]/fieldset/div[5]/button[2]
;		driver.FindElementByXPath(Xpath).click()




		
		; Buyer Notes 내용 CustomerNoteOnWeb.txt 파일에 저장하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[1]
		FG.WringValueIntoAFile(Xpath, "CustomerNoteOnWeb")

		; Additional Info 내용 AdditionalInfo.txt 파일에 저장하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[2]
		FG.WringValueIntoAFile(Xpath, "AdditionalInfo")

		; Staff Notes 내용 StaffOnlyNote.txt 파일에 저장하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[2]/div/textarea
		FG.WringValueIntoAFile(Xpath, "StaffOnlyNote")
		

		
	}









	; ########################
	; DropDown 메뉴 값을 읽은 후 argument 로 받은 값이 아니면 받은 값으로 바꾸기
	; ########################	
	compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, StatusToBePreferred, PreferredStatusValue){
		
		driver := ChromeGet()
		
		; DropDown Manu 상태를 읽은 후 원하는 상태가 아니면 원하는 상태로 바꾸기
		CurrentStatus := driver.FindElementByXPath(Xpath).Attribute("value") ; 상태 값을 얻기
		
		;~ MsgBox, % CurrentStatus
		
		if(CurrentStatus != StatusToBePreferred){ ; 상태값이 원하는 값이 아니면 if문으로 들어가서 원하는 상태로 바꾸기
			driver.FindElementByXPath(Xpath).SendKeys(PreferredStatusValue)
		}
	}



	; ########################	
	; 빈칸에 값 입력 후 엔터 누르기
	; ########################	
	PutAValueInABlank(Xpath, Value){
		
		driver := ChromeGet()
		
		driver.FindElementByXPath(Xpath).SendKeys(Value)
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.Enter)
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS)
	}



	; ########################
	; 현재 페이지에서 원하는 키워드 검색 후 찾았으면 클릭하기, 못 찾았으면 1 리턴하기
	; ########################	
	FindTextAndClickIt(WantToFindText){
		
		driver := ChromeGet()
		
		Xpath = //*[text() = '%WantToFindText%']
		
		if(driver.FindElementByXPath(Xpath))
			driver.FindElementByXPath(Xpath).click()
		else
			return 1
	}



	; ########################
	; Xpath 와 저장될 파일 이름을 받은 후 Xpath 에 있는 내용 받은 파일명에 저장하기
	; ########################	
	WringValueIntoAFile(Xpath, FileNameToBeWritten){
		
		driver := ChromeGet()

		value := driver.FindElementByXPath(Xpath).Attribute("value")
		StringUpper, value, value  ; 대문자로 바꾸기
		
		;~ MsgBox, % value
		
		; value 안에 있는 내용 FileNameToBeWritten로 받은 파일이름에 저장하기
		FileAppend, %value%, %A_ScriptDir%\CreatedFiles\%FileNameToBeWritten%.txt
	}


	
	

	; ########################
	; 열려있는 브라우저 창 닫기
	; ########################
	ClosingCurrentlyOpenedBrowser(){
		
		driver := ChromeGet()
		;~ MsgBox, % driver.Window.Title "`n" driver.Url
		
		;~ driver.quit() ; 이거 작동하지 않음 closing all the browsers
		driver.close() ; closing just one tab of the browser
	}
	




	
	
} ; 클래스의 끝
