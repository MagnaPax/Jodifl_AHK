#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include CommWeb.ahk
#Include MethodsForAll.ahk
#Include ChromeGet.ahk
#Include Quickselect.ahk
#Include HasVal.ahk


;~ #Include ChromeGet.ahk ; 이미 열려있는 창을 사용할 수 있게 해주는 함수. 
global Progress



class FG extends CommWeb{
	
	;~ driver :=


	; ########################
	; 공통된 Order 처리 위한 메소드
	; 원하는 검색 url 에서 PONumber 검색하기
	; ########################	
	ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder){

		; 새로운 FG 창을 열기
		CommWeb.OpenNewBrowser(URL)

		driver := ChromeGet()
		;~ MsgBox, % driver.Window.Title "`n" driver.Url
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		
		; 현재 페이지가 FG 로그인 페이지면 Login 메소드 실행해서 로그인 후 New Orders 메뉴로 들어가기
		if(RegExMatch(driver.Url, "imU)login")){
			;~ MsgBox, It's FG login page
			;~ FG.Login()
			CommWeb.Login(URL)

/* ; 이게 되면 좋을텐데 이상하게 TIMEOUT 에러가 나서 그냥 로그인 한 창 닫고 다시 시작하는것으로 바꿨음
			; 로그인 뒤 다시 해당 페이지를 열기(뉴오더면 뉴오더 페이지, 전체 검색이면 전체 오더 페이지)
			driver := ChromeGet()
			driver.Get(URL)
*/
		}

		
		driver := ChromeGet()


		; 뉴오더 고객일때 실행
		if(IsItFromNewOrder == 1){

			; New Order 화면 보여지는 갯수 50개로 고치기
	;		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[3]/div/fg-per-page/div/div/select
	;		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 3, "50")  ; ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기
	;		Sleep 150


			; 페이지 숫자 알아내서 #ofPages 변수에 넣기
			#ofPages_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[2]/fg-pagination/ul/li[3]/div/div[2]/div
			#ofPages := driver.FindElementByXPath(#ofPages_Xpath).Attribute("textContent")
			RegExMatch(#ofPages, "imU)([0-9])", SubPat)
			#ofPages := SubPat
			
;			MsgBox, % "#ofPages : " . #ofPages . "`n`n" . "SubPat1 : " . SubPat1
			
			; 페이지가 1페이지 이상이면 페이지 번호 넣는 창에 페이지 숫자 입력해서 이동하기
			; 지금 찾는 고객은 마지막 페이지에 있을 확률이 높으므로 페이지 숫자 넣어서 일단 마지막 페이지로 이동함
			;~ if(TheBlankOfCurrentPage_Xpath != 1){
			if(#ofPages != 1){
;				MsgBox, #ofPages IS MORE THAN 1
				TheBlankOfCurrentPage_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[2]/fg-pagination/ul/li[3]/div/div[1]/div/input
				driver.FindElementByXPath(TheBlankOfCurrentPage_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(#ofPages).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 #ofPages 입력 후 엔터쳐서 이동하기	
			}

			
			
			; PREV 버튼의 Xpath PreviouseButton_Xpath 변수에 넣기
			NextButton_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[2]/fg-pagination/ul/li[4]/button
			PreviouseButton_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[2]/fg-pagination/ul/li[2]/button
			
			
			; 화면에서 Customer PO 찾고 그 값에 해당하는 url 들어가기
			Loop{
				; 현재 화면에서 고객 po 찾았으면 실행
				if(driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]")){
					
					; 화면에서 Customer PO 에 있는 outer HTML 값 outerHTMLofTheCustomerPO 에 저장하기
					outerHTMLofTheCustomerPO := driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("outerHTML")			
					;~ outerHTMLofTheCustomerPO := driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").Attribute("innerText")
					
;					MsgBox, % "outerHTMLofTheCustomerPO IS `n" . outerHTMLofTheCustomerPO
					
					; outerHTML 에서 url 페이지 들어가기 위해 Customer PO 의 고유 아이디만 uniqueIDofCustPO 변수에 저장
					UnquotedOutputVar = imU)href="#/order/(.*)">
					RegExMatch(outerHTMLofTheCustomerPO, UnquotedOutputVar, SubPat)

					
;					MsgBox, % "THE UNIQUE NUMBER OF CUST'S ORDER PAGE ID IS : " . SubPat1
					


					; 위의 동작에서 얻은 고유 아이디에 기본 url 주소를 붙이면 해당 Customer PO 의 주문 페이지가 된다.
					; 그곳으로 이동한다
					; SubPat1 를 찾았을 때만 이동한다
					if(SubPat1){
						URLofCustPO := "https://vendoradmin.fashiongo.net/#/order/" . SubPat1				
						driver.Get(URLofCustPO)
					}
					
					
					
					; 가끔 outerHTMLofTheCustomerPO 읽을때 
					; <a _ngcontent-c10="" href="#/order/13408498">MTR1F38167D12</a>
					; 이런 형식이 아닌 다른 이상한? 값의 형태가 있을때가 있다.
					; 그런 값에서는 RegExMatch 이용해서 주문페이지의 고유 아이디를 추출할 수 없다
					; 또한 PONumber 에 해당하는 링크를 클릭할 수도 없다
					; 왜 그런지 모르겠다. 그냥 에러 메세지 띄우고 직접 클릭하라고 하는게 지금까지 찾아낸 유일한 해법이다.
					
					; MTR1F38167D12 <- 이게 클릭이 안 된다
					
					if(!SubPat1){
						SoundPlay, %A_WinDir%\Media\Ring02.wav
						MsgBox, 262144, Title, AN ERROR OCCURRED. PLEASE CLICK THE CUSTOMER CODE BELOW MANUALLY.`n`n`n%PONumber%
						;~ driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").click()
						;~ driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
						driver := ChromeGet()
					}


;~ /*					
					; 오더 페이지에 제대로 들어갔는지 확인하기
					; 현재 열린 창과 고객의 주문 페이지가 맞지 않으면
					; 현재 창 리프레쉬 한 뒤 continue 로 루프 다시 시작해보기
					; SubPat1 를 찾았을 때만 제대로 들어갔는지 확인한다
					; 찾지 못했을 땐 위에서 그냥 수동으로 클릭했다.
					if(SubPat1){
						driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
						Sleep 1000
						CurrentURL := driver.Url

						if(CurrentURL != URLofCustPO)
						{
							MsgBox, 262144, url not matched, IT'S NOT CUSTOMER'S ORDER PAGE`nRESTART LOADING THE ORDER PAGE AGAIN`n`nCurrentURL : `n%CurrentURL%`nURLofCustPO : `n%URLofCustPO%
							driver.refresh()
							FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder)
							;~ continue
						}
					}

*/

					; 오더 페이지 제대로 들어갔으니 리턴으로 메소드 끝내기
					return
					
					
		;			MsgBox, % "SubPat : " . SubPat . "`n`n" . "SubPat1 : " . SubPat1 . "`n`n" . "outerHTMLofTheCustomerPO : " . outerHTMLofTheCustomerPO

				}
/*				
				; 현재 화면에서 Customer PO 못 찾았으면 PREV 버튼 눌러서 페이지 이동한 뒤 LOOP 를 통해 이동한 화면에서 Customer PO 다시 찾아보기
				else if(driver.FindElementByXPath(PreviouseButton_Xpath).isEnabled()){
					Sleep 100
					driver.FindElementByXPath(PreviouseButton_Xpath).click()
					driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
					Sleep 100
				}
*/				

				; #ofPages 변수가 1이라는 뜻은 위의 코드를 통해 첫 번째 페이지에서도 고객 PO를 못 찾았다는 뜻이므로 루프 돌 필요가 없으니 break 사용해서 나가기
				else if(#ofPages == 1){
					MsgBox, 262144, NO PO on this page, IT'S PAGE # 1 BUT CANNOT FINE %PONumber% ON NEW ORDER PAGES`n`nPLEASE FIND THE ORDER PAGE AND CLICK IT.
					return
				}	
				

				; 현재 화면에서 Customer PO 못 찾았으면 현재 페이지보다 이전 페이지를 입력한 뒤 GO 버튼 눌러서 이동하기
				else if(#ofPages){
					Sleep 100
				
					#ofPages--  ; 현재 페이지에서 1 빼기
					
					; 페이지 번호 입력 후 엔터
					TheBlankOfCurrentPage_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[2]/fg-pagination/ul/li[3]/div/div[1]/div/input
					driver.FindElementByXPath(TheBlankOfCurrentPage_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(#ofPages).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 #ofPages 입력 후 엔터쳐서 이동하기	
						
					driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
					Sleep 100
				}
				
			} ; loop 끝

		}




		; CUSTOMER PO 입력창이 나타났으면 검색란에 PO Number 입력
		; 없으면 ProcessingCommonStepOfOrderProcessing 메소드 재귀호출하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[3]/input
		Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[3]/input)[2]

		i = 0
		Loop{
			Sleep 500
			
			if(driver.FindElementByXPath(Xpath).isDisplayed()){
				driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(PONumber)
				Sleep 500
				break
			}
			else if(i == 4){
				MsgBox, 입력칸 못 찾았음, %i%
				FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder)
			}
			
			i++			
			
		}


		; PO Number 로 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[2]/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, po, "PO Number")  ; ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기
		Sleep 500
		

		; Input Period 상태를 Last 365 Days 로 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[1]/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 33, "Last 365 Days")  ; ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기
		Sleep 500

		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다

		
		
		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*		
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")

/*
		; MTR 로 시작되는 패션고 PO Number 뽑아서 배열에 저장		
		Array_PONumber := object() ; 사용할 배열 선언
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "imU)>(MTR.*)</a>", SubPat, FoundPos + strLen(SubPat)))
			Array_PONumber.Insert(SubPat1)
*/

		FoundPos := RegExMatch(HTML_Source, "imU)>(MTR.*)</a>")
		
;		MsgBox, % ErrorLevel

		Sleep 200

		; 현재 화면에서 검색된 Customer PO 가 없어서 배열에 값이 저장되어 있지 않으면
		; 현재 페이지를 refresh 한 뒤 ProcessingCommonStepOfOrderProcessing 메소드 재귀호출 하기
;		if(Array_PONumber[1] == ""){
;		if(IsThereSearchedResultOnThisPage == ""){
		;~ if(SubPat1 == ""){
		if ErrorLevel
		{
			
			MsgBox, 262144, Alert, NO SEARCHED RESULT ON THIS PAGE`n`nREFRESH CHROME BROWSER THEN CONTINUE SERARCH CUSTOMER PO AGAIN.

			driver.refresh()
			Sleep 4000


			FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder)
		}
		
		
		return	
		
	} ; ProcessingCommonStepOfOrderProcessing 메소드 끝
	



	; ########################
	; 아이템 업데이트 위해 오더 페이지 열기
	; 현재 화면에 있는 모든 Customer PO 얻기
	; ########################
	;~ OpenOrderPageForItemUpdate(PONumber, Array_PONumber){
	OpenOrderPageForItemUpdate(PONumber){
		
;		MsgBox, in OpenOrderPageForItemUpdate Method


/*
		; 창을 새로 열고 시작하고 싶으면 이것을 사용하면 된다
		URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소		
		
		; 패션고에서 PO Number 얻기 위해 공통으로 처리하는 과정들을 담은 메소드 실행해서 PO Number 검색결과 얻기
		FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber)
*/		

		; 검색되어 나온 고객 PO 모두 저정하기 위한 배열 선언
		Array_PONumber := object()

		; 프로그램 처음 시작할 때 크롬창을 열고 시작했기 때문에 그 창 다시 사용하기 위해
		driver := ChromeGet()


/*		; ProcessingCommonStepOfOrderProcessing 메소드 실행되서 이미 현재 페이지에서는 검색된 CUSTOMER PO 가 있기 때문에 아래 코드 사용 안하도록 했음

		; 검색란에 PO Number 입력하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[3]/input
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(PONumber)
		Sleep 500


		; PO Number 로 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[2]/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, po, "PO Number")  ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기
		Sleep 500
		

		; Input Period 상태를 Last 365 Days 로 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[1]/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 33, "Last 365 Days")  ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기
		Sleep 500

		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
*/		
		
		
		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*		
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")




		; MTR 로 시작되는 패션고 PO Number 뽑아서 배열에 저장
		UnquotedOutputVar = imU)/order/\d*">(MTR.*)</a>
		
		
		
		FoundPos = 1
		;~ while(FoundPos := RegExMatch(HTML_Source, "imU)>(MTR.*)</a>", SubPat, FoundPos + strLen(SubPat)))
		while(FoundPos := RegExMatch(HTML_Source, UnquotedOutputVar, SubPat, FoundPos + strLen(SubPat)))
			Array_PONumber.Insert(SubPat1)


/*
		; MTR 로 시작되는 패션고 PO Number 뽑아서 배열에 저장
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, "imU)>(MTR.*)</a>", SubPat, FoundPos + strLen(SubPat)))
			Array_PONumber.Insert(SubPat1)
*/

		; 현재 화면에서 검색된 Customer PO 가 없어서 배열에 값이 저장되어 있지 않으면
		; 현재 페이지를 refresh 한 뒤 OpenOrderPageForItemUpdate 메소드 재귀호출 하기
		if(Array_PONumber[1] == ""){
			
			MsgBox, 262144, Alert, NO SEARCH RESULT`n`nREFRESH THIS PAGE THEN SEARCH CUSTOMER PO AGAIN TO GET ALL CUSTOMER POS TO UPDATE ITEMS IF OK BUTTON CLICKED.

			driver.refresh()
			Sleep 4000

			FG.OpenOrderPageForItemUpdate(PONumber)
		}


/*
		; PO 제대로 잘 저장했는지 확인하는 코드
		Loop % Array_PONumber.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Array_PONumber[A_Index]
		
			if Array_PONumber[A_Index] not contains PONumber
			{
				MsgBox, It's not correct PO
			}
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


		
		; 브라우저 닫기
;		driver.close() ; closing just one tab of the browser
		
		return, Array_PONumber	
		
	}













	; pre order 날짜 있는 아이템들 찾아서 날짜가 같은 것들 체크한 뒤 백오더 버튼 누르고 해당하는 날짜 입력한 뒤 다른 날짜 찾아서 같은 동작 반복하는 메소드
	; UpdateItems 메소드 안에서 호출된다		
	ProcessOfItemsHavingPreOrderDates(Array_AvailableDate, Array_StyleNo, Array_StylyColor, Array_StylyQty){
			
		driver := ChromeGet()		
		
			
		; Array_AvailableDate 배열의 available date 날짜 순서대로 정렬 후 Array_AvailableDate_Sorted 배열에 저장
		QS := new Quickselect
		Loop, % Array_AvailableDate.MaxIndex(){
			;~ MsgBox, % QS.Select(Array_AvailableDate, A_Index)
			Array_AvailableDate_Sorted.Insert(QS.Select(Array_AvailableDate, A_Index))
		}
			
/*
		; Array_AvailableDate_Sorted 배열에 값이 어떻게 들어갔는지 확인키 위해. 값을 확인하기 위한 것으로서 실제 사용땐 영향을 미치는 동작들이 아니다. 실행을 위해 주석처리 해도 무방함
		Loop, % Array_AvailableDate_Sorted.MaxIndex()
		{
			MsgBox % "Element number " . A_Index . " is " . Array_AvailableDate_Sorted[A_Index]
			;~ MsgBox, % driver.FindElementByXPath("//*[text() = '" TheAvailableDate "']//parent::td//child::a").Attribute("innerText")
				

;			MsgBox, % driver.FindElementByXPath("//*[text() = '" Array_AvailableDate_Sorted[A_Index] "']//parent::td//child::a").Attribute("innerText") ; available date 에 연관된 style 값. 첫 번째 값 					
;			MsgBox, % driver.FindElementByXPath("(//*[text() = '" Array_AvailableDate_Sorted[A_Index] "'])[2]//parent::td//child::a").Attribute("innerText") ; available date 에 연관된 style 값. 두 번째 값

				
		}
				
			
			
		Loop % Array_AvailableDate.Maxindex(){
				
			TheAvailableDate := Array_AvailableDate[A_Index]
			
			;~ MsgBox % "Element number " . A_Index . " is " . Array_AvailableDate[A_Index]
			MsgBox % "Element number " . A_Index . " is " . TheAvailableDate
				
				
	
			;~ if(driver.FindElementByXPath("//*[text() = '" TheAvailableDate "']")){
				;~ MsgBox, % driver.FindElementByXPath("//*[text() = '" TheAvailableDate "']//parent::td//child::a").Attribute("innerText")
			;~ }
					

			if(driver.FindElementByXPath("//*[text() = '" Array_AvailableDate[A_Index] "']")){
				MsgBox, % driver.FindElementByXPath("//*[text() = '" Array_AvailableDate[A_Index] "']//parent::td//child::a").Attribute("innerText")
			}
					

				
		}
*/
		
		
		
		pointer_for_date = 0
		
		Loop{				
				
			#ofsamedate = 0
			Current_date := Array_AvailableDate_Sorted[++pointer_for_date]
			#ofsamedate++
				
			Loop{
				Next_date := Array_AvailableDate_Sorted[++pointer_for_date]
				#ofsamedate++
					
				MsgBox, % "Current_date : " . Current_date . "`nNext_date" . Next_date . "`n`npointer_for_date : " . pointer_for_date . "`n#ofsamedate : " . #ofsamedate
					
				if(Current_date == Next_date){						
				}
				else
					break
			}
			
			#ofsamedate--
			pointer_for_date--
			
			point_for_check := pointer_for_date - #ofsamedate + 1
				
			MsgBox, loop out
			MsgBox, % "Current_date : " . Current_date . "`nNext_date" . Next_date . "`n`npointer_for_date : " . pointer_for_date . "`n#ofsamedate : " . #ofsamedate . "`npoint_for_check : " . point_for_check
				
		}


		MsgBox, % "master(?) loop out"

		
		
		
		return
		
	} ; ProcessOfItemsHavingPreOrderDates 메소드 끝















	


	; ########################
	; 현재 페이지의 정보들 얻기 (Style No, Style Color, Pre Order 여부, cc번호, Billing Add, Shipping Add)
	; ########################		
	;~ GettingInfoFromCurrentPage(PONumber, Arr_StyleNo, Arr_#ofColors){
	GettingInfoFromCurrentPage(PONumber, IsItFromNewOrder, IsItFromExcelFile){
		
		M_driver := new MethodsForAll
		
		Arr_CC := object()
		Arr_Memo := object()
		
/*		
		driver := ChromeGet()
		
		; 지금 열려있는 크롬 창이 FG 창이면
		if(RegExMatch(driver.Url, "imU)fashiongo")){

		}
*/

		; 패션고 창이 열려있지 않으면 열기		
		; 이렇게 하면 이미 열린 창이 오토핫키에서 사용할 수 없으면 에러 나겠지만
		; 이건 어차피 나만 쓰는 거니까 이렇게 하는게 시간이 더 적게 들 수 있다.
		IfWinNotExist, FashionGo Vendor Admin - Google Chrome
		{
			; 원하는 검색 url 에서 PONumber 검색하기		
			; IsItFromNewOrder 변수값이 1이면 뉴오더 검색이므로 url을 뉴오더로 하기
			if(IsItFromNewOrder == 1){
				URL = https://vendoradmin.fashiongo.net/#/order/orders/new ; 뉴 오더 검색창 주소
				FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder)
			}
			else
			{
				URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
				FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder)
			}
			; 다시 현재 창 컨트롤 할 수 있게 얻기
			driver := ChromeGet()
			
			; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾아서 클릭하기
			driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
		}
		
		; 이미 열려있으면 그 창 다시 사용하기
		; 
		else{
			driver := ChromeGet()
;			driver.refresh()
;			MsgBox, % driver.Window.Title "`n" driver.Url

			URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
			driver.Get(URL)
			
			driver := ChromeGet()
			
			
			; 창이 열려있을때만해도 로그인 되어있었는데 그 사이에 만약 로그아웃되어 현재 페이지가 FG 로그인 페이지면 Login 메소드 실행해서 로그인 후 New Orders 메뉴로 들어가기
			if(RegExMatch(driver.Url, "imU)login")){
				
				MsgBox, IT'S LOGGED OUT, SO LOG IN AGAIN

				CommWeb.Login(URL)
				
				driver := ChromeGet()
			}

			
;			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
;			Sleep 300
			

			; 페이지 숫자 알아내서 #ofPages 변수에 넣기
			#ofPages_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[2]/fg-pagination/ul/li[3]/div/div[2]/div
			#ofPages := driver.FindElementByXPath(#ofPages_Xpath).Attribute("textContent")
			RegExMatch(#ofPages, "imU)([0-9])", SubPat)
			#ofPages := SubPat
			
;			MsgBox, % "#ofPages : " . #ofPages . "`n`n" . "SubPat1 : " . SubPat1
			
			; 페이지가 1페이지 이상이면 페이지 번호 넣을 수 있는 공란에 페이지 숫자 입력해서 이동하기
			; 지금 찾는 고객은 마지막 페이지에 있을 확률이 높으므로 페이지 숫자 넣어서 일단 마지막 페이지로 이동함
			if(TheBlankOfCurrentPage_Xpath != 1){
				
				TheBlankOfCurrentPage_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/div/div/div[2]/fg-pagination/ul/li[3]/div/div[1]/div/input
				
				; 페이지 번호 넣을 수 있는 곳이 보일때까지 루프 돌면서 기다리다가 나오면 페이지 숫자 입력해서 이동
				Loop
				{					
					if(driver.FindElementByXPath(TheBlankOfCurrentPage_Xpath).isDisplayed()){
						
						driver.FindElementByXPath(TheBlankOfCurrentPage_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(#ofPages).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 #ofPages 입력 후 엔터쳐서 이동하기	
						break
					}
					Sleep 500
				}
				
			}			
			

			; 화면에서 Customer PO 찾고 그 값에 해당하는 url 들어가기
			counter_ForSearchingCustPOonThisScreen = 0
			Loop{
				
				Xpath = //*[text() = '%PONumber%']			
				; 현재 화면에서 해당 Customer PO를 찾았으면
				if(driver.FindElementByXPath(Xpath).Attribute("innerText")){
					
;					MsgBox, 262144, Founded, FOUND THE CUST'S PO ON THIS SCREEN					
					
					; 화면에서 Customer PO 에 있는 outer HTML 값 outerHTMLofTheCustomerPO 에 저장하기
					outerHTMLofCustPO := driver.FindElementByXPath(Xpath).Attribute("outerHTML")
					
					; outerHTML 에서 url 페이지 들어가기 위해 Customer PO 의 고유 아이디만 uniqueIDofCustPO 변수에 저장					
					UnquotedOutputVar = .*/order/(\d*).*>.*
					uniqueIDofCustPO := RegExReplace(outerHTMLofCustPO, UnquotedOutputVar, "$1")  ; $1 역참조를 사용하여 Customer PO 의 고유 아이디만 돌려줍니다
					
					; 위의 동작에서 얻은 고유 아이디에 기본 url 주소를 붙이면 해당 Customer PO 의 주문 페이지가 된다.
					; 그 주문 페이지 주소를 URLofCustPO 변수에 저장
					basicURLofOrderPage = https://vendoradmin.fashiongo.net/#/order/
					URLofCustPO := basicURLofOrderPage . uniqueIDofCustPO
					
					; URLofCustPO 변수에 저장된 고객의 주문 페이지로 이동
					driver.Get(URLofCustPO)

;~ /*
					; 오더 페이지에 제대로 들어갔는지 확인하기
					; 현재 열린 창과 고객의 주문 페이지가 맞지 않으면
					; 현재 창 리프레쉬 한 뒤 continue 로 루프 다시 시작해보기
					driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
					Sleep 700					
					CurrentURL := driver.Url
						
					if(URLofCustPO != CurrentURL)
					{
						MsgBox, 262144, url not matched, IT'S NOT CUSTOMER'S ORDER PAGE`nRESTART LOADING THE ORDER PAGE AGAIN`n`nCurrentURL : `n%CurrentURL%`nURLofCustPO : `n%URLofCustPO%
;						MsgBox, IT'S CALLED FROM GettingInfoFromCurrentPage METHOD
						driver.refresh()
						FG.ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder)
						;~ continue
					}
*/					
					
					; 고객의 주문 페이지로 이동했으니 루프 끝내고 나가기
					break
				}				
				
				; 화면에서 Customer PO 를 2번 찾아보다가 못 찾았으면 PREV 버튼 눌러서 페이지 이동한 뒤 LOOP 처음으로 돌아가서 Customer PO 다시 찾아보기
				else if(counter_ForSearchingCustPOonThisScreen == 2){
					
					; PREV 버튼이 화면에 표시됐으면 버튼 클릭하기
					if(driver.FindElementByXPath(PreviouseButton_Xpath).isEnabled()){
						Sleep 100
						driver.FindElementByXPath(PreviouseButton_Xpath).click()
						driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
						Sleep 100
					}
					
					counter_ForSearchingCustPOonThisScreen = 0
				}
				
				counter_ForSearchingCustPOonThisScreen++
				
				Sleep 500
			}
			
			Sleep 1000
			
			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		}


		
		; 다시 현재 창 컨트롤 할 수 있게 얻기
;		driver := ChromeGet()

		
		
/*		
		; PO Number 링크 나타날때까지 기다림
		Xpath = //*[text() = '%PONumber%']
		Sleep 500
		while(!driver.FindElementByXPath(Xpath).Attribute("innerText"))
			Sleep 100
*/


/*
		; PO Number 링크 나타날때까지 기다리다가 나오면 클릭하기
		counter_ForSearchingCustPOonThisScreen = 0
		Loop{
			
			Xpath = //*[text() = '%PONumber%']			
			; 현재 화면에서 해당 Customer PO를 찾았으면
			IF(driver.FindElementByXPath(Xpath).Attribute("innerText")){
				
				MsgBox, 262144, Founded, FOUND THE CUST'S PO ON THIS SCREEN
				
				; 해당 Xpath 의 값을 outerHTMLofCustPO 변수에 저장한 뒤
				outerHTMLofCustPO := driver.FindElementByXPath(Xpath).Attribute("outerHTML")
				
				; Customer PO 의 고유 아이디만 uniqueIDofCustPO 변수에 저장
				UnquotedOutputVar = .*/order/(\d*).*>.*
				uniqueIDofCustPO := RegExReplace(outerHTMLofCustPO, UnquotedOutputVar, "$1")  ; $1 역참조를 사용하여 Customer PO 의 고유 아이디만 돌려줍니다
				
				; 위의 동작에서 얻은 고유 아이디에 기본 url 주소를 붙여서 URLofCustPO 변수에 저장
				basicURLofOrderPage = https://vendoradmin.fashiongo.net/#/order/
				URLofCustPO := basicURLofOrderPage . uniqueIDofCustPO
				
				; URLofCustPO 변수에 저장된 고객의 주문 페이지로 이동
				driver.Get(URLofCustPO)
				

				; 오더 페이지에 제대로 들어갔는지 확인하기
				; 현재 열린 창과 고객의 주문 페이지가 맞지 않으면
				; 현재 창 리프레쉬 한 뒤 continue 로 루프 다시 시작해보기
				driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
				sleep 700
				CurrentURL := driver.Window.Title "`n" driver.Url
					
				if(URLofCustPO != CurrentURL){
					MsgBox, 262144, url not matched, IT'S NOT CUSTOMER'S ORDER PAGE`nRESTART LOADING THE ORDER PAGE AGAIN`n`nCurrentURL : `n%CurrentURL%`nURLofCustPO : `n%URLofCustPO%
					driver.refresh()
					continue
				}
				
				
				; 고객의 주문 페이지로 이동했으니 루프 끝내고 나가기
				break
			}			
			
			; 화면에서 10번 찾아보다가 못 찾았으면 현재 페이지를 리프래쉬 한 뒤 CONTINUE 로 루프 다시 시작해서 시도해보기
			if(counter_ForSearchingCustPOonThisScreen == 10){
				driver.refresh()
				MsgBox, 262144, Retry, Recall this Method`n`n메소드 재귀호출
				FG.GettingInfoFromCurrentPage(PONumber, IsItFromNewOrder, IsItFromExcelFile)
			}
			
			counter_ForSearchingCustPOonThisScreen++
			
			Sleep 500
		}
		
		Sleep 1000
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
*/		




/*
		; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾아서 클릭하기
		; 예를들어 MTR1D39747D26 로 검색했으면 검색된 화면에는 MTR1D39747D26 뿐만 아니라 MTR1D39747D26-BO1 등도 같이 표시되기 때문에 딱 원하는 PO Number만 클릭하기 위해		
		Xpath = //*[text() = '%PONumber%']
		
		if(driver.FindElementByXPath(Xpath))
			driver.FindElementByXPath(Xpath).click()
		
		;이 코드를 써도 됨 이게 두줄밖에 안 되서 더 간단함
;		if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
;			driver.FindElementByXPath("//*[text() = '" PONumber "']").click()


		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 1000
*/		
		
		
		Sleep 1000
		
		
		; 뉴오더 처리하는 것에서 호출됐을때만 실행
		; 주문 페이지의 Order Status 컨펌으로 바꾸기
		;~ if(!IsItFromExcelFile){
		if(IsItFromNewOrder){
		
			; Order Status 바꾸기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select
			HasDropdownMenuChanged := FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 2, "Confirmed Orders")
			
			driver := ChromeGet()
			
			; 위에서 Order Status 가 Confirmed Orders 로 바뀌었으면 Update 버튼 클릭하기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[1]
			if(HasDropdownMenuChanged == 1)
				driver.FindElementByXPath(Xpath).click()
		}




		; Shipping Method 상태 알아내기
		; UPS Ground 가 아니면 
		Shipping_Method_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[2]/ul/li[1]/span/span/select
		Shipping_Method_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[2]/ul/li/span/span/select
		ShippingMethodStatus := driver.FindElementByXPath(Shipping_Method_Xpath).Attribute("value")
		





/*	; CC 입력할 때 활성화 하기

		; CC 링크 나타날때까지 기다림
		Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/a)
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[3]/div/a
		Sleep 500
		while(!driver.FindElementByXPath(Xpath).Attribute("innerText"))
			Sleep 100
		
		Sleep 1000

		; 카드 정보 얻기 위해 CC 링크 클릭
		Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/a)
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[3]/div/a
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
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[3]/div/ul/li[1]/span[2]
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
		

		; CC #
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/ul/li[2]/span[2]
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[3]/div/ul/li[2]/span[2]
		CC# := driver.FindElementByXPath(Xpath).Attribute("innerText")
		StringReplace, CC#, CC#, -, , All ; 대쉬(-) 없애기
		CC# := RegExReplace(CC#, "[^0-9]", "") ;숫자만 저장
		Arr_CC[2] := CC#
		
		; CVV
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/ul/li[2]/span[4]
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[3]/div/ul/li[2]/span[4]
		Arr_CC[5] := driver.FindElementByXPath(Xpath).Attribute("innerText")		
		
		; Exp
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[2]/div/ul/li[3]/span[2]		
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li/span[3]/div/ul/li[3]/span[2]	
		ExpDate := driver.FindElementByXPath(Xpath).Attribute("innerText")
		
		ExpDate := M_driver.Refine_ExpDate(ExpDate)
		Arr_CC[4] := ExpDate
*/		
		
		
		; Name
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[1]/ul/li[2]/span[2]
		Name := driver.FindElementByXPath(Xpath).Attribute("innerText")
		StringUpper, Name, Name ; 대문자로 바꾸기
		Arr_CC[1] := Name		
		
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
		Arr_Memo[3] := driver.FindElementByXPath(Xpath).Attribute("value")
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
		UnquotedOutputVar = im)(( unit| Suite| Ste| #| Apt| SPACE| BLDG| Building| Sujite| Sujite).*)
		;~ UnquotedOutputVar = im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite|Sujite).*)
		;~ UnquotedOutputVar = im)((\sunit|\sSuite|\sSte|\s#|\sApt|\sSPACE|\sBLDG|\sBuilding|\sSujite|\sSujite).*)
		Arr_BillingADD[5] := M_driver.FindAdd2_In_Add1(Arr_BillingADD[1], UnquotedOutputVar) ; Arr_BillingADD[1] 에 들어있는 전체 주소를 넘겨서 ADD2 만 Arr_BillingADD[5] 에 넣기
		Arr_BillingADD[1] := M_driver.DeleteAdd2_In_Add1(Arr_BillingADD[1], UnquotedOutputVar) ; 전체 주소 중 ADD2를 지운뒤 Arr_BillingADD[1]에 넣기
		
;		MsgBox, % "add1 : " Arr_BillingADD[1] . "`n" . "add2 : " . Arr_BillingADD[5]


		; ZIP 찾아서 배열 6번째에 넣고 3번째에는 State(州)만 넣기
		UnquotedOutputVar = im)(\d.*)
		Arr_BillingADD[6] := M_driver.FindAdd2_In_Add1(Arr_BillingADD[3], UnquotedOutputVar) ; Arr_BillingADD[3] 에 있는 State(州) + Zip 을 Zip만 Arr_BillingADD[6] 에 넣기
		Arr_BillingADD[3] := M_driver.DeleteAdd2_In_Add1(Arr_BillingADD[3], UnquotedOutputVar) ; Arr_BillingADD[3] 값에서 Zip 지운 뒤 State(州) 만 Arr_BillingADD[3] 에 넣기
		
		
;		MsgBox, % "State(州) : " Arr_BillingADD[3] . "`n" . "Zip : " . Arr_BillingADD[6]




		; Arr_ShippingADD2 찾아서 배열 5번째에 넣고 1번째에는 ADD1만 남기기
		UnquotedOutputVar = im)(( unit| Suite| Ste| #| Apt| SPACE| BLDG| Building| Sujite| Sujite).*)
		;~ UnquotedOutputVar = im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite|Sujite).*)
		;~ UnquotedOutputVar = im)((\sunit|\sSuite|\sSte|\s#|\sApt|\sSPACE|\sBLDG|\sBuilding|\sSujite|\sSujite).*)
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


		; 카드 정보 순서 변경하기
		Arr_Temp := Arr_BillingADD.Clone()		
		
		Arr_BillingADD[1] := Arr_Temp[1] ; ADD1
		Arr_BillingADD[2] := Arr_Temp[5] ; ADD2
		Arr_BillingADD[3] := Arr_Temp[2] ; CITY
		Arr_BillingADD[4] := Arr_Temp[3] ; STATE
		Arr_BillingADD[5] := Arr_Temp[6] ; ZIP
		Arr_BillingADD[6] := Arr_Temp[4] ; COUNTRY
		Arr_BillingADD[7] := Arr_Temp[7] ; PHONE
		

		Arr_Temp := []
		Arr_Temp := Arr_ShippingADD.Clone()
		
		
		Arr_ShippingADD[1] := Arr_Temp[1]
		Arr_ShippingADD[2] := Arr_Temp[5]
		Arr_ShippingADD[3] := Arr_Temp[2]
		Arr_ShippingADD[4] := Arr_Temp[3]
		Arr_ShippingADD[5] := Arr_Temp[6]
		Arr_ShippingADD[6] := Arr_Temp[4]
		Arr_ShippingADD[7] := Arr_Temp[7]
		



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
		
			; ADD2 는 대문자로 바꾸지 않기 위해. #를 StringUpper로 처리하면 에러 난다. ADD2 에서 #를 지운다
			if(A_Index == 2){
				StringReplace, Data, Data, #, , All
				Arr_BillingADD[A_Index] := Data
				continue
			}
						
			StringUpper, Data, Data
			Arr_BillingADD[A_Index] := Trim(Data)
			;~ MsgBox % "Element number " . A_Index . " is " . Arr_BillingADD[A_Index]
		}


		; Arr_ShippingADD 배열의 모든 값을 대문자로 바꾸고 양쪽에 공란 없애기
		Loop % Arr_ShippingADD.Maxindex(){
			
			Data := Arr_ShippingADD[A_Index]
		
			; ADD2 는 대문자로 바꾸지 않기 위해. #를 StringUpper로 처리하면 에러 난다. ADD2 에서 #를 지운다
			if(A_Index == 2){
				StringReplace, Data, Data, #, , All
				Arr_ShippingADD[A_Index] := Data
				continue
			}
			
			StringUpper, Data, Data
			Arr_ShippingADD[A_Index] := Trim(Data)
			;~ MsgBox % "Element number " . A_Index . " is " . Arr_ShippingADD[A_Index]
		}



		


		return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo, ShippingMethodStatus] ; 밑에 완성하고 나면 여러개의 배열을 리턴할테니. Arr_later은 나중에 더할 배열 





	} ; GettingInfoFromCurrentPage(PONumber) 메소드 끝
















; ############################################################################################################################################################################################################################################################
; ############################################################################################################################################################################################################################################################
; 아이템 업데이트 관련 메소드들
; ############################################################################################################################################################################################################################################################
; ############################################################################################################################################################################################################################################################


	
	
	; ######################################
	; 페이지에 아이템 업데이트 하는 메인 메소드
	; ######################################
	;~ UpdateItems(Array_StyleNo, Array_StylyColor, ShippingFee, InvoiceNo, TrackingNo, PONumber){
	UpdateItems(Array_StyleNo, Array_StylyColor, Array_StylyQty, ShippingFee, InvoiceNo, TrackingNo, PONumber){
		
		; Pre Order 아이템의 Available Date 저장하는 배열 선언
		Array_AvailableDate := object()
		
		; 정렬된 Available Date 저장하는 배열 선언
		Array_AvailableDate_Sorted := object()



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

		
		; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾은 뒤 해당 PO의 탭버튼 클릭해서 새 탭에서 열기
		; 예를들어 MTR1D39747D26 로 검색했으면 검색된 화면에는 MTR1D39747D26 뿐만 아니라 MTR1D39747D26-BO1 등도 같이 표시되기 때문에 딱 원하는 PO Number에 해당되는 탭버튼만 클릭하기 위해
		;~ if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
		if(driver.FindElementByXPath("//*[text() = '" PONumber "']").isDisplayed())
			driver.FindElementByXPath("//*[text() = '" PONumber "']//parent::td//child::i").click() ; 해당 PO의 탭버튼 클릭
				
		

		; 탭을 클릭해서 창을 새로 열었으니 현재 창을 다시 컨트롤 할 수 있도록 함수호출
		driver := ChromeGet()
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 2000

;		MsgBox, ready to update items
		
		
		

		; Order Status 상태를 읽어서 Block 되어있으면 메소드 나가기
		OrderStatus_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div/ul/li[2]/span/span[1]/select
		IsThisOrderPageBlocked := driver.FindElementByXPath(OrderStatus_Xpath).isEnabled()
		
		; Confirmed Orders : -1
		
		if(IsThisOrderPageBlocked == 0){
			MsgBox, IT'S BLOCKED ORDER PAGE
			
			; 전체 크롬 창이 아니라 새 탭으로 열린 이 탭만 닫기
			driver.close() ; closing just one tab of the browser
			Sleep 1000
			return
		}


		; Order Status 상태가 Partially Shipped 이거나 Fully Shipped 이면 메소드 나가기
		; Order Status 값이 3 이면 Partially Shipped 이고 4이면 Fully Shipped 이다.
		
		CurrentOrderStatus := driver.FindElementByXPath(OrderStatus_Xpath).Attribute("value") ; Order Status 상태 값을 얻기
		
;		MsgBox, % "CurrentOrderStatus : " . CurrentOrderStatus

		if(CurrentOrderStatus == 3){			
			MsgBox, IT'S FULLY SHIPPED OR PARTIALLY SHIPPED ORDERS
			
			; 전체 크롬 창이 아니라 새 탭으로 열린 이 탭만 닫기
			driver.close() ; closing just one tab of the browser
			Sleep 1000
			return			
		}
		else if(CurrentOrderStatus == 4){
			MsgBox, IT'S FULLY SHIPPED OR PARTIALLY SHIPPED ORDERS
			
			; 전체 크롬 창이 아니라 새 탭으로 열린 이 탭만 닫기
			driver.close() ; closing just one tab of the browser
			Sleep 1000
			return			
		}
		
		
		
		
		
		

		
		; 이 화면에 보낼 아이템이 있는지 찾기
		; 없으면 페이지 닫고 메소드 나가기
		; 아이템 찾는 동안 보여줄 프로그래스 바 
		TotalLoops = 57
		Gui, -Caption +AlwaysOnTop +LastFound
		Gui, Add, Text, x12 y9 w100 h20 , S E A R C H I N G . . .
		Gui, Add, Progress, w410 Range0-%TotalLoops% vProgress
		Gui, Show, w437 h84, SEARCHING ITEMS
				
		
		IsThereCheckBoxesOfItemsToUpdateOnThisPage = 0 ; 화면에 체크박스 있는지 확인하기 위해. 있다면 isDisplayed() 메소드에 의해 -1 값이 이 변수에 저장될 것이다
		
		Loop % Array_StyleNo.Maxindex(){

			StyleNo := Array_StyleNo[A_Index]
			StyleColor := Array_StylyColor[A_Index]
			StyleQty := Array_StylyQty[A_Index]

;			MsgBox % "Style No. " . A_Index . " is " . Array_StyleNo[A_Index]  "`n"  "Style Color" . A_Index . " is " . Array_StylyColor[A_Index] . "`n" . "Style Qty" . A_Index . " is " . Array_StylyQty[A_Index] ; 업데이트 할 스타일 번호와 색깔 확인하기 위해
		
			
			; StyleNo , StylyColor , StyleQty 셋 다 맞는 맞는 체크박스의 Xpath
			Xpath = //a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%StyleColor%')]//parent::td//following-sibling::td[contains(text(),'%StyleQty%')]//following-sibling::td//child::div[@class='check-square']
			IsThereCheckBoxesOfItemsToUpdateOnThisPage := driver.FindElementByXPath(Xpath).isDisplayed() ; 만약 해당하는 Xpath 위치에 있는 element 가 화면에 있다면 isDisplayed() 메소드에 의해 -1 값을 반환함
			
			; 프로그래스 바 5씩 증가
			GuiControl,,Progress, +5 ; 5씩 증가
		}
		
		Gui Destroy

;		MsgBox, % "IsThereCheckBoxesOfItemsToUpdateOnThisPage : " . IsThereCheckBoxesOfItemsToUpdateOnThisPage


		; 이 화면에 보낼 아이템이 하나도 없으면 탭 닫고 다음 Customer PO 열기 위해 현재 열린 탭 닫고 메소드 나가기
		; IsThereCheckBoxesOfItemsToUpdateOnThisPage 변수 값이 -1이 아니면 이 페이지에는 이번에 보낼 아이템이 없어서 업데이트 할 것이 없다는 뜻이므로
		if(IsThereCheckBoxesOfItemsToUpdateOnThisPage != -1){
			
			driver.close() ; closing just one tab of the browser
			Sleep 1000
			return
		}
		
;		MsgBox, pause, out of loop



		; 혹시 모르니 변수 다시 초기화 하기
		StyleNo := ""
		StyleColor := ""
		StyleQty := ""





		

	

		; ############################################################################################################
		; 여기서부터는 Order Status를 Confirmed Orders 로 바꾸는 동작. 실제 실행시에 주석처리 없애줘야 됨
		; ############################################################################################################

		
;~ /*
		; Order Status 상태를 읽은 후 Confirmed Orders 가 아니면 Confirmed Orders로 상태를 바꾼 후 Update 버튼 클릭
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div/ul/li[2]/span/span[1]/select
		CurrentOrderStatus := driver.FindElementByXPath(Xpath).Attribute("value") ; Order Status 상태 값을 얻기
		if(CurrentOrderStatus != 2){ ; Order Status 값이 2가 아니면 if문으로 들어가서 상태 바꾸기
			driver.FindElementByXPath(Xpath).SendKeys("Confirmed Orders") ; Order Status 상태를 Confirmed Orders 로 바꾸기
			
			;~ Sleep 500
			

			; Update 버튼 나타날때까지 기다린 뒤 클릭
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[1]			
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div/ul/li[2]/span/button[1]
			;~ while(driver.FindElementByXPath(Xpath).Attribute("value"))
				;~ Sleep 100
			;~ driver.FindElementByXPath(Xpath).click()
			
			loop{
				if(driver.FindElementByXPath(Xpath).isDisplayed()){
					driver.FindElementByXPath(Xpath).click()
					break
				}				
				Sleep 100
			}
			
			; Successfully updated order 메세지가 없어질때까지 기다리기
			; ComObjError(false) 명령어가 없으면 Successfully updated order 메세지가 없어진 뒤에 값을 찾을 수 없다고 에러메세지 나타남
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[1]/div/div[2]/fg-toast/div/div/div/span[1]
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[1]/div/div[2]/fg-toast/div/div			
			Sleep 500
			while(driver.FindElementByXPath(Xpath).Attribute("innerText"))
				Sleep 100

			MsgBox, Successfully updated order message has been vanished.
		}
		
		
		; Total 옆에 있는 전체 선택 체크박스 체크해주기
;		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/thead/tr/th[10]/div/label/div
;		driver.FindElementByXPath(Xpath).click()
;		Sleep 500

*/		
		
		
		; ##################################################################
		; 인스탁으로 나오는 아이템들 중에서 못 보내는 아이템들 백오더로 처리하기		
		; ##################################################################
		
		; 아이템에 pre order 날짜가 없어서 지금 보내야 되지만 어떤 이유에서든지 지금 보낼 수 없는 아이템들 백오더 처리하기
		; 그런 아이템들은 백오더 처리하려고 하면 언제 보낼건지 날짜를 물어보기 때문에 Domestic 아이템은 한 달 뒤로, Import 아이템들은 두 달 뒤로 날짜를 입력한 뒤 백오더 처리하기		
		; IsItSecondCallToProcessDomesticItems 변수는 지금 호출된 메소드가 Domestic 처리를 하기 위해 호출된것인지 Import 아이템들을 처리하기 위해서인지 결정하게 해준다
		; 메소드 처음 호출 땐 Import 아이템들을 처리한다
		; 그렇기 때문에 Domestic 오더 처리인지 표시하기 위한 IsItSecondCallToProcessDomesticItems 변수값을 거짓인 0값으로 넣고 메소드 호출
		; Domestic 아이템들 처리를 위한 두 번째 호출은 메소드 자체에서 스스로 재귀호출 한다
		; 그렇게 메소드 안에서 재귀호출 할 때는 IsItSecondCallToProcessDomesticItems 변수 값을 참 값인 1로 넣고 호출한다
	
		
		IsItSecondCallToProcessDomesticItems = 0
		FG.ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems, Array_StyleNo, Array_StylyColor, Array_StylyQty)



		

		
		

		; #################################################################
		; 남아있는 모든 아이들을 전체선택 해준 뒤 보낼 아이템만 체크박스 해제하기
		; #################################################################
		
		; 우선 아이템 번호와 색깔에 맞는 백오더 값의 위치를 확인 한 뒤 'Back Ordered' 라고 백오더 되어있지 않으면
		; 해당하는 아이템을 체크한 뒤 (실제로는 위에서 모든 아이템을 체크 해줬기 때문에 여기서 체크하면 체크가 취소되는 것)
		; 배열에서 해당 아이템의 값을 지우고 다시 루프의 첫번째로 돌아가기
		
		
		; Total 옆에 있는 전체 선택 체크박스 체크해주기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/thead/tr/th[10]/div/label/div
		driver.FindElementByXPath(Xpath).click()
		Sleep 500


		; 현재 화면에 있는 보낼 아이템들의 체크박스 체크하는 메소드 호출
		; 이 메소드가 백오더 처리 중에 호출되는 것인지 묻는 IsItCalledMiddleOfBOItemsProcessing 변수 값을 거짓 값인 0으로 설정해서 호출하기
		IsItCalledMiddleOfBOItemsProcessing = 0
		FG.CheckItemsInTheArrays(Array_StyleNo, Array_StylyColor, Array_StylyQty, IsItCalledMiddleOfBOItemsProcessing)
		





		; ################################################################
		; 위에서 체크 해제되지 않은 아이템들 백오더 버튼 눌러서 백오더로 넘기기
		; ################################################################
		
		; 아이템이 한개라도 체크 돼서 백오더로 넘겨줄 게 있으면
		; 백오더 버튼 클릭해서 해당 아이템들을 백오더로 넘긴 뒤
		; 공란에 필요한 정보들을 입력하고 Order Status 상태를 Fully Shipped 로 바꾼 뒤
		; Update & Notify Buyter 버튼 클릭해서 이 페이지에서 처리할 것들 완료하기


		if(CheckedItemExistsOrNot == 1){


MsgBox, Ready to Click BO Button			
			; 백오더 버튼 클릭해서 체크되지 않은(지금 다루고 있는 주문이 아닌) 아이템들 백오더로 넘기기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[1]/div[2]/button[3]
;			driver.FindElementByXPath(Xpath).click()


MsgBox, Shipping & Handling Fee
			; Shipping & Handling Fee 입력
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[3]/div/div[6]/div[2]/div/input
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
			driver.FindElementByXPath(Xpath).SendKeys(ShippingFee)

MsgBox, Invoice #
			; InvoiceNo 입력
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[2]/ul/li[1]/span/div/input
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div/ul/li[1]/span[5]/div/input
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
			driver.FindElementByXPath(Xpath).SendKeys(InvoiceNo)

MsgBox, Tracking #
			; TrackingNo 입력
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[2]/ul/li[2]/span/div/input
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[2]/ul/li[2]/span/div/input
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a") ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A
			driver.FindElementByXPath(Xpath).SendKeys(TrackingNo)

MsgBox, Ready to Click SAVE Button
			; Save 버튼 클릭
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[1]/div/div[3]/button
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[1]/div/div[3]/button
;			driver.FindElementByXPath(Xpath).click()


			; Order Status 바꾸기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select
			FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 4, "Fully Shipped Orders")
			Sleep 500
			
			; Update & Notify Buyter 버튼 클릭하기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[2]
			driver.FindElementByXPath(Xpath).click()
			Sleep 1000


MsgBox, Ready to Click Send Mail Button
			; Send Mail 버튼 클릭하기
			Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/fg-notify-buyer-modal/div/div[2]/div/div/div[2]/fieldset/div[5]/button[2]
;			driver.FindElementByXPath(Xpath).click()
		}




;		MsgBox, Close tab
		
		; 브라우저 닫기
		driver.close() ; closing just one tab of the browser
		
		
MsgBox, Finished Updating of this page
		
		
		
	} ; UpdateItems 메소드 끝
	
	








	
	
	; ##################################################################################################################
	; 현재 화면에서 Array_StyleNo, Array_StylyColor, Array_StylyQty 배열에 값이 있는, 즉 이번에 보낼 아이템 체크박스 클릭하기
	; IsItCalledMiddleOfBOItemsProcessing 값이 1이면 백오더 처리하는 메소드 중간에 호출됐다는 뜻
	; IsItCalledMiddleOfBOItemsProcessing 값이 0이면 백오더 처리하는 메소드는 이미 끝났고 실제로 아이템을 보내기 위한 과정에서 호출됐다는 뜻
	; ##################################################################################################################
	;~ CheckItemsInTheArrays(Array_StyleNo, Array_StylyColor, Array_StylyQty, IsItCalledMiddleOfBOItemsProcessing){
	CheckItemsInTheArrays(Array_StyleNo, Array_StylyColor, Array_StylyQty, IsItCalledMiddleOfBOItemsProcessing, Arr_Items_Having_NoPreOrderDates){
		
		driver := ChromeGet()
		
		Loop % Array_StyleNo.Maxindex(){
			Loop % Array_StylyColor.Maxindex(){

				StyleNo := Array_StyleNo[A_Index]
				StyleColor := Array_StylyColor[A_Index]
				StyleQty := Array_StylyQty[A_Index]
				
				MsgBox % "Style No. " . A_Index . " is " . Array_StyleNo[A_Index]  "`n"  "Style Color" . A_Index . " is " . Array_StylyColor[A_Index] . "`n" . "Style Qty" . A_Index . " is " . Array_StylyQty[A_Index] ; 업데이트 할 스타일 번호와 색깔 확인하기 위해


				;~ Xpath = //*[text() = '%StyleNo%']


				; 빈 배열이 아니라면 즉, 배열에 아이템이 한개라도 들어있다면 시작
				;~ if(StyleNo != ""){
				if(Array_StyleNo[A_Index] != ""){

					; 아래 코드들은 에러 메세지 표시 안해주는 함수 써야됨. 안 그러면 아래 Xpath들에 해당하지 않는 값을 찾으려고 하면 No Such Element 에러 나타냄
					; 만약에 현재 페이지에 찾는 StyleNo , StylyColor , StyleQty 셋 다 맞는 아이템이 있다면
					; 그 아이템에 해당하는 백오더 상태를 읽고 백오더가 아니라면 (근데 사실 셋 다 맞는 아이템이 있으면 백오더 값이 있을 수 없다. 왜냐면 백오더 된 아이템은 수량 값인 StyleQty 가 0이기 때문에 아예 if문이 실행이 안됨)
					; 그 아이템 체크 클릭해서 체크 해제하기
					Xpath = //*[text() = '%StyleNo%']//parent::div//parent::td//following-sibling::td//child::div[@class='text-s' and contains(text(),'%StyleColor%')]//parent::td//following-sibling::td[@class='align-mid text-right' and contains(text(),'%StyleQty%')]
					
					
					if(driver.FindElementByXPath(Xpath)){
					;~ if(driver.FindElementByXPath("//*[text() = '" StyleNo "']")){
					
						MsgBox, Item founded.
						
						BackOrderStatus := ""
						
						; StyleNo , StylyColor , StyleQty 셋 다 맞는 아이템의 백오더 값의 위치
						; StyleNo , StylyColor 두개의 값이 맞는 백오더 값의 위치. StyleQty까지 셋 다 맞는 것을 찾을 수 없는 이유는 백오더 된 아이템은 Qty가 0이기 때문
						Xpath2 = //a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%StyleColor%')]//following-sibling::div						
						;~ Xpath2 = //a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%StyleColor%')]//parent::td/following-sibling::td[contains(text(),'%StyleQty%')]//preceding-sibling::td//child::div[@class='order-table__state']
						BackOrderStatus := driver.FindElementByXPath(Xpath2).Attribute("innerText")
						
						MsgBox, % "Back Order is " . BackOrderStatus
							
						; 만약 백오더 값이 없다면 if문 실행해서 현재 배열값 초기화 시키고 체크박스 클릭하기
						if(!BackOrderStatus){


							; IsItCalledMiddleOfBOItemsProcessing 변수에 값이 있다면 백오더 처리중에 호출되었다는 뜻
							; 백오더 처리중에 호출되었다면 배열에서 아이템 정보를 삭제하면 안된다. 그렇게 되면 실제로 아이템을 보낼 때 보낼 정보가 없기 때문
							
							; IsItCalledMiddleOfBOItemsProcessing 변수에 값이 없다면 실제로 아이템을 내보내기 위한 동작중에 이 메소드가 호출되었다는 뜻
							; 실제로 아이템을 내보내기 위한 동작이라면 아이템에 해당하는 체크박스를 클릭한 뒤 클릭된 아이템의 정보를 Array_StyleNo, Array_StylyColor, Array_StylyQty 배열에서 삭제해야 된다. 그래야 다른 화면에서 또 클릭하지 않으니까
							; 아래 if문에서 삭제해준다

							if(!IsItCalledMiddleOfBOItemsProcessing){								
								
								MsgBox, It's for sending items
								
								; 체크박스 클릭한 아이템은 배열에서 지우기
								Array_StyleNo[A_Index] := ""
								Array_StylyColor[A_Index] := ""
								Array_StylyQty[A_Index] := ""

							}


							; StyleNo , StylyColor , StyleQty 셋 다 맞는 맞는 체크박스의 Xpath
							Xpath = //a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%StyleColor%')]//parent::td//parent::tr//child::td[@class='align-mid text-center']//child::div[@class='fg-checkbox']
							AllMatchedItem_Xpath = //a[contains(text(),'%StyleNo%')]//parent::div//parent::td//following-sibling::td//child::div[contains(text(),'%StyleColor%')]//parent::td//following-sibling::td[contains(text(),'%StyleQty%')]//following-sibling::td//child::div[@class='check-square']
							
							; 이거를 세개를 다 맞는게 아니라 StyleNo 만 맞는 체크박스를 찾아야 되나? 그럼 체크박스가 없는것도 체크되려나? 밑에서 계속 하나만 체크되는게 문제다
							

							
							; 이렇게 AlreadyUnCheckedItemColor 배열에 쓰레기값(?)이 없으면 아래 for index, element in AlreadyUnCheckedItemColor 가 실행되지 않는다.
							; StyleNo 값이 Arr_Items_Having_NoPreOrderDates 배열 안에 있다면 어차피 최소 한 번은 체크되어야 한다
							; 일단 아래의 NoColor라는 실제하지 않는 색깔값에 의해 for 문이 실행되서 체크박스가 클릭되면 AlreadyUnCheckedItemColor 배열의 첫번째 값은 클릭된 아이템이 갖고있는 색깔 값으로 바뀐다 그 이후로는 정상동작하게 된다
							AlreadyUnCheckedItemColor := array()
							AlreadyUnCheckedItemColor[1] := "NoColor" 

							; IsItCalledMiddleOfBOItemsProcessing 변수에 값이 있다면 백오더 처리중에 호출되었다는 뜻
							; 실제 아이템을 보내는 동작이 아닌 현재 페이지에는 pre order 날짜는 없어서 인스탁으로 나오지만 실제로는 물건이 없어서 지금 보내지 못하는 아이템들 백오더로 넘어줘야 하는 처리 실행
							if(IsItCalledMiddleOfBOItemsProcessing){
								

								; 실제로 보낼 아이템인데 체크 박스가 체크되어서 백오더로 넘어가게 됐다면 체크박스 클릭해서 체크된 것 해제하기
								; 체크박스가 체크되어 있지 않다면 그냥 넘어가기
								; driver.FindElementByXPath(StyleNoMatched_Xpath).isSelected() 통해서 체크박스가 체크되었는지 안 되었는지 확인한 뒤 처리하면 얼마나 좋겠냐만은 무엇때문인지 체크 유무에 상관없이 0값만 반환한다
								; 할 수 없이 아래의 방법을 이용한다
								
								; Arr_Items_Having_NoPreOrderDates 배열에는 이 메소드를 호출한 ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems 메소드를 통해 처리된, 현재 화면에서 pre order 날짜가 없는 Style No 값들 중에서 체크박스가 체크된 Style No 값이 있다
								; 그리고 StyleNo 변수에는 실제로 보낼 아이템의 Style No 값이 들어있다.
								; 만약 StyleNo 변수 안에 들어있는 값이 Arr_Items_Having_NoPreOrderDates 배열에도 있다면 그건 ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems 메소드를 통해서 이미 체크됐다는 뜻이므로 해당 체크박스를 클릭해서 체크를 해제해줘야된다
								; 값이 들어있지 않다면 클릭하지 않고 넘어간다. 값이 들어있는데 클릭한다면 체크를 해제하는 게 아니라 새롭게 체크를 하게 되는 것이므로 현재 페이지에서 보내야 될 아이템이 백오더 화면으로 넘어가게 된다.
								
								; Arr_Items_Having_NoPreOrderDates 배열에는 Style No 값만 있는 게 아니라 몇 번 체크됐는지도 들어있다. 무슨 뜻이냐면 만약 A1234가 두 번 클릭됐다면 A1234 값이 배열에 두 번 들어가 있다
								
;								#ofCheckBox = 1
								Loop % Arr_Items_Having_NoPreOrderDates.Maxindex()
								{
									
									; Style No 가 갖고있는 체크박스의 Xpath
									; [%#ofCheckBox%] 를 통해 #ofCheckBox 변수 값이 증가할 때마다 첫번째, 두번째, 세번째 ... 체크박스의 위치를 저장한다
;									StyleNoMatched_Xpath = (//a[contains(text(),'%StyleNo%')]//ancestor::tr//descendant::div[@class='check-square'])[%#ofCheckBox%]
									
									MsgBox % "Element number " . A_Index . " is " . Arr_Items_Having_NoPreOrderDates[A_Index]
									
									; [A_Index] 이렇게는 밑의 if 문에서 비교할 수 없으니 해당 배열 값을 StyleNoHasToBeUnCheckIfMatchedWith 변수에 넣는다
									StyleNoHasToBeUnCheckIfMatchedWith := Arr_Items_Having_NoPreOrderDates[A_Index]


/*
									; 체크박스가 없으면 루프 중단하고 나가기
									if(!driver.FindElementByXPath(StyleNoMatched_Xpath).isDisplayed()){
										MsgBox, % "No checkbox of #" . #ofCheckBox . " check box`n`nbreak from this checking boxes loop"
										#ofCheckBox = 1
										break
									}
*/									

									; 만약 StyleNo 값이 Arr_Items_Having_NoPreOrderDates 배열 안에 있다면
									if StyleNo in %StyleNoHasToBeUnCheckIfMatchedWith%
									{
										MsgBox, % "ITEM " . StyleNo . " HAS TO BE SHIPPED BUT IT HAS BEEN CHECKED SO TO RELEASE THE CHECKING" . "`n`n" . "실제로 보낼 아이템인데 체크가 됐기 때문에 체크를 해제합니다"
										

										; AlreadyUnCheckedItemColor 배열에는 색깔이 들어있는데 배열에 들어있는 값만큼(index 변수 크기만큼) 반복하면서 그 값을 element 변수에 넣는다
										for index, element in AlreadyUnCheckedItemColor
										{
											; 만약 고객에게 보내려는 StyleColor 가 element 변수랑 같지 않다면
											; pre order 날짜가 없어서 이미 체크됐지만 체크 해제하기 위해 체크박스를 클릭하지 않았다는 뜻이므로 in문 안으로 들어가서 클릭하기
											; 만약 StyleColor 와 element 가 같다면 이미 클릭해서 해제했다는 뜻이므로 또다시 클릭하면 안됨
											; in 문이 실행되면 체크박스 클릭해서 체크 해제해준뒤 그 색깔 값을 AlreadyUnCheckedItemColor 배열에 추가하기 이렇게 추가된 값은 다음 처리때 이미 클릭 해제됐는지 확인하기 위한 이번 동작에 사용됨
											if StyleColor not in %element%
											{
												; StyleNo , StylyColor , StyleQty 셋 다 맞는 맞는 체크박스 체크 해제하기
												driver.FindElementByXPath(AllMatchedItem_Xpath).click()
												AlreadyUnCheckedItemColor[A_Index] := StyleColor
											}
										}
																				
										
										
										
;										driver.FindElementByXPath(StyleNoMatched_Xpath).click()
;										break

										AlreadyUnCheckedStyleNo := StyleNoHasToBeUnCheckIfMatchedWith
									}
									
									; 다음 체크박스를 체크하기 위해 변수에 1 증가시키기
;									#ofCheckBox++

								
								}
								
								; 다른 Style No 의 색깔 값들을 저장하기 위해 배열 초기화하기
								AlreadyUnCheckedItemColor := []
								*/
;	MsgBox, out of the array
								
								; 브레이크 실행해서 실제 아이템을 보내기 위한 아래 동작들을 실행하지 않고 밖으로 나갑니다
								break
							}
	MsgBox



							; 백오더 처리중에 호출된 것이 아니라 실제로 아이템을 보내기 위한 동작이라면 
							; ; StyleNo , StylyColor , StyleQty 셋 다 맞는 맞는 체크박스를 클릭해서 체크를 해제해준다
							driver.FindElementByXPath(AllMatchedItem_Xpath).click()
							
							; 실제 삭제 동작은 위의 if 문을 통해 이미 처리됐음. 시각적으로 체크박스가 체크된 후 메세지 창을 나타내기 위함
							MsgBox % "Style No. " . A_Index . " is " . StyleNo  "`n"  "Style Color" . A_Index . " is " . StyleColor "`n" . "StyleQty" . A_Index . " is " . StyleQty . "`n`n"  "IS CHECKED THEN DELETED FROM THE ARRAY." ; 체크되고 지워진 배열값 확인하기 위해
							
							CheckedItemExistsOrNot = 1
						}
						else{
							MsgBox % "ITEM " . Array_StyleNo[A_Index] . " " . Array_StylyColor[A_Index] . " " . Array_StylyQty[A_Index] . " IS ON THIS PAGE, BUT IT'S ALREADY PROCESSED AS BACK ORDER, SO THIS APPLICATION IS GOING TO OPEN ANOTHER ORDER DETAILS THEN UPDATE THIS ITEM ON THAT"
						}
					}
					; 만약에 현재 페이지에 찾는 StyleNo , StylyColor , StyleQty 셋 다 맞는 아이템이 없다면 break로 끝내기
;					else
;						break
					
				}
				
			} ; Loop % Array_StylyColor.Maxindex() 루프 끝
			
		break ; 이걸 해줘야 Style No 와 Style Color 값이 서로 맞게 돌아간다

		} ; Loop % Array_StyleNo.Maxindex() 루프 끝

		return
		
	} ; CheckItemsInTheArrays 메소드 끝
	
	



































	
	; #####################################################################
	; 인스탁으로 나오는 아이템들 중에서 못 보내는 아이템들 백오더로 처리하기
	
	; 1) 체크 박스가 활성화 된 아이템들을 탐지(비활성화 된 아이템들은 Back Ordered 된 것)
	; 2) 그중에서 pre-order 날짜가 없는 아이템들을 탐지
	; 3) 이 메소드가 Import 아이템들을 위한 첫 번째 호출일 때
	;   3)-a 만약 Domestic 아이템이 탐지됐다면 현재 화면에 Domestic 아이템이 있다는 뜻으로 IsThereDomesticItem 변수에 표시만하고 Domestic 아이템들은 체크박스에 체크하지 않기
	;   3)-b Import 아이템들을 모두 체크해주기
	;   3)-c 보낼 아이템들을 체크 해제해 주기 위해서 CheckItemsInTheArrays 메소드 호출하기
	;   3)-d 백오더 버튼 클릭 후 나오는 달력에서 두 달 뒤의 15일을 선택하기
	;   3)-e Domestic 아이템들을 처리하기 위해 이 메소드 재귀호출하기
	;   
	; 4) 이 메소드가 Domestic 아이템들을 위한 두 번째 호출일 때
	;   4)-a 이미 첫 번째 호출에서 Import 아이템들이 처리됐기 때문에 pre-order 날짜가 없는 아이템은 모두 체크
	;   4)-b 보낼 아이템들을 체크 해제해 주기 위해서 CheckItemsInTheArrays 메소드 호출하기
	;   4)-c 백오더 버튼 클릭 후 나오는 달력에서 한 달 뒤의 15일을 선택하기	
	;
	; 5) 메소드 끝내기
	; IsItSecondCallToProcessDomesticItems 값이 0 이면 Import 아이템들 처리
	; IsItSecondCallToProcessDomesticItems 값이 1 이면 Domestic 아이템들 처리
	; #####################################################################
	
	;~ ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems){
	ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems, Array_StyleNo, Array_StylyColor, Array_StylyQty){
		
		#ofCheckBoxes := ""
		
		; pre order 날짜가 없는 Style No 저장하기 위한 배열선언
		; 이 배열은 CheckItemsInTheArrays 메소드로 넘겨져서 체크되지 않은 체크박스를 해제한답시고 다시 체크하지 않는데 사용된다
		Arr_Items_Having_NoPreOrderDates := object()
		
		
;~ /*		
		; 이 메소드가 첫 번째 호출인지 두 번째 호출인지 표시해주기 
		if(IsItSecondCallToProcessDomesticItems == 0){
			MsgBox, It's the first call for Import Items
		}
		else if(IsItSecondCallToProcessDomesticItems == 1){
			MsgBox, It's the second call for Domestic Items
		}
*/		
		
		; Domestic 아이템 몇 개 있는지 확인하는 변수
		IsThereDomesticItem = 0
				
		; 체크박스 몇 개 있는지 확인키 위한 배열 선언
		Str_#ofCheckBoxes := object()
				
		; 이 화면에 체크된 아이템이 있는지 표시하기 위해
		AreThereCheckedItemsOnThisScreen = 0
		


		driver := ChromeGet()


;~ MsgBox, % driver.Window.Title "`n" driver.Url



				
		; 현재 페이지의 HTML 소스 코드 읽기
		Xpath = //*
		HTML_Source := driver.FindElementByXPath(Xpath).Attribute("outerHTML")

		Sleep 700


			
		; 백오더 된 아이템은 체크박스가 활성화 되지 않는다
		; 현재 화면에 활성화 된 체크박스가 몇 개 있는지 알아보기 위해
		; 현재 화면에 있는 모든 활성화 된 체크박스 Str_#ofCheckBoxes 배열에 저장
		UnquotedOutputVar = imU)<div\s_ngcontent-c7=""\sclass="check-square"><div
		UnquotedOutputVar = class="check-square"><div

		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, UnquotedOutputVar, SubPat, FoundPos + strLen(SubPat)))
		{
;			MsgBox, % SubPat1
			Str_#ofCheckBoxes.Insert(SubPat1)
		}
			
		; 소스코드에서 읽을 때는 아이템에 있는 체크박스 갯수보다 2개가 더 많다. 아마도 Total 옆에 있는 체크박스 갯수포함 다른 것까지 세는 것 같다. 그래서 Str_#ofCheckBoxes 배열 갯수에서 2개를 빼준다
		#ofCheckBoxes := Str_#ofCheckBoxes.Maxindex() - 2
		
		MsgBox, % "# of CheckBoxes : " . #ofCheckBoxes
		
		; #ofCheckBoxes 변수에 값이 없으면 다시 재귀호출 하기
		if(#ofCheckBoxes == ""){
			
			MsgBox, NO #ofCheckBoxes ERROR. RECALL THIS METHOD AGAIN TO GET CORRECT # OF CHECKBOXES
			FG.ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems, Array_StyleNo, Array_StylyColor, Array_StylyQty)
		}
				
;		MsgBox, % "# of Check Boxes : " . #ofCheckBoxes
			
			
		; 체크박스 하나에 Style No 와 pre order 날짜값 2개가 있을 수 있기 때문에 체크박스 숫자만큼만 돌리면 안된다
		; 그렇다고 무한정 돌릴 수는 없고 Style No 와 pre order 날짜 값 갯수는 무조건 체크박스 * 2 갯수보다 작기 때문에 그만큼 루프를 돌린다
		; 루프는 Xpath 에 값이 없으면 아래의 if 문에서 빠져나감
		point_for_StyleNo_and_PreOrderDate = 1
		Loop, % #ofCheckBoxes * 2{
					
;			MsgBox, % "point_for_StyleNo_and_PreOrderDate : " . point_for_StyleNo_and_PreOrderDate
					
					
			; 활성화 된 체크박스(백오더 된 아이템은 체크박스가 활성화 되지 않음)의 Style No 와 Pre-order 날짜의 위치를 가리키는 Xpath. 
			; 둘의 값을 따로 나누지 못하는 이유는 Style No 와 Pre-order 날짜의 Xpath 값이 같기 때문이다
			; 다음, 그다음 값을 읽기 위해 point_for_StyleNo_and_PreOrderDate 변수 값을 증가시키면서 Xpath값을 전진시킨다
			Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody//child::div[@class='check-square']//ancestor::tbody//child::div[@class='order-table__no'])[%A_Index%]
			Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody//child::div[@class='check-square']//ancestor::tbody//child::div[@class='order-table__no'])[%point_for_StyleNo_and_PreOrderDate%]
					

					
					

			; Xpath 에 해당되는 Element(Style No 와 Pre-order 날짜값)가 있으면 if문 실행
			if(driver.FindElementByXPath(Xpath).isDisplayed()){
				
				; innerTextValueWhichIsSearched 변수에 Xpath 값 넣기
				innerTextValueWhichIsSearched := driver.FindElementByXPath(Xpath).Attribute("innerText")
						
;				MsgBox, % "innerTextValueWhichIsSearched : " . innerTextValueWhichIsSearched
						
						
				; 만약 innerTextValueWhichIsSearched 에 값이 없다면 pre-order 날짜가 없는 것이므로 체크해주기				
				if(!innerTextValueWhichIsSearched){
							
;					MsgBox, % "The value of this Xpath is null which means that it has to be checked"
							
					; 날짜가 없는 아이템의 색깔이 여러개일 수 있으니 해당되는 체크박스 계속 체크하다가 없으면 루프 빠져나가며 체크 중단하기
					j = 1
					; 같은 Style No 안에 체크박스가 많을 수 있으니 구분하기 위한 변수
					#ofCheckBox = 1
					Loop{
								
						; pre-order 날짜가 없는 해당 체크박스 체크해주기
						; 위에서 확인한 point_for_StyleNo_and_PreOrderDate 변수값이 가리키는(pro-order 값이 없는) Xpath로 일단 고정한 뒤
						; 그 값에서 파생되는 체크박스의 갯수는 #ofCheckBox 변수를 증가시키면서 확인하기
						XpathForCheck = ((/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody//child::div[@class='check-square']//ancestor::tbody//child::div[@class='order-table__no'])[%point_for_StyleNo_and_PreOrderDate%]//ancestor::tbody//child::div[@class='check-square'])[%#ofCheckBox%]
								
								
						; pre-order 날짜가 없어서 체크되어야 하는 체크박스의 Style No. 알아낸 뒤
						; 그 값 StyleNo_of_NoPreOrderDateCheckBox_Which_IsJustToBeChecked 변수에 넣기
						point_for_StyloNo_of_this_Xpath := point_for_StyleNo_and_PreOrderDate - 1 ; Style No 값은 직전 Xpath 에 있기 때문에 1을 빼주기
						pre_Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody//child::div[@class='check-square']//ancestor::tbody//child::div[@class='order-table__no'])[%point_for_StyloNo_of_this_Xpath%]
						StyleNo_of_NoPreOrderDateCheckBox_Which_IsJustToBeChecked := driver.FindElementByXPath(pre_Xpath).Attribute("innerText")
						MsgBox, % "Style No. of this Xpath is : " . StyleNo_of_NoPreOrderDateCheckBox_Which_IsJustToBeChecked

						
						; 지금 진행하는 상태가 이 메소드의 첫 번째 호출인지(Import처리) 두 번째 호출(Domestic처리)인지 구분해서 
						; Import 아이템 처리 위한 첫 번째 호출일때만 아래 코드 실행
						; 첫 번째 호출일때는 만약 현재 화면에 Domestic 아이템이 있다면 그 존재 여부만 IsThereDomesticItem 변수에 표시하고 넘어간 뒤
						; Import 아이템들만 체크해서 처리한다
						; 첫 번째 호출시 만약 현재 화면에 Domestic 아이템들이 있다는 것이 탐지되면 IsThereDomesticItem 값을 1로 바꿔서 현재 화면에 두 번째 호출에서 처리할 Domestic 아이템이 남아있다는 것을 표시한다.
						; 메소드 끝날때쯤의 if문에서 IsThereDomesticItem 값이 1이라면 Domestic 아이템들을 처리하기 위해 이 메소드를 다시 재귀호출 한다

						; IsItSecondCallToProcessDomesticItems 값이 0이면 이번이 첫 번째 호출이라는 뜻
						; 첫 번째 호출때는 Import 아이템들만 처리하기 때문에 pre order 날짜가 없는 아이템이라도 Domestic 아이템이면 break 통해서 체크하지 않고 그냥 나간다
						if(!IsItSecondCallToProcessDomesticItems){
							
							MsgBox, % "Because The value of variable 'IsItSecondCallToProcessDomesticItems' is : " . IsItSecondCallToProcessDomesticItems . ", It's The First Call of this method"
								
								
							; 현재 값이 Domestic 이면 IsThereDomesticItem 값에 Domestic 아이템이 있다는 것만 표시하고 해당 체크박스를 체크하지 말고 루프 끝내기
							; 맨 앞이 Domestic 코드일때만. 플러스 사이즈에는 아이템 번호와 관계 없이 마지막에 -P가 붙기 때문
							;~ if StyleNo_of_NoPreOrderDateCheckBox_Which_IsJustToBeChecked contains D,P,Y,LX,M,C
							if(RegExMatch(StyleNo_of_NoPreOrderDateCheckBox_Which_IsJustToBeChecked, "i)^(D|P|Y|LX|M|C)"))
							{
								MsgBox, it's a domestic item
										
								; Domestic 아이템 있다는 뜻으로 IsThereDomesticItem 값을 1로 해주기
								IsThereDomesticItem = 1
										
										
								; 체크박스 갯수 구분하기 위한 변수값 초기화 하기
								#ofCheckBox = 1
										
								; Style No 와 pre-order 날짜를 하나씩 읽기 위한 변수 증가하기
								point_for_StyleNo_and_PreOrderDate++
								
								; 아래 체크박스 체크하는 것 등을 실행하지 않고 루프 중단하고 나가기
								break
							}
						}
						


						; 체크박스가 없으면 루프 중단하고 나가기
						if(!driver.FindElementByXPath(XpathForCheck).isDisplayed()){
							MsgBox, % "No checkbox of #" . #ofCheckBox . " check box`n`nbreak from this checking boxes loop"
							#ofCheckBox = 1
							break
						}

						
						; 체크박스에 체크할 pre order 날짜가 없는 Style No 배열에 넣기
						; 이 배열은 CheckItemsInTheArrays 메소드 호출할 때 넘어가서 체크되지 않은 체크박스들을 체크한답시고 체크하는 것을 방지한다
						Arr_Items_Having_NoPreOrderDates.Insert(StyleNo_of_NoPreOrderDateCheckBox_Which_IsJustToBeChecked)
						

						; 체크박스가 있으면 체크박스 체크하기
						driver.FindElementByXPath(XpathForCheck).click()
								
						; 이 화면에 체크된 아이템이 있는지 표시하기 위해
						AreThereCheckedItemsOnThisScreen = 1
								
						Sleep 100
								
						MsgBox, % #ofCheckBox . " check box clicked"						
								
						; 다음 체크박스 체크하기 위해 변수값 1증가
						#ofCheckBox++
								
					}
				}
			}
			; Xpath 에 해당되는 Element 가 없으면 루프 중단하고 나가기
			else
				break
			
			; Style No 와 pre-order 날짜를 하나씩 읽기 위한 변수 증가하기
			point_for_StyleNo_and_PreOrderDate++
		}
		

		; 보낼 아이템들을 다시 체크해서 결과적으로 체크 해제하기
		; IsItCalledMiddleOfBOItemsProcessing 값을 참 값인 1로 설정해서 호출하기
		IsItCalledMiddleOfBOItemsProcessing = 1
		;~ FG.CheckItemsInTheArrays(Array_StyleNo, Array_StylyColor, Array_StylyQty, IsItCalledMiddleOfBOItemsProcessing)
		FG.CheckItemsInTheArrays(Array_StyleNo, Array_StylyColor, Array_StylyQty, IsItCalledMiddleOfBOItemsProcessing, Arr_Items_Having_NoPreOrderDates)
		
		
		driver := ChromeGet()



		; 체크된 아이템이 있으면 백오더 버튼 누르고 처리하기
		; 현재 화면에 pre-order 날짜가 없는 아이템이 하나도 없을수도 있으니까.
		if(AreThereCheckedItemsOnThisScreen){
			MsgBox, There are checked check boxes on the screen.
					
			; 백오더 버튼 클릭
			BackOrderButton_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[1]/div[2]/button[3]
			driver.FindElementByXPath(BackOrderButton_Xpath).click()
					
			Sleep 300

			; 날짜 입력창 뜨면 받는 날짜를 2달 후로 입력
			Backorder_Confirmation_Window_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/fg-backorder-modal/div/div[2]/div/div/div[2]					
			if(driver.FindElementByXPath(Backorder_Confirmation_Window_Xpath).isDisplayed()){
				MsgBox, DISPLAYED
				
				Sleep 300
						
				; 달력 화면 불러오기 위해서 날짜 입력칸을 클릭하기
				AvailableByBlank_Xpath = //input[@class='datepicker hasDatepicker']
				driver.FindElementByXPath(AvailableByBlank_Xpath).click()
						
				Sleep 300
						
				; 달력이 표시되면 오른쪽 화살표 클릭해서 한 달이나 두 달 뒤로 이동하기
				RightArrowOfTheCalendar_Xpath = //span[@class='ui-icon ui-icon-circle-triangle-e']
				if(driver.FindElementByXPath(RightArrowOfTheCalendar_Xpath).isDisplayed()){
					
					
					; Domestic 아이템이 있다고 표시됐으면 이번것은 Import 아이템 이므로 두 번 클릭해서 두 달 뒤로 만들기
					; 만약 IsThereDomesticItem 값이 1이 아니면 이번 것은 Domextic 아이템 이므로 한 번만 클릭해서 한 달 뒤로 만들기
					if(IsThereDomesticItem == 1){
						; 클릭 한 번 해서 한 달 뒤 만들기
						driver.FindElementByXPath(RightArrowOfTheCalendar_Xpath).click()
						Sleep 300						
					}
					; 클릭 한 번 해서 한 달 뒤 만들기
					driver.FindElementByXPath(RightArrowOfTheCalendar_Xpath).click()
							
					; 한 달 혹은 두 달 뒤의 달력에서 15일을 찾아서 클릭하기
					Date_15_Xpath = //a[contains(text(), '15')]
					Date_15_Xpath = //a[@class='ui-state-default' and contains(text(), '15')]					
					driver.FindElementByXPath(Date_15_Xpath).click()
							
							
					Sleep 5000

					; Confirm & NotifyBuyer_Xpath 버튼 클릭해서 백오더 완료하기
					Confirm_And_NotifyBuyer_Xpath = //button[@class='btn btn--min-width btn-blue margin-right-8']
					driver.FindElementByXPath(Confirm_And_NotifyBuyer_Xpath).click()
					
					Sleep 300


					; 현재창 Reload
					driver.refresh()
					
					driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
				
					; IsThereDomesticItem 값이 1이라면 이 메소드는 첫번째 호출된것이고 또한 현재 화면에 Domestic 아이템들이 탐지됐다는 뜻
					; 만약 지금의 호출이 두 번째 호출이라면 IsItSecondCallToProcessDomesticItems 값이 0이 아닌 1이 되어 IsThereDomesticItem 값이 1로 바뀌는 if문이 실행되지 않는다
					; 첫 번째 호출로 인해 Import 아이템들이 처리됐고 이제 남아있는 Domestic 아이템들을 처리할 차례
					; 만약 현재 화면에 Import 아이템들만 있다면 IsThereDomesticItem 값이 1로 바뀌지 않기 때문에 Domestic 아이템 처리를 위한 아래의 if문이 실행되지 않는다
					; 즉, 현재가 첫 번째 메소드 호출로써 위에서 Import 아이템들을 처리해줬고 IsThereDomesticItem 값이 1인것 때문에 현재 화면에 Domestic 아이템들이 있다는 두 가지 조건이 충족할 때만
					; 아래 if 문을 실행해서 남아있는 Domestic 아이템들을 처리한다 
					if(IsThereDomesticItem == 1){
					;~ if(IsThereDomesticItem){
						
						; IsItSecondCallToProcessDomesticItems 값을 1로 바꿔서 지금의 호출이 Domestic 아이템들을 위한 두 번째 호출임을 표시
						IsItSecondCallToProcessDomesticItems = 1
						
						; 남아있는 Domestic 아이템들을 처리하기 위한 두 번째 호출
						;~ FG.ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems)
						FG.ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems, Array_StyleNo, Array_StylyColor, Array_StylyQty)
						
					}					

				}

			}

		}


	return	
		
	} ; ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems 메소드 끝
	
	
	
















































; ############################################################################################################################################################################################################################################################
; ############################################################################################################################################################################################################################################################
; 짧고 자잘한 메소드들
; ############################################################################################################################################################################################################################################################
; ############################################################################################################################################################################################################################################################










	
	
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




/*		
		; Buyer Notes 내용 CustomerNoteOnWeb.txt 파일에 저장하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[1]
		FG.WringValueIntoAFile(Xpath, "CustomerNoteOnWeb")

		; Additional Info 내용 AdditionalInfo.txt 파일에 저장하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[2]
		FG.WringValueIntoAFile(Xpath, "AdditionalInfo")

		; Staff Notes 내용 StaffOnlyNote.txt 파일에 저장하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[2]/div/textarea
		FG.WringValueIntoAFile(Xpath, "StaffOnlyNote")
*/		

		
	}







	; ########################
	; 현재 상태를 Confirmed Orders로 바꾸기
	; ########################	
	ChangeOrderStatusToConfirmedOrders(){
		
		driver := ChromeGet()
		
	
		; Order Status 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select
		FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 2, "Confirmed Orders")
		
		; Update 버튼 클릭하기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[1]
		driver.FindElementByXPath(Xpath).click()		
		
		; Update & Notify Buyter 버튼 클릭하기
;		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[2]
;		driver.FindElementByXPath(Xpath).click()
				
		; Send Mail 버튼 클릭하기
;		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/fg-notify-buyer-modal/div/div[2]/div/div/div[2]/fieldset/div[5]/button[2]
;		driver.FindElementByXPath(Xpath).click()		
		
		return
	}








	; ########################
	; DropDown 메뉴 값을 읽은 후 argument 로 받은 값이 아니면 받은 값으로 바꾸기
	; StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)
	; ########################	
	compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, StatusToBePreferred, PreferredStatusValue){
		
		driver := ChromeGet()
		
		; Xpath 에 있는 element 가 활성화 되어 수정 가능한지 확인하기. 0이 반환되면 사용 불가
		IsItEnabled := driver.FindElementByXPath(Xpath).isEnabled()
		
		if(IsItEnabled == 0){
			MsgBox, It's Disabled.
		}
		
		; DropDown Manu 상태를 읽은 후 원하는 상태가 아니면 원하는 상태로 바꾸기
		CurrentStatus := driver.FindElementByXPath(Xpath).Attribute("value") ; 상태 값을 얻기
		
		;~ MsgBox, % CurrentStatus
		
		if(CurrentStatus != StatusToBePreferred){ ; 상태값이 원하는 값이 아니면 if문으로 들어가서 원하는 상태로 바꾸기
			driver.FindElementByXPath(Xpath).SendKeys(PreferredStatusValue)
			
			; 만약 상태값이 바뀌었으면 1을 리턴하기
			return 1
		}
		
		; 상태값이 바뀌지 않았으면 0을 리턴하기
		return 0
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
