#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.





#Include CommWeb.ahk
#Include MethodsForAll.ahk
#Include ChromeGet.ahk


;~ #Include ChromeGet.ahk ; 이미 열려있는 창을 사용할 수 있게 해주는 함수. 
;~ global driver



class eLAMBS extends CommWeb{

	; ###########
	; eLAMBS 열어서 값 얻어오기
	; ###########
	Get_SOInfo_ofLAMBS(OrderID){
		
;		MsgBox, % "OrderID : " . OrderID
		
		URL = http://jodifl.elambs.com/page_Sales/SalesOrder_list.aspx
		CommWeb.OpenNewBrowser(URL)


		driver := ChromeGet()
		
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		
		; 현재 페이지가 LAMBS 로그인 페이지면 Login 메소드 실행해서 로그인 후 New Orders 메뉴로 들어가기
		if(RegExMatch(driver.Url, "imU)Login")){
			;~ MsgBox, It's FG login page
			CommWeb.Login()
			
			; 로그인 뒤 다시 해당 페이지를 열기(뉴오더면 뉴오더 페이지, 전체 검색이면 전체 오더 페이지)
			driver := ChromeGet()
			driver.Get(URL)			

			driver.executeScript("return document.readyState").toString().equals("complete")
			Sleep 1000
		}



		; Search Icon 클릭하기. 만약 없으면 그냥 넘기기
		Xpath = //*[@id="contents-header"]/div[4]
		if(driver.FindElementByXPath(Xpath))
			driver.FindElementByXPath(Xpath).click()
	
		driver.executeScript("return document.readyState").toString().equals("complete")
		Sleep 2000
		
		
		
		; Put Order Id in the blank
		Xpath = //*[@id="order-id-field"]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(OrderID).sendKeys(driver.Keys.ENTER)
		
		driver.executeScript("return document.readyState").toString().equals("complete")
		Sleep 1000




		; Order Id에 관련된 Customer 클릭하기
		Xpath = //a[contains(text(),'%OrderID%')]//parent::td//following-sibling::td//child::a[@class='edit-link context-menu']
		driver.FindElementByXPath(Xpath).click()		
		
		driver.executeScript("return document.readyState").toString().equals("complete")
		Sleep 1000




		; Billing Addr 얻기
		Xpath = //*[@id="ContentPlaceHolder1_BillingAddress_Address"]
		BillingAddrOneLAMBS := driver.FindElementByXPath(Xpath).Attribute("value")
		
		driver.executeScript("return document.readyState").toString().equals("complete")
		Sleep 1000
		






		; Shipping Addr 얻기
		Xpath = //*[@id="ContentPlaceHolder1_ShippingAddress_ShipAddress1"]
		ShippingAddrOneLAMBS := driver.FindElementByXPath(Xpath).Attribute("value")
		
		driver.executeScript("return document.readyState").toString().equals("complete")
		Sleep 1000


		
		
		; 이전 페이지로 돌아가기
		driver.executeScript("window.history.go(-1)")
		
		driver.executeScript("return document.readyState").toString().equals("complete")
		Sleep 1000
						
		
		
		; Order Id 클릭
		Xpath = //a[contains(text(),'%OrderID%')]
		driver.FindElementByXPath(Xpath).click()
		
		driver.executeScript("return document.readyState").toString().equals("complete")
		Sleep 1000
				
		
		
		
		
		

		

		;~ MsgBox, % ShippingAddrOneLAMBS

		
		
		
		;~ return ShippingAddrOneLAMBS
		return [BillingAddrOneLAMBS, ShippingAddrOneLAMBS]
	}


	; ###############
	; 배열에 있는 아이템들의 체크박스에 체크하기
	; ###############
	CheckCheckBoxOfItems(Str_Style){
		
		Style := Str_Style[4]
		Color := Str_Style[5]
		
;		MsgBox, % Style . "`n" . Color
		
		driver := ChromeGet()
		
		Xpath = //span[contains(text(),'%Style%')]//parent::td//following-sibling::td//child::span[contains(text(), '%Color%')] //parent::td//preceding-sibling::td//child::input
		driver.FindElementByXPath(Xpath).click()
		
		
		return
	}
	
	
	
	; ###############
	; 아이템 상태를 void로 바꾸고 메모창에 TRANSFERRED TO N41 적은 후 저장하고 끝내기
	; ###############
	WrapUpeLAMBS(){
		
		
		driver := ChromeGet()
		
		; Status 를 Void로 바꾸기
		Xpath = //*[@id="ContentPlaceHolder1_DetailsPOStatus"]
		driver.FindElementByXPath(Xpath).SendKeys(driver.Keys.ArrowDown).SendKeys(driver.Keys.ArrowDown).sendKeys(driver.Keys.ENTER)

;~ /*
		; 확인 버튼 누르기
		Xpath = //*[@id="contents-wrapper"]/div/div[10]/div[2]/div/div[1]/div[2]/span[5]/input
		driver.FindElementByXPath(Xpath).click()
		Sleep 300
*/

		; Sales Order Memo에 TRANSFERRED TO N41 적기
		Xpath = //*[@id="ContentPlaceHolder1_SalesOrderMemo"]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.ENTER).sendKeys(driver.Keys.ENTER).SendKeys("TRANSFERED TO N41")
		Sleep 300
		
		

		; 저장 버튼 클릭
		Xpath = //*[@id="save-button"]
		;~ driver.FindElementByXPath(Xpath).click()
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.ENTER)



		driver.close() ; closing just one tab of the browser
		
		return
	}









}