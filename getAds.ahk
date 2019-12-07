#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)






Loop{
  
    run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " )
    driver := ChromeGet()
    
    
    
    
	; 타겟 셀의 상태가 선택할 수 있게 될때까지 계속 반복
	Loop{	
		

		; Marketing Tools -> Fixed Price Ads -> Purchase Ads 로 이동		
		URL = https://vendoradmin.fashiongo.net/#/ads/purchase
		driver.Get(URL)

		
		
		; 지금 페이지가 로그인 페이지라면 로그인하기
		if(RegExMatch(driver.Url, "imU)login")){
			
			driver := FG_LoginOnly(driver)
			
			; Marketing Tools -> Fixed Price Ads -> Purchase Ads 로 이동 
			;~ driver.Get(URL)
		}

;~ /*
		; 광고 사는 페이지로 제대로 넘어갔는지 확인
		Loop{
			
			Sleep 100

			; 창으로 제대로 이동했으면 루프 나가기
			if(RegExMatch(driver.Url, "imU)purchase"))
				break
			else
				driver.Get(URL)	; Marketing Tools -> Fixed Price Ads -> Purchase Ads 로 이동 
		}
*/



		; Page 드롭다운 박스 상태를 Mobile Main Page 로 바꾸기
		page_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-purchase-ads/div[3]/div[1]/div[2]/div[2]/div[1]/div/select
;		changeStatusOfTheDropBoxUntilItsChanged(page_Xpath, "Mobile Main Page", 15, driver) ; Mobile Main Page 메뉴 상태 값이 15라서 제대로 바뀌었는지 상태 확인 하는 값으로 15를 넘겨줌		
;		changeStatusOfTheDropBoxUntilItsChanged(page_Xpath, "New Arrivals Page", 1, driver) ; New Arrivals Page
;       changeStatusOfTheDropBoxUntilItsChanged(page_Xpath, "Women Page", 5, driver) ; Women Page Page
		
		
        
        
        
        


		
		; Spot 드롭다운 박스 상태를 Mobile Main Small 로 바꾸기
		sPot_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-purchase-ads/div[3]/div[1]/div[2]/div[2]/div[2]/div/select
;		changeStatusOfTheDropBoxUntilItsChanged(sPot_Xpath, "Mobile Main Small", 71, driver) ; Mobile Main Small 메뉴 상태 값이 71이라 제대로 바뀌었는지 상태 확인 하는 값으로 71를 넘겨줌
;		changeStatusOfTheDropBoxUntilItsChanged(sPot_Xpath, "New Arrival Left", 7, driver) ; New Arrival Left
;		changeStatusOfTheDropBoxUntilItsChanged(sPot_Xpath, "New Arrival Right", 8, driver) ; New Arrival Right
;		changeStatusOfTheDropBoxUntilItsChanged(sPot_Xpath, "Top Right", 3, driver) ; Women Page -> Top Right		
		changeStatusOfTheDropBoxUntilItsChanged(sPot_Xpath, "New In Right", 8, driver) ; New In -> New In Right
        

		Sleep 1700	



		


		
		
		
;MsgBox, 262144, Title, 나 때문에 다른 사람도 RE-CAPCHA 경고 뜬 것 같음. 실험을 위해 다음에는 프로그램 사용 말고 그냥 손으로 해보자.`n`n먼저 로그인 해 둔 뒤 [[02:59:56]] 에 프로그램 시작하기. 5초가 살짝 안되게 4.X초 걸림. 미리 로그인 해두었으면 페이지 넘어가는데 5초 걸림.
		
		
		;~ Xpath = //*[@title='09/01/2019 (Sun)']
		Xpath = //*[@title='12/15/2019 (Sun)']		

		
        
		; 해당 Xpath 날짜 칸의 innerText 값을 innerText 변수에 넣기
		innerText := driver.FindElementByXPath(Xpath).Attribute("innerText")	;;; Available
		
		

		; innerText 변수 안에 Sold Out 이 있는지 확인 후 있으면 계속 루프 돌기
        needle_SoldOut = Sold Out
        if(RegExMatch(innerText, needle_SoldOut))
          continue
        
        ; innerText 변수 안에 Available 이 있는지 확인 후 있으면 선택할 수 있게 활성화 된 것이니 루프 빠져나감
		needle = Available
		if(RegExMatch(innerText, needle))
		{
			;MsgBox, 262144, Title, innerText : %innerText%에서 %needle% 찾았음. 루프 빠져나감
			break
		}



;MsgBox, 아직 활성화 안돼서 광고 살 수 없음		
		driver.refresh()
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		;~ Sleep 1000
		
		
	}
	



;MsgBox, 262144, Title, 이제 광고 구입할 수 있음
	
    
		; ** 광고 구입 할 수 있게 활성화 된 뒤 **        

/*
		dayAndDate0 = 09/01/2019 (Sun)
		dayAndDate1 = 09/02/2019 (Mon)
		dayAndDate2 = 09/03/2019 (Tue)
		dayAndDate3 = 09/04/2019 (Wed)
		dayAndDate4 = 09/05/2019 (Thu)
		dayAndDate5 = 09/06/2019 (Fri)
		dayAndDate6 = 09/07/2019 (Sat)
*/		
/*		
		dayAndDate0 = 10/20/2019 (Sun)
		dayAndDate1 = 10/21/2019 (Mon)
		dayAndDate2 = 10/22/2019 (Tue)
		dayAndDate3 = 10/23/2019 (Wed)
		dayAndDate4 = 10/24/2019 (Thu)
		dayAndDate5 = 10/25/2019 (Fri)
		dayAndDate6 = 10/26/2019 (Sat)
*/

        ; 화요일만 잡기
;       dayAndDate2 = 12/03/2019 (Tue)
; 		Xpath_temp = //*[@title='%dayAndDate2%']
;		driver.FindElementByXPath(Xpath_temp).click()
        
        ; 수요일만 잡기
;        dayAndDate3 = 11/06/2019 (Wed)
;  		Xpath_temp = //*[@title='%dayAndDate3%']
;		driver.FindElementByXPath(Xpath_temp).click()
        
        ; 목요일만 잡기
        dayAndDate4 = 12/12/2019 (Thu)
		Xpath_temp = //*[@title='%dayAndDate4%']
		driver.FindElementByXPath(Xpath_temp).click()
        
        
/*		
		; 여러 칸 누르기 위해 컨트롤 키 누름
		driver.sendKeys(driver.Keys.CONTROL)

		Xpath_temp = //*[@title='%dayAndDate0%']
		driver.FindElementByXPath(Xpath_temp).click()

		Xpath_temp = //*[@title='%dayAndDate1%']
		driver.FindElementByXPath(Xpath_temp).click()
		
		Xpath_temp = //*[@title='%dayAndDate2%']
		driver.FindElementByXPath(Xpath_temp).click()		
		
		Xpath_temp = //*[@title='%dayAndDate3%']
		driver.FindElementByXPath(Xpath_temp).click()
		
		Xpath_temp = //*[@title='%dayAndDate4%']
		driver.FindElementByXPath(Xpath_temp).click()

		Xpath_temp = //*[@title='%dayAndDate5%']
		driver.FindElementByXPath(Xpath_temp).click()
		
		Xpath_temp = //*[@title='%dayAndDate6%']
		driver.FindElementByXPath(Xpath_temp).click()		
		
		; 컨트롤 해제
		driver.sendKeys(driver.Keys.CONTROL)
*/		
        


/*		
; 클릭이 됐는지 확인하기 위해
outerHTML := driver.FindElementByXPath(Xpath).Attribute("outerHTML") ;;;; ui-selected
needle = selected
if(RegExMatch(outerHTML, needle))
{
	MsgBox, outerHTML : %outerHTML%에서 %needle% 찾았음. 해당 Xpath 칸이 클릭됐다는 뜻
}		
*/


;MsgBox, 262144, Title, 칸 다 선택했음




		; #### Purchase Ad 버튼 클릭하기 ####        
        ; 화면에서 버튼 그림 찾아 클릭하기. 못 찾으면 찾을때까지 계속 재귀호출로 반복, 즉 버튼 나올때까지 기다린다는 뜻
        findAndClick_thePurchaseAdButton()

/*
        ; Purchase Ad 버튼을 Xpath 를 사용하여 클릭하기
		Xpath_Purchase = //button[contains(text(), 'Purchase Ad')]
		driver.FindElementByXPath(Xpath_Purchase).click()		
		Sleep 700
*/		







;MsgBox, 262144, Title, 이제 CONFIRM 버튼 찾을 예정

		; #### Confirm 버튼 위치로 마우스 옮기기 ####
        ; 화면에서 버튼 그림 찾아 클릭하기. 못 찾으면 찾을때까지 계속 재귀호출로 반복, 즉 버튼 나올때까지 기다린다는 뜻
        findAndClick_theConfirmButton()
        
    
    
    
    
  
        MsgBox, 4100, Wintitle, 처음부터 다시 반복하시겠습니까?
        IfMsgBox, Yes
        {
          WinClose, ahk_exe chrome.exe
          continue
        }
        IfMsgBox, No
        {
          break
        }
  
  
}


MsgBox, 262144, Title, 프로그램을 종료합니다.


ExitApp























































    ;~ run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " )
    ;~ driver := ChromeGet()


	;~ driver:= ComObjCreate("Selenium.CHROMEDriver")	;Chrome driver
    ;~ driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	;~ driver.AddArgument("--start-maximized")	;창 최대화 하기 단, 처음 만들때만 적용됨 창을 연 뒤에는 적용 안됨





	; 타겟 셀의 상태가 선택할 수 있게 될때까지 계속 반복
	Loop{	
		

		; Marketing Tools -> Fixed Price Ads -> Purchase Ads 로 이동 
		URL = https://vendoradmin.fashiongo.net/#/ads/purchase
		driver.Get(URL)
		
		
		; 지금 페이지가 로그인 페이지라면 로그인하기
		if(RegExMatch(driver.Url, "imU)login")){
			
			driver := FG_LoginOnly(driver)
			
			; Marketing Tools -> Fixed Price Ads -> Purchase Ads 로 이동 
			;~ driver.Get(URL)
		}

;~ /*
		; 광고 사는 페이지로 제대로 넘어갔는지 확인
		Loop{
			
			Sleep 100

			; 창으로 제대로 이동했으면 루프 나가기
			if(RegExMatch(driver.Url, "imU)purchase"))
				break
			else
				driver.Get(URL)	; Marketing Tools -> Fixed Price Ads -> Purchase Ads 로 이동 
		}
*/



		; Page 드롭다운 박스 상태를 Mobile Main Page 로 바꾸기
		page_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-purchase-ads/div[3]/div[1]/div[2]/div[2]/div[1]/div/select
;		changeStatusOfTheDropBoxUntilItsChanged(page_Xpath, "Mobile Main Page", 15, driver) ; Mobile Main Page 메뉴 상태 값이 15라서 제대로 바뀌었는지 상태 확인 하는 값으로 15를 넘겨줌		
;		changeStatusOfTheDropBoxUntilItsChanged(page_Xpath, "New Arrivals Page", 1, driver) ; New Arrivals Page
        


		
		; Spot 드롭다운 박스 상태를 Mobile Main Small 로 바꾸기
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-purchase-ads/div[3]/div[1]/div[2]/div[2]/div[2]/div/select
;		changeStatusOfTheDropBoxUntilItsChanged(Xpath, "Mobile Main Small", 71, driver) ; Mobile Main Small 메뉴 상태 값이 71이라 제대로 바뀌었는지 상태 확인 하는 값으로 71를 넘겨줌
		changeStatusOfTheDropBoxUntilItsChanged(Xpath, "New Arrival Left", 7, driver) ; New Arrival Left
		Sleep 1000



		


		
		
		
;MsgBox, 262144, Title, pause
		
		
		;~ Xpath = //*[@title='09/01/2019 (Sun)']
		Xpath = //*[@title='09/08/2019 (Sun)']		

		
        
		; 해당 Xpath 날짜 칸의 innerText 값을 innerText 변수에 넣기
		innerText := driver.FindElementByXPath(Xpath).Attribute("innerText")	;;; Available
		
		

		; innerText 변수 안에 Sold Out 이 있는지 확인 후 있으면 계속 루프 돌기
        needle_SoldOut = Sold Out
        if(RegExMatch(innerText, needle_SoldOut))
          continue
        
        ; innerText 변수 안에 Available 이 있는지 확인 후 있으면 선택할 수 있게 활성화 된 것이니 루프 빠져나감
		needle = Available
		if(RegExMatch(innerText, needle))
		{
			;MsgBox, 262144, Title, innerText : %innerText%에서 %needle% 찾았음. 루프 빠져나감
			break
		}



;MsgBox, 아직 활성화 안돼서 광고 살 수 없음		
		driver.refresh()
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		;~ Sleep 1000
		
		
	}
	



;MsgBox, 262144, Title, 이제 광고 구입할 수 있음
	
    
		; ** 광고 구입 할 수 있게 활성화 된 뒤 **        

/*
		dayAndDate0 = 09/01/2019 (Sun)
		dayAndDate1 = 09/02/2019 (Mon)
		dayAndDate2 = 09/03/2019 (Tue)
		dayAndDate3 = 09/04/2019 (Wed)
		dayAndDate4 = 09/05/2019 (Thu)
		dayAndDate5 = 09/06/2019 (Fri)
		dayAndDate6 = 09/07/2019 (Sat)
*/		
;~ /*		
		dayAndDate0 = 09/08/2019 (Sun)
		dayAndDate1 = 09/09/2019 (Mon)
		dayAndDate2 = 09/10/2019 (Tue)
		dayAndDate3 = 09/11/2019 (Wed)
		dayAndDate4 = 09/12/2019 (Thu)
		dayAndDate5 = 09/13/2019 (Fri)
		dayAndDate6 = 09/14/2019 (Sat)
*/		
		
		; 여러 칸 누르기 위해 컨트롤 키 누름
		driver.sendKeys(driver.Keys.CONTROL)

		Xpath_temp = //*[@title='%dayAndDate0%']
		driver.FindElementByXPath(Xpath_temp).click()

		Xpath_temp = //*[@title='%dayAndDate1%']
		driver.FindElementByXPath(Xpath_temp).click()
		
		Xpath_temp = //*[@title='%dayAndDate2%']
		driver.FindElementByXPath(Xpath_temp).click()		
		
		Xpath_temp = //*[@title='%dayAndDate3%']
		driver.FindElementByXPath(Xpath_temp).click()
		
		Xpath_temp = //*[@title='%dayAndDate4%']
		driver.FindElementByXPath(Xpath_temp).click()

		Xpath_temp = //*[@title='%dayAndDate5%']
		driver.FindElementByXPath(Xpath_temp).click()
		
		Xpath_temp = //*[@title='%dayAndDate6%']
		driver.FindElementByXPath(Xpath_temp).click()		
		
		; 컨트롤 해제
		driver.sendKeys(driver.Keys.CONTROL)
		
        
/*		
; 클릭이 됐는지 확인하기 위해
outerHTML := driver.FindElementByXPath(Xpath).Attribute("outerHTML") ;;;; ui-selected
needle = selected
if(RegExMatch(outerHTML, needle))
{
	MsgBox, outerHTML : %outerHTML%에서 %needle% 찾았음. 해당 Xpath 칸이 클릭됐다는 뜻
}		
*/

;MsgBox, 칸 다 선택했음




		; #### Purchase Ad 버튼 클릭하기 ####        
        ; 화면에서 버튼 그림 찾아 클릭하기. 못 찾으면 찾을때까지 계속 재귀호출로 반복, 즉 버튼 나올때까지 기다린다는 뜻
        findAndClick_thePurchaseAdButton()





/*
        ; Purchase Ad 버튼을 Xpath 를 사용하여 클릭하기
		Xpath_Purchase = //button[contains(text(), 'Purchase Ad')]
		driver.FindElementByXPath(Xpath_Purchase).click()		
		Sleep 700
*/		



;MsgBox, 262144, Title, 이제 CONFIRM 버튼 찾을 예정





		; #### Confirm 버튼 위치로 마우스 옮기기 ####
        ; 화면에서 버튼 그림 찾아 클릭하기. 못 찾으면 찾을때까지 계속 재귀호출로 반복, 즉 버튼 나올때까지 기다린다는 뜻
        findAndClick_theConfirmButton()
        



MsgBox, 262144, Title, end


MsgBox, 262144, Title, 결제 버튼 누르기 전 일시정지

/*		
        ; Confirm 버튼을 Xpath 를 사용하여 클릭하기 (이거는 작동 안 할 가능성 높음 Purcha Ad 버튼을 눌러서 driver가 컨트롤을 잃었기 때문에. Confirm 클릭하려면 driver 를 새로 뜬 조그만 창으로 옮겨야 됨)
		Xpath_Confirm = //button[contains(text(), 'Confirm')]
		driver.FindElementByXPath(Xpath_Purchase).click()
*/			
	
	
	
	
MsgBox, 262144, Title, 끝
	
	
	
	
	
	
	
	ExitApp






Esc::
 Exitapp


















	; FG 로그인만 하는 메소드
	FG_LoginOnly(driver){


		; FG 로그인 페이지로 이동
		URL = https://vendoradmin.fashiongo.net/#/auth/login
		driver.Get(URL)


		Loginname = CHUNHEEHAN
		Password = Jodifl123.


		; 아이디 입력		
		Xpath = //input[@formcontrolname='username']
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Loginname) ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A 한 뒤 로그인 입력


		; 비밀번호 입력
		Xpath = //input[@formcontrolname='password']
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Password).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 비밀번호 입력 후 엔터쳐서 로그인하기


		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		return driver
	}






	; 드롭박스 메뉴를 원하는 상태로 바뀔때까지 계속해서 바꾸기
	changeStatusOfTheDropBoxUntilItsChanged(Xpath, StatusToBePreferred, statusWhichHasToBe, driver){		
		
		Loop{
		
		
;MsgBox 드롭박스 메뉴 ||  %StatusToBePreferred%  || 로 바꾸기
		
			; 드롭박스 메뉴를 원하는 상태로 바꾼 뒤 확인하는 의미로 탭키 누르기
			driver.FindElementByXPath(Xpath).SendKeys(StatusToBePreferred).sendKeys(driver.Keys.TAB)
			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
			
			; 드롭박스의 현재 상태 값을 얻기
			CurrentStatus := driver.FindElementByXPath(Xpath).Attribute("value")
			Sleep 100
			
;MsgBox, % "현재 드롭박스 상태 : " . CurrentStatus
			
			; 드롭박스 상태값이 원하는 것을 바뀌었으면 루프 빠져나가기
			if CurrentStatus contains %statusWhichHasToBe%
			;~ if(CurrentStatus == StatusToBePreferred)
				break
				
			Sleep 150
		}

		
		return
	}
	
	
	



        ; #### Purchase Ad 버튼 클릭하기 ####
        ; 화면에서 버튼 그림 찾아 클릭하기. 못 찾으면 찾을때까지 계속 재귀호출로 반복, 즉 버튼 나올때까지 기다린다는 뜻
        findAndClick_thePurchaseAdButton(){
          
          Text:="|<purchase Ad>*174$79.zzzzzDzzzzzzyE7zzzbzzzzzXz81zzznzzzzzlzYwb8MM63UkTky2SHY8430UM7sC1C9mMmNyHdnwbAUAtAzAk8A1wn6ECQaTaFb2Ty0n9zCH6H8mMDy0NYzU9k9Y14UDD42Ts4wAn0kMDbn1"
          if ok:=FindText(3087,628,150000,150000,0,0,Text)
          {
              CoordMode, Mouse
              X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
              MouseMove, X+W//2, Y+H//2
              Click
              Sleep 700
          }
          else
          {
              ; 못 찾았으면 재귀호출해서 계속 찾기
              Sleep 500
              findAndClick_thePurchaseAdButton()
          }
          
          return          
        }



        ; #### Confirm 버튼 위치로 마우스 옮기기 ####
        ; 화면에서 버튼 그림 찾아 클릭하기. 못 찾으면 찾을때까지 계속 재귀호출로 반복, 즉 버튼 나올때까지 기다린다는 뜻
        findAndClick_theConfirmButton(){
          
          Text:="|<confirm>*170$49.zzzzkTzzsDzzkDzzs3zzszzzstUk81004wU828V00TnYl4FaADtmMW9n67UtAF4tX1kQa8WQlY10H4FCMn1kNW8bAN"
          if ok:=FindText(3104,772,150000,150000,0,0,Text)
          {
              CoordMode, Mouse
              X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
              MouseMove, X+W//2, Y+H//2
              Click
              Sleep 500
          }
          else
          {
              ; 못 찾았으면 재귀호출해서 계속 찾기
              Sleep 500
              findAndClick_theConfirmButton()
          }
          
          return
          
        }





;===== Copy The Following Functions To Your Own Code Just once =====


; Note: parameters of the X,Y is the center of the coordinates,
; and the W,H is the offset distance to the center,
; So the search range is (X-W, Y-H)-->(X+W, Y+H).
; err1 is the character "0" fault-tolerant in percentage.
; err0 is the character "_" fault-tolerant in percentage.
; Text can be a lot of text to find, separated by "|".
; ruturn is a array, contains the X,Y,W,H,Comment results of Each Find.

FindText(x,y,w,h,err1,err0,text)
{
  xywh2xywh(x-w,y-h,2*w+1,2*h+1,x,y,w,h)
  if (w<1 or h<1)
    return, 0
  bch:=A_BatchLines
  SetBatchLines, -1
  ;--------------------------------------
  GetBitsFromScreen(x,y,w,h,Scan0,Stride,bits)
  ;--------------------------------------
  sx:=0, sy:=0, sw:=w, sh:=h, arr:=[]
  Loop, 2 {
  Loop, Parse, text, |
  {
    v:=A_LoopField
    IfNotInString, v, $, Continue
    Comment:="", e1:=err1, e0:=err0
    ; You Can Add Comment Text within The <>
    if RegExMatch(v,"<([^>]*)>",r)
      v:=StrReplace(v,r), Comment:=Trim(r1)
    ; You can Add two fault-tolerant in the [], separated by commas
    if RegExMatch(v,"\[([^\]]*)]",r)
    {
      v:=StrReplace(v,r), r1.=","
      StringSplit, r, r1, `,
      e1:=r1, e0:=r2
    }
    StringSplit, r, v, $
    color:=r1, v:=r2
    StringSplit, r, v, .
    w1:=r1, v:=base64tobit(r2), h1:=StrLen(v)//w1
    if (r0<2 or h1<1 or w1>sw or h1>sh or StrLen(v)!=w1*h1)
      Continue
    ;--------------------------------------------
    if InStr(color,"-")
    {
      r:=e1, e1:=e0, e0:=r, v:=StrReplace(v,"1","_")
      v:=StrReplace(StrReplace(v,"0","1"),"_","0")
    }
    mode:=InStr(color,"*") ? 1:0
    color:=RegExReplace(color,"[*\-]") . "@"
    StringSplit, r, color, @
    color:=Round(r1), n:=Round(r2,2)+(!r2)
    n:=Floor(255*3*(1-n)), k:=StrLen(v)*4
    VarSetCapacity(ss, sw*sh, Asc("0"))
    VarSetCapacity(s1, k, 0), VarSetCapacity(s0, k, 0)
    VarSetCapacity(rx, 8, 0), VarSetCapacity(ry, 8, 0)
    len1:=len0:=0, j:=sw-w1+1, i:=-j
    ListLines, Off
    Loop, Parse, v
    {
      i:=Mod(A_Index,w1)=1 ? i+j : i+1
      if A_LoopField
        NumPut(i, s1, 4*len1++, "int")
      else
        NumPut(i, s0, 4*len0++, "int")
    }
    ListLines, On
    e1:=Round(len1*e1), e0:=Round(len0*e0)
    ;--------------------------------------------
    if PicFind(mode,color,n,Scan0,Stride,sx,sy,sw,sh
      ,ss,s1,s0,len1,len0,e1,e0,w1,h1,rx,ry)
    {
      rx+=x, ry+=y
      arr.Push(rx,ry,w1,h1,Comment)
    }
  }
  if (arr.MaxIndex())
    Break
  if (A_Index=1 and err1=0 and err0=0)
    err1:=0.05, err0:=0.05
  else Break
  }
  SetBatchLines, %bch%
  return, arr.MaxIndex() ? arr:0
}

PicFind(mode, color, n, Scan0, Stride
  , sx, sy, sw, sh, ByRef ss, ByRef s1, ByRef s0
  , len1, len0, err1, err0, w, h, ByRef rx, ByRef ry)
{
  static MyFunc
  if !MyFunc
  {
    x32:="5589E583EC408B45200FAF45188B551CC1E20201D08945F"
    . "48B5524B80000000029D0C1E00289C28B451801D08945D8C74"
    . "5F000000000837D08000F85F00000008B450CC1E81025FF000"
    . "0008945D48B450CC1E80825FF0000008945D08B450C25FF000"
    . "0008945CCC745F800000000E9AC000000C745FC00000000E98"
    . "A0000008B45F483C00289C28B451401D00FB6000FB6C02B45D"
    . "48945EC8B45F483C00189C28B451401D00FB6000FB6C02B45D"
    . "08945E88B55F48B451401D00FB6000FB6C02B45CC8945E4837"
    . "DEC007903F75DEC837DE8007903F75DE8837DE4007903F75DE"
    . "48B55EC8B45E801C28B45E401D03B45107F0B8B55F08B452C0"
    . "1D0C600318345FC018345F4048345F0018B45FC3B45240F8C6"
    . "AFFFFFF8345F8018B45D80145F48B45F83B45280F8C48FFFFF"
    . "FE9A30000008B450C83C00169C0E803000089450CC745F8000"
    . "00000EB7FC745FC00000000EB648B45F483C00289C28B45140"
    . "1D00FB6000FB6C069D02B0100008B45F483C00189C18B45140"
    . "1C80FB6000FB6C069C04B0200008D0C028B55F48B451401D00"
    . "FB6000FB6C06BC07201C83B450C730B8B55F08B452C01D0C60"
    . "0318345FC018345F4048345F0018B45FC3B45247C948345F80"
    . "18B45D80145F48B45F83B45280F8C75FFFFFF8B45242B45488"
    . "3C0018945488B45282B454C83C00189454C8B453839453C0F4"
    . "D453C8945D8C745F800000000E9E3000000C745FC00000000E"
    . "9C70000008B45F80FAF452489C28B45FC01D08945F48B45408"
    . "945E08B45448945DCC745F000000000EB708B45F03B45387D2"
    . "E8B45F08D1485000000008B453001D08B108B45F401D089C28"
    . "B452C01D00FB6003C31740A836DE001837DE00078638B45F03"
    . "B453C7D2E8B45F08D1485000000008B453401D08B108B45F40"
    . "1D089C28B452C01D00FB6003C30740A836DDC01837DDC00783"
    . "08345F0018B45F03B45D87C888B551C8B45FC01C28B4550891"
    . "08B55208B45F801C28B45548910B801000000EB2990EB01908"
    . "345FC018B45FC3B45480F8C2DFFFFFF8345F8018B45F83B454"
    . "C0F8C11FFFFFFB800000000C9C25000"
    x64:="554889E54883EC40894D10895518448945204C894D288B4"
    . "5400FAF45308B5538C1E20201D08945F48B5548B8000000002"
    . "9D0C1E00289C28B453001D08945D8C745F000000000837D100"
    . "00F85000100008B4518C1E81025FF0000008945D48B4518C1E"
    . "80825FF0000008945D08B451825FF0000008945CCC745F8000"
    . "00000E9BC000000C745FC00000000E99A0000008B45F483C00"
    . "24863D0488B45284801D00FB6000FB6C02B45D48945EC8B45F"
    . "483C0014863D0488B45284801D00FB6000FB6C02B45D08945E"
    . "88B45F44863D0488B45284801D00FB6000FB6C02B45CC8945E"
    . "4837DEC007903F75DEC837DE8007903F75DE8837DE4007903F"
    . "75DE48B55EC8B45E801C28B45E401D03B45207F108B45F0486"
    . "3D0488B45584801D0C600318345FC018345F4048345F0018B4"
    . "5FC3B45480F8C5AFFFFFF8345F8018B45D80145F48B45F83B4"
    . "5500F8C38FFFFFFE9B60000008B451883C00169C0E80300008"
    . "94518C745F800000000E98F000000C745FC00000000EB748B4"
    . "5F483C0024863D0488B45284801D00FB6000FB6C069D02B010"
    . "0008B45F483C0014863C8488B45284801C80FB6000FB6C069C"
    . "04B0200008D0C028B45F44863D0488B45284801D00FB6000FB"
    . "6C06BC07201C83B451873108B45F04863D0488B45584801D0C"
    . "600318345FC018345F4048345F0018B45FC3B45487C848345F"
    . "8018B45D80145F48B45F83B45500F8C65FFFFFF8B45482B859"
    . "000000083C0018985900000008B45502B859800000083C0018"
    . "985980000008B45703945780F4D45788945D8C745F80000000"
    . "0E90B010000C745FC00000000E9EC0000008B45F80FAF45488"
    . "9C28B45FC01D08945F48B85800000008945E08B85880000008"
    . "945DCC745F000000000E9800000008B45F03B45707D368B45F"
    . "04898488D148500000000488B45604801D08B108B45F401D04"
    . "863D0488B45584801D00FB6003C31740A836DE001837DE0007"
    . "8778B45F03B45787D368B45F04898488D148500000000488B4"
    . "5684801D08B108B45F401D04863D0488B45584801D00FB6003"
    . "C30740A836DDC01837DDC00783C8345F0018B45F03B45D80F8"
    . "C74FFFFFF8B55388B45FC01C2488B85A000000089108B55408"
    . "B45F801C2488B85A80000008910B801000000EB2F90EB01908"
    . "345FC018B45FC3B85900000000F8C05FFFFFF8345F8018B45F"
    . "83B85980000000F8CE6FEFFFFB8000000004883C4405DC390"
    MCode(MyFunc, A_PtrSize=8 ? x64:x32)
  }
  return, DllCall(&MyFunc, "int",mode
    , "uint",color, "int",n, "ptr",Scan0, "int",Stride
    , "int",sx, "int",sy, "int",sw, "int",sh
    , "ptr",&ss, "ptr",&s1, "ptr",&s0
    , "int",len1, "int",len0, "int",err1, "int",err0
    , "int",w, "int",h, "int*",rx, "int*",ry)
}

xywh2xywh(x1,y1,w1,h1,ByRef x,ByRef y,ByRef w,ByRef h)
{
  SysGet, zx, 76
  SysGet, zy, 77
  SysGet, zw, 78
  SysGet, zh, 79
  left:=x1, right:=x1+w1-1, up:=y1, down:=y1+h1-1
  left:=left<zx ? zx:left, right:=right>zx+zw-1 ? zx+zw-1:right
  up:=up<zy ? zy:up, down:=down>zy+zh-1 ? zy+zh-1:down
  x:=left, y:=up, w:=right-left+1, h:=down-up+1
}

GetBitsFromScreen(x,y,w,h,ByRef Scan0,ByRef Stride,ByRef bits)
{
  VarSetCapacity(bits,w*h*4,0), bpp:=32
  Scan0:=&bits, Stride:=((w*bpp+31)//32)*4
  Ptr:=A_PtrSize ? "UPtr" : "UInt", PtrP:=Ptr . "*"
  win:=DllCall("GetDesktopWindow", Ptr)
  hDC:=DllCall("GetWindowDC", Ptr,win, Ptr)
  mDC:=DllCall("CreateCompatibleDC", Ptr,hDC, Ptr)
  ;-------------------------
  VarSetCapacity(bi, 40, 0), NumPut(40, bi, 0, "int")
  NumPut(w, bi, 4, "int"), NumPut(-h, bi, 8, "int")
  NumPut(1, bi, 12, "short"), NumPut(bpp, bi, 14, "short")
  ;-------------------------
  if hBM:=DllCall("CreateDIBSection", Ptr,mDC, Ptr,&bi
    , "int",0, PtrP,ppvBits, Ptr,0, "int",0, Ptr)
  {
    oBM:=DllCall("SelectObject", Ptr,mDC, Ptr,hBM, Ptr)
    DllCall("BitBlt", Ptr,mDC, "int",0, "int",0, "int",w, "int",h
      , Ptr,hDC, "int",x, "int",y, "uint",0x00CC0020|0x40000000)
    DllCall("RtlMoveMemory","ptr",Scan0,"ptr",ppvBits,"ptr",Stride*h)
    DllCall("SelectObject", Ptr,mDC, Ptr,oBM)
    DllCall("DeleteObject", Ptr,hBM)
  }
  DllCall("DeleteDC", Ptr,mDC)
  DllCall("ReleaseDC", Ptr,win, Ptr,hDC)
}

base64tobit(s)
{
  ListLines, Off
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  StringCaseSense, On
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:=(i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,A_LoopField,v)
  }
  StringCaseSense, Off
  s:=SubStr(s,1,InStr(s,"1",0,0)-1)
  s:=RegExReplace(s,"[^01]+")
  ListLines, On
  return, s
}

bit2base64(s)
{
  ListLines, Off
  s:=RegExReplace(s,"[^01]+")
  s.=SubStr("100000",1,6-Mod(StrLen(s),6))
  s:=RegExReplace(s,".{6}","|$0")
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:="|" . (i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,v,A_LoopField)
  }
  ListLines, On
  return, s
}

ASCII(s)
{
  if RegExMatch(s,"(\d+)\.([\w+/]{3,})",r)
  {
    s:=RegExReplace(base64tobit(r2),".{" r1 "}","$0`n")
    s:=StrReplace(StrReplace(s,"0","_"),"1","0")
  }
  else s=
  return, s
}

MCode(ByRef code, hex)
{
  ListLines, Off
  bch:=A_BatchLines
  SetBatchLines, -1
  VarSetCapacity(code, StrLen(hex)//2)
  Loop, % StrLen(hex)//2
    NumPut("0x" . SubStr(hex,2*A_Index-1,2), code, A_Index-1, "char")
  Ptr:=A_PtrSize ? "UPtr" : "UInt"
  DllCall("VirtualProtect", Ptr,&code, Ptr
    ,VarSetCapacity(code), "uint",0x40, Ptr . "*",0)
  SetBatchLines, %bch%
  ListLines, On
}

; You can put the text library at the beginning of the script,
; and Use Pic(Text,1) to add the text library to Pic()'s Lib,
; Use Pic("comment1|comment2|...") to get text images from Lib
Pic(comments, add_to_Lib=0) {
  static Lib:=[]
  if (add_to_Lib)
  {
    re:="<([^>]*)>[^$]+\$\d+\.[\w+/]{3,}"
    Loop, Parse, comments, |
      if RegExMatch(A_LoopField,re,r)
        Lib[Trim(r1)]:=r
  }
  else
  {
    text:=""
    Loop, Parse, comments, |
      text.="|" . Lib[Trim(A_LoopField)]
    return, text
  }
}


/***** C source code of machine code *****

int __attribute__((__stdcall__)) PicFind(int mode
  , unsigned int c, int n, unsigned char * Bmp
  , int Stride, int sx, int sy, int sw, int sh
  , char * ss, int * s1, int * s0
  , int len1, int len0, int err1, int err0
  , int w, int h, int * rx, int * ry)
{
  int x, y, o=sy*Stride+sx*4, j=Stride-4*sw, i=0;
  int r, g, b, rr, gg, bb, e1, e0;
  if (mode==0)  // Color Mode
  {
    rr=(c>>16)&0xFF; gg=(c>>8)&0xFF; bb=c&0xFF;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
      {
        r=Bmp[2+o]-rr; g=Bmp[1+o]-gg; b=Bmp[o]-bb;
        if (r<0) r=-r; if (g<0) g=-g; if (b<0) b=-b;
        if (r+g+b<=n) ss[i]='1';
      }
  }
  else  // Gray Threshold Mode
  {
    c=(c+1)*1000;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
        if (Bmp[2+o]*299+Bmp[1+o]*587+Bmp[o]*114<c)
          ss[i]='1';
  }
  w=sw-w+1; h=sh-h+1;
  j=len1>len0 ? len1 : len0;
  for (y=0; y<h; y++)
  {
    for (x=0; x<w; x++)
    {
      o=y*sw+x; e1=err1; e0=err0;
      for (i=0; i<j; i++)
      {
        if (i<len1 && ss[o+s1[i]]!='1' && (--e1)<0)
          goto NoMatch;
        if (i<len0 && ss[o+s0[i]]!='0' && (--e0)<0)
          goto NoMatch;
      }
      rx[0]=sx+x; ry[0]=sy+y;
      return 1;
      NoMatch:
      continue;
    }
  }
  return 0;
}

*/


;================= The End =================

;










