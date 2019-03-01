#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include ChromeGet.ahk

;~ class CommWeb{
class CommWeb{
	
	; ########################
	; URL을 받아서 새로운 크롬 브라우저 창을 엽니다. (이 메소드로 열린 창은 재사용 될 수 있습니다)
	; ########################
	OpenNewBrowser(URL){
		run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)
		Sleep 3000
		
/*		
		driver := ChromeGet()
		
		; 아래 두 코드 동작되지 않음
		driver.AddArgument("--start-maximized") ; 창 최대화 하기
		driver.ExecuteScript("document.body.style.zoom = '100%';") ; 폰트 크기 100% 로 하기
		
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		Sleep 500
*/		
		return
	}
	
	
	

	; ########################
	; 로그인 합니다
	; ########################	
	Login(URL){
		
		driver := ChromeGet()
;		MsgBox, % driver.Window.Title "`n" driver.Url
		
		;~ sURL := driver.Url
		;~ MsgBox, % sURL
		
		; 로그인 페이지가 아니면 메소드 메세지 출력 후 메소드 빠져나가기
		if(!RegExMatch(driver.Url, "imU)login")){
			MsgBox, It's login page
			return
		}

		;~ MsgBox, % driver.Url


		; FG 로그인 페이지 처리
		if(RegExMatch(driver.Url, "imU)fashiongo")){
			
			;~ MsgBox, It's FG login page
			
			; ######
			; 천희
			; ######						
			Loginname = customer3
			Loginname = CHUNHEEHAN
			Password = Jo123456789
			Password = Jo123456789

			
			; ######
			; 루니
			; ######						
			;~ Loginname = customer2
			;~ Password = Jo123456789

			
			; ######
			; 디제이
			; ######						
			;~ Loginname = customer1
			;~ Password = Jo123456789


			; 아이디 입력
			Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[1]/input		
			Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[1]/input[2]
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Loginname) ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A 한 뒤 로그인 입력


			; 비밀번호 입력
			Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[2]/input
			Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[2]/input[2]
			;~ driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Password) ; Ctrl+A 한 뒤 비밀번호 입력
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Password).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 비밀번호 입력 후 엔터쳐서 로그인하기
			
			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
			Sleep 2000
			
			
			; 그냥 driver.Get(URL) 사용해서 원래 열려던 페이지를 열면 간단하고 좋을텐데 이상하게 이 코드가 timeout 에러를 발생시켜서 할 수 없이 로그인 한 지금 창을 닫은 뒤 다시 처음부터 시작하기
			driver.close() ; closing just one tab of the browser
			CommWeb.OpenNewBrowser(URL)
			
/*			
			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
			Sleep 2000
			
			MsgBox, % URL
			; 로그인 뒤 다시 해당 페이지를 열기(뉴오더면 뉴오더 페이지, 전체 검색이면 전체 오더 페이지)
			driver.Get(URL)
			
			MsgBox, "finish loading`n" URL
*/			
			return
		}
		

		; lashowroom 로그인 페이지 처리
		if(RegExMatch(driver.Url, "imU)lashowroom")){
			
			;~ MsgBox, It's LASHOWROOM login page
			
			; ######
			; 천희
			; ######			
			Loginname = jodifl3
			Password = jodifl
			SecurityCode = 7864
			CCPWD = FC83D28D
			
			; ######
			; 루니
			; ######					
			;~ Loginname = jodifl2
			;~ Password = jodifl
			;~ SecurityCode = 7864
			
			; ######
			; 디제이
			; ######		
			;~ Loginname = jodifl5
			;~ Password = jodifl
			;~ SecurityCode = 7864			
			
			

			; 아이디 입력
			Xpath = //*[@id="uname"]
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Loginname) ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A 한 뒤 로그인 입력


			; 비밀번호 입력
			Xpath = //*[@id="login_pwd"]
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Password).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 비밀번호 입력 후 엔터쳐서 로그인하기


			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
			Sleep 2000
			


			; Security Code 입력창인지 검사하기 위해 다시 현재 페이지 읽기
			driver := ChromeGet()
			
			; 얻은 현재 url이 Security Verification Code 입력 화면이면 보안코드 입력하기
			; https://admin.lashowroom.com/login_verify.php
			if(RegExMatch(driver.Url, "imU)login_verify")){
				
				; Security Code 입력
				Xpath = //*[@id="verification_code"]
				driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(SecurityCode).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 SecurityCode 입력 후 엔터쳐서 로그인하기				
				Sleep 500

				driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
				Sleep 2000
				
				; 로그인 뒤 다시 해당 페이지를 열기(뉴오더면 뉴오더 페이지, 전체 검색이면 전체 오더 페이지)
;				driver.Get(URL)
			}
			
			; 로그인 뒤 다시 해당 페이지를 열기(뉴오더면 뉴오더 페이지, 전체 검색이면 전체 오더 페이지)
;			driver.Get(URL)
			
			return
		}




		; eLAMBS 로그인 페이지 처리
		if(RegExMatch(driver.Url, "imU)elambs")){
			
			
			;~ MsgBox, 262144, Log In, PLEASE LOG IN eLAMBS
			;~ return
			
			
			
			;~ MsgBox, It's FG login page
			
			Loginname = admin
			Password = Jo2135057428.


			; 아이디 입력
			Xpath = //*[@id="tbUserName"]
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Loginname) ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A 한 뒤 로그인 입력


			; 비밀번호 입력
			Xpath = //*[@id="tbPassword"]
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Password).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 비밀번호 입력 후 엔터쳐서 로그인하기
			
			
			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
			
			return
		}
		






	}
	
}

