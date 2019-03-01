#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk

; 오토핫키 urldownloadtofile

;해외 배송인지 확인하는 것 먼저

; CONSOLIDATION으로 들어온 주문을 배송하려 했더니 경고창 나왔음 (ERROR MESSAGE 폴더에 저장했음)

;21. Gui를 여러개 만들기 - Gui
; http://jogamja.tistory.com/m/21


	GetInfoFromFashiongo(PO_F, FGServer){
		
		; 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기
		WinClose, http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%PO_F% - Internet Explorer

		Clipboard :=
		
		
		Loginname = customer3
		Password = Jo123456789
		URL = http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%PO_F%

		WB := ComObjCreate("InternetExplorer.Application")
		WB.Visible := true
		WB.Navigate(URL)
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		   Sleep, 10



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
			   Sleep, 10
/*		   
		   ; 주소 등 정보 얻기 위해 PO 창 다시 열기
	   		WB.Navigate(URL)
			While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
				Sleep, 10
*/
		   
		;	MsgBox, found login
		}





		; 신용카드 번호 알아내기 위한 동작
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; 이 부분이 실행되면 add1등 읽어 온 모든 html source 값이 초기화 됨. 이유는 모르겠음 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		wb.document.getElementsByTagName("SPAN")[15].Click() ; 신용카드 번호 하이퍼링크 클릭
		wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderGeneralDetail_tbCCAccessPassword").value := "Jo123456789"  ;PASSWORD 창에 비밀번호 입력
		wb.document.getElementsByTagName("A")[23].Click() ; Submit 버튼 누르기
		

		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 10

		; FASHIONGO의 CC정보 읽어오기
		; 아래 두 개 모두 작동함
		Sleep 1000
		;CCinfoOnFASHIONGO := wb.document.getElementsByTagName("SPAN")[16].innerText
		CCinfoOnFASHIONGOVal := wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderGeneralDetail_lblCreditCardDetail").innerText
		
		
		; CCinfoOnFASHIONGOVal 문자열 왼쪽의 공란과 특수문자 제거하기
		; 공란만 제거했으면 " " 가 되고 특수문자만 제거하려면 "`t"인데 두 개를 제거하려고 " `t" 사용
		CCinfoOnFASHIONGOVal := LTrim(CCinfoOnFASHIONGOVal, " `t`r")
		
		
		; 혹시 cc 값을 못 얻었으면 다시 시도해서 얻기
		if(!CCinfoOnFASHIONGOVal)
			CCinfoOnFASHIONGOVal := wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderGeneralDetail_lblCreditCardDetail").innerText
			
		
		;MsgBox, % CCinfoOnFASHIONGOVal


		;CCinfoOnFASHIONGO := wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderGeneralDetail_lblCreditCardDetail").innerText
		
		; CCinfoOnFASHIONGO 안에 있는 신용카드정보 CCinfoOnFASHIONGO.txt 파일에 저장하기
		FileAppend, %CCinfoOnFASHIONGOVal%, %A_ScriptDir%\CreatedFiles\CCinfo.txt
		











		; html source 얻기
		htmlSourcecode := WB.Document.All[0].outerhtml
		;htmlSourcecode := WB.Document.All[0].Value
		
		; html source 얻기
		;htmlSourcecode := wb.document.all.primary_nav.innerText
		
		; html source 얻기
		;htmlSourcecode := wb.document.documentElement.innerText


		; 패션고 페이지에서 고객정보 읽어오기
		FindInfoInFASHIONGO(htmlSourcecode)



;		MsgBox, CustomerNoteOnWeb is %CustomerNoteOnWeb%
		


		; Customer notes 읽어오기
		;CustomerNoteOnWeb := wb.document.getElementsById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderItemDetail_tbCustomerNote").value
		CustomerNoteOnWebVal := wb.document.getElementsByTagName("TEXTAREA")[2].innerText
		StringUpper, CustomerNoteOnWebVal, CustomerNoteOnWebVal ; 고객 메모 대문자로 바꾸기
		
		; CustomerNoteOnWebVal 안에 있는 고객 메모 CustomerNoteOnWebVal.txt 파일에 저장하기
		FileAppend, %CustomerNoteOnWebVal%, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt
		



		
		
		; Staff only notes 읽어오기
		;StaffOnlyNoteVal := wb.document.getElementsById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderItemDetail_tbStaffOnlyNoteVal").value
		StaffOnlyNoteVal := wb.document.getElementsByTagName("TEXTAREA")[3].innerText
		StringUpper, StaffOnlyNoteVal, StaffOnlyNoteVal ; Staff only notes 대문자로 바꾸기
		
		; StaffOnlyNoteVal 안에 있는 직원용 내부 메모  StaffOnlyNoteVal.txt 파일에 저장하기		
		FileAppend, %StaffOnlyNoteVal%, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt


		
		
		
		;MsgBox, Customer's memo is : %Cmemo%

		

		
		


		

		






/*
		Clipboard :=
		
		; 백 오더에는 po번호에 -BO가 붙는다 BO BO1 BO2 이런 식으로
		URL = http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%PO_F%
		
		; 크롬 새창으로 열기
		run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL
		
		
;		WinActivate, %URL% - Google Chrome

		Sleep 2000
		
		; 소스코드 창 열기 Ctrl + u
		SendInput, ^u
		Sleep 2000

		; 소스코드 전체선택 후 복사
		SendInput, ^a^c
		Sleep 1000
		
		HTMLSource := % Clipboard
		Sleep 1000

		; 소스코드 탭 닫기
		SendInput, ^w
		
		;HTMLSource = orderGeneralDetail_lblShipCompanyName">Goose Feathers</span></td>


		; 패션고 페이지에서 정보 읽어오기
		FindInfoInFASHIONGO(HTMLSource)
		
		
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
		DeleteAdd2InAdd1(Address1)		
*/








/*		
	; 패션고 정보가 LAMBS와 맞는지 묻고 Yes 눌렀으면 LAMBS에서 정보 얻는 GetInfoFromLAMBS() 함수 호출
	;MsgBox, 4, UPS Label Print out, 고객 메모 확인`n`nIs Info on FASHIONGO same as LAMBS?
	MsgBox, 4100, UPS Label Print out, 고객 메모 확인`n`nIs Info on FASHIONGO same as LAMBS?
	IfMsgBox, Yes
	{
		GetInfoFromLAMBS()		
	}
	else IfMsgBox, No
	{
		Reload		
	}		
*/


	return
	}
	
	
	
asdfasdfasdfasdf(PO_F, FGServer){
	
	
	
		Clipboard :=
		
		
		Loginname = customer3
		Password = Jo123456789
		URL = http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%PO_F%

		WB := ComObjCreate("InternetExplorer.Application")
		WB.Visible := true
		WB.Navigate(URL)
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		   Sleep, 10



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
			   Sleep, 10
		   
		;	MsgBox, found login
		}	
	
	
	wb.document.getElementsByTagName("SPAN")[15].Click() ; 신용카드 번호 하이퍼링크 클릭
	wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderGeneralDetail_tbCCAccessPassword").value := "Jo123456789"  ;PASSWORD 창에 비밀번호 입력
	wb.document.getElementsByTagName("A")[23].Click() ; Submit 버튼 누르기
	

	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 10


	
Sleep 1000

;CCinfoOnFASHIONGO := wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderGeneralDetail_lblCreditCardDetail").value
;CCinfoOnFASHIONGO := wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderGeneralDetail_lblCreditCardDetail").innerText
;CCinfoOnFASHIONGO := wb.document.getElementsByTagName("SPAN")[5].value

;1324ZXCVQWER := wb.document.getElementsByTagName("SPAN")[5].value ;:= 76035
CCinfoOnFASHIONGO := wb.document.getElementsByTagName("SPAN")[16].innerText ;:= 76035
;wb.document.getElementsByTagName("SPAN")[5].value := 76035

MsgBox, % CCinfoOnFASHIONGO
MsgBox, % 1324ZXCVQWER

	
	
	
	
	
	
	
	
	return
}	