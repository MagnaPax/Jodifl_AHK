#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk

;해외 배송인지 확인하는 것 먼저





GetInfoFromLASHOROOM(PO_L){
	
		; 변수 PO_L의 7번째 부터 PO_L의 길이만큼(PO_L길이 끝까지) 읽어오는 함수
		PartiallyPO := substr(PO_L,7,StrLen(PO_L))
	

		URL = https://admin.lashowroom.com/order_edit_v1.php?order_id=%PartiallyPO%&list_option=new
		
		
		
		
		
		
		
		

		Clipboard :=
		
		
		Loginname = jodifl
		Password = j123456789
		SecurityCode = 7864
		

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
			wb.document.getElementById("uname").value := Loginname  ;ID 입력
			wb.document.getElementById("login_pwd").value := Password ; 비밀번호 입력
			wb.document.getElementsByTagName("INPUT")[2].Click() ; 로그인 버튼 누르기

			While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			   Sleep, 10
		   
		;	MsgBox, found login
		
		
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



			; 얻은 현재 url이 Security Verification Code 입력 화면이면 보안코드 입력하기
			; https://admin.lashowroom.com/login_verify.php
			if(RegExMatch(sURL, "imU)login_verify")){
				wb.document.getElementById("verification_code").value := SecurityCode  ;Security Code 입력			
				wb.document.getElementsByTagName("INPUT")[2].Click() ; 로그인 버튼 누르기

				While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
				   Sleep, 10	
				   
		
			}


		   ; 주소 등 정보 얻기 위해 PO 창 다시 열기
			WB.Navigate(URL)
			While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
				Sleep, 10

		}
		
		

				
		
		
		
		
		
		
		
		
		; html source 얻기
;		htmlSourcecode := WB.Document.All[0].outerhtml
		
		; html source 얻기
		;htmlSourcecode := wb.document.all.primary_nav.innerText
		
		; html source 얻기
		;htmlSourcecode := wb.document.documentElement.innerText


		
		
		;MsgBox, Customer's memo is : %Cmemo%

		
		; LA쇼룸 페이지에서 정보 읽어오기
;		FindInfoInLASHOWROOM(htmlSourcecode)



		; CompanyName
		CompanyName := wb.document.getElementsByTagName("TD")[36].innerText
		;StringUpper, CompanyName, CompanyName ; CompanyName 대문자로 바꾸기 회사 이름은 대소문자 섞인 게 나은 것 같음
		
		; Address11
		Address1 := wb.document.getElementsByTagName("TD")[38].innerText
		
		; Address2
		Address2 := wb.document.getElementsByTagName("TD")[40].innerText
		
		; City
		City := wb.document.getElementsByTagName("TD")[42].innerText
		
		; State
		State := wb.document.getElementsByTagName("TD")[44].innerText
		
		; ZipCode
		ZipCode := wb.document.getElementsByTagName("TD")[46].innerText
		
		; Country
		Country := wb.document.getElementsByTagName("TD")[48].innerText
		
		; Attention
		Attention := wb.document.getElementsByTagName("TD")[51].innerText
		
		; Phone
		Phone := wb.document.getElementsByTagName("TD")[53].innerText
		
		; Email
		Email := wb.document.getElementsByTagName("TD")[29].innerText
		
		; Customer notes 읽어오기
		CustomerNoteOnWeb := wb.document.getElementsByTagName("TD")[66].innerText
		StringUpper, CustomerNoteOnWeb, CustomerNoteOnWeb ; 고객 메모 대문자로 바꾸기		
		
		; Staff only notes 읽어오기 ;이건 에러나는데 필요하면 유튜브 
		StaffOnlyNote := wb.document.getElementsByTagName("TEXTAREA")[5].innerText
		StringUpper, StaffOnlyNote, StaffOnlyNote ; Staff only notes 대문자로 바꾸기


		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
		DeleteAdd2InAdd1(Address1)		
		
		Start()
		Start_Invoice_1()		
		
		MouseMove, 857, 313
		MouseGetPos, , , , control
		ControlGetText, SubTotal, %control%, LAMBS

		; Invoice 1 탭 클릭하기
		MouseClick,l , 41, 265




/*

		; Customer notes 읽어오기
		;CustomerNoteOnWeb := wb.document.getElementById("ctl00_ctl00_ContentPlaceHolder1_cphBody_orderItemDetail_tbCustomerNote").value
		CustomerNoteOnWeb := wb.document.getElementsByTagName("TD")[66].innerText
		StringUpper, CustomerNoteOnWeb, CustomerNoteOnWeb ; 고객 메모 대문자로 바꾸기
		
		; Staff only notes 읽어오기 ;이건 에러나는데 필요하면 유튜브 
		StaffOnlyNote := wb.document.getElementsByName("cancel_comment").innerText
		StringUpper, StaffOnlyNote, StaffOnlyNote ; Staff only notes 대문자로 바꾸기
*/				
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		

/*	
		run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL
		
;		MsgBox, % PO_L

		Sleep 1500
		
;		Run, https://admin.lashowroom.com/order_edit_v1.php?order_id=25166&list_option=new



		; 소스코드 창 열기 Ctrl + u
		SendInput, ^u
		Sleep 1500

		; 소스코드 전체선택 후 복사
		SendInput, ^a^c
		Sleep 1000
		
		htmlSourcecode := % Clipboard
		Sleep 1000

		; 소스코드 탭 닫기
		SendInput, ^w
		
		;HTMLSource = orderGeneralDetail_lblShipCompanyName">Goose Feathers</span></td>




		; LA쇼룸 페이지에서 정보 읽어오기
		FindInfoInLASHOWROOM(htmlSourcecode)
*/		
		
		
		
		
		
/*
	; LA쇼룸 정보가 LAMBS와 맞는지 묻고 Yes 눌렀으면 LAMBS에서 정보 얻는 GetInfoFromLAMBS() 함수 호출
	MsgBox, 4100, UPS Label Print out, 고객 메모 확인`n`nIs Info on LASHOWROOM same as LAMBS?
	IfMsgBox, Yes
		GetInfoFromLAMBS()
	else IfMsgBox, No
		Reload
*/



		return
}