#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

;#Include function.ahk



CConFashiongo(PO_F, FGServer){
	

/*	
	URL = http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%PO_F%
	;URL = http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR1C14B4FA11
	
	
	
	wb := ComObjCreate("InternetExplorer.Application")

;	ComObjConnect(wb, "IE_")

	wb.Visible := true
	wb.Navigate(URL)

	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 10	
*/	
	
	
	
/*	
	wbGet(WinTitle="ahk_class IEFrame", Svr#=1) {               ;// based on ComObjQuery docs
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
*/	
	
	
	
	
	; PO 가 맞는 창을 찾기 위해 CurrentFashionGoIEWinTitle 에 맞는 url 할당
	CurrentFashionGoIEWinTitle = http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%PO_F%
		
	
	; 현재 존재하는(이미 열어둔) IE창 접속 하기 위해 함수 호출
	; 현재 존재하는 창 중에 맞는 PO를 찾아서 접속
	; 단, wb를 전역변수로 선언해야 작동함.
	; 아래 세 함수 모두 작동함
	;WBGet(WinTitle="ahk_class IEFrame", Svr#=1)             ;// based on ComObjQuery docs
	;IEGet(name="")
	;wb := IEGet("http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR1C26891212 - Internet Explorer")
	
	
	; 이미 열려있는 CurrentFashionGoIEWinTitle 값의 url 주소가 있는 IE 창 사용하기
	wb := IEGet(CurrentFashionGoIEWinTitle)
	
	
	; 만약 맞는 PO 창이 없으면 창 새로 만들고 로그인하기
	if(!wb){
		
				
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
		}
	}
	
		;MsgBox, "%PO_F%" PAGE IS NOT EXIST
	


	;IEGet(name = %CurrentFashionGoIEWinTitle%)
	
	;IEGet(CurrentFashionGoIEWinTitle)
	;WBGet(WinTitle = "http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%PO_F% - Internet Explorer", Svr#=1)             ;// based on ComObjQuery docs
	;WBGet(WinTitle = "http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR1C26891212 - Internet Explorer", Svr#=1)             ;// based on ComObjQuery docs
	
	
	
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
	
	;OrganizingCCinfo()



	;MsgBox, CC info `n`n %CCinfoOnFASHIONGOVal%
	
	
	return
}

/*
OrganizingCCinfo(){
	
	MsgBox, function starts
		
	; CCinfoOnFASHIONGO.txt 내용을 RawCCinfo 변수에 저장하기
	FileRead, RawCCinfo, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt
	
	; CCNumbers, ExpDate, CVV
	
	MsgBox, % RawCCinfo
	
	;CCNumbers := RegExMatch(RawCCinfo, "mU)(lblShipAttention.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
	FoundPos := RegExMatch(RawCCinfo, "mU)(CardNumber.\s(.*) CVV)", SubPat)

	; CCNumbers 찾았으면 전역변수인 CCNumbers 에 값 넣기
	if(FoundPos){
		CCNumbers := SubPat2
		
		; CCNumbers 에서 모든 - 제거하기.
		StringReplace, CCNumbers, CCNumbers, -, , All
		
		MsgBox, 함수 안에서 찾은 것 CCNumbers`n%CCNumbers%
		
		;FileAppend, %CCNumbers%, %A_ScriptDir%\CreatedFiles\CCNumbers.txt
		FileAppend, %CCNumbers%, C:\Users\JODIFL4\Desktop\000000000\LAMBS\CreatedFiles\CCNumbers.txt
		;FileRead, CCNumbers, %A_ScriptDir%\CreatedFiles\CCNumbers.txt
		FileRead, CCNumbers, C:\Users\JODIFL4\Desktop\000000000\LAMBS\CreatedFiles\CCNumbers.txt


;C:\Users\JODIFL4\Desktop\000000000\LAMBS\CreatedFiles
		MsgBox, 함수 안에서 찾은 것 CCNumbers`n%CCNumbers%
		
		;MsgBox, % FoundPos
	}
	
	
		
	return
}
*/





aaIEGet(Name, WindowTitleToCompare){
		IfEqual, Name, , "http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR1C26891212 - Internet Explorer", Name, ahk_class IEFrame     ;// Get active window if no parameter
			Name := (Name="New Tab - Windows Internet Explorer")? "about:Tabs":RegExReplace(Name, " - (Windows|Microsoft)? ?Internet Explorer$")
		for wb in ComObjCreate("Shell.Application").Windows()
			if wb.LocationName=Name and InStr(wb.FullName, "iexplore.exe")
				return wb
	}