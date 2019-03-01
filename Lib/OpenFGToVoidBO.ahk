#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include GetActiveBrowserURL.AHK



OpenFGToVoidBO(CurrentPONumber){
	
	;~ MsgBox, % CurrentPONumber
	

	; ### 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기 ###
	;WinClose, http://vendoradmin%FGServer%.fashiongo.net/OrderDetails.aspx?po=%CurrentPONumber% - Internet Explorer
	
	; ### 혹시 이전에 같은 PO 창이 열려있었으면 닫고 시작하기 ###
	; 이전 IE 창 열려있으면 자꾸 이런저런 에러 나서 아예 IE 창들을 모두 그룹으로 묶고 한꺼번에 닫고 시작
	;~ GroupAdd,ExplorerGroup, ahk_class CabinetWClass
	GroupAdd,ExplorerGroup, ahk_class IEFrame
	WinClose,ahk_group ExplorerGroup


;~ /* 이거 작동하긴 했는데 이 코드 추가한 뒤 에러가 나는 것 같음
	; Do you want to close all tabs 경고창 나오면 Close all tabs 버튼 눌러서 IE창 모두 닫기
	Sleep 200
	IfWinExist, Internet Explorer
	{
		WinActivate, Internet Explorer
		ControlSend, Button1, {Enter}, Internet Explorer
	}
*/
	; 0x800704a6 - A system shut down has already been scheduled. 에러 나는 것 처리 위해
	; #################################################################################
	; https://autohotkey.com/board/topic/97731-unknown-error/page-2
	; https://autohotkey.com/board/topic/89355-suspend-script-while-process-is-running/
	; #################################################################################
	Process , Exist, iexplore.exe ;
	while (ErrorLevel){
	suspend, on
	Process, Exist, iexplore.exe
	sleep, 1000
	}
	suspend, off

	Clipboard :=



	; New Orders 검색화면 열기
	Loginname = customer3
	Password = Jo123456789
	;~ URL = https://vendoradmin.fashiongo.net/#/order/orders/new
	URL = https://vendoradmin.fashiongo.net/#/order/orders
	WinMaximize, ahk_class IEFrame

	WB := ComObjCreate("InternetExplorer.Application")
	WB.Visible := true
	WB.Navigate(URL)
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 1000



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
			
		;~ wb.document.getElementById("tbUserID").value := Loginname  ;ID 입력
		;~ wb.document.getElementsByTagName("INPUT")[0].innerText := Loginname ;ID 입력
		
		;~ wb.document.getElementById("tbPassword").value := Password ; 비밀번호 입력
		;~ wb.document.getElementsByTagName("INPUT")[1].innerText := Password ;비밀번호 입력
			
			
		; 얘가 기계로 임력하는 걸 아는건지 위와 같은 일반적인 방법으로 로그인 하려고 하면 자꾸 에러가 나서 다음과 같이 이동 후 입력하는 방법 사용
		wb.document.getElementsByTagName("INPUT")[0].focus() ;ID 입력
		SendInput, % Loginname
		Sleep 100
			
		wb.document.getElementsByTagName("INPUT")[1].focus() ;비밀번호 입력
		SendInput, % Password
		Sleep 100
			
		wb.document.getElementsByTagName("BUTTON")[0].Click() ; 로그인 버튼 누르기
			
			
		; 로그인 후 New Orders 검색화면 열기
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100

		Sleep 500
		;~ wb.document.getElementsByTagName("I")[4].Click() ; 메뉴 바의 All Orders 의 ˅ 버튼 누르기
		;~ wb.document.getElementsByTagName("A")[13].Click() ; All Orders 안에 있는 New Orders 버튼 누르기
		
		
		wb.document.getElementsByTagName("DIV")[33].Click() ; CLICK All Orders
			

		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
		
	}
	
	
	; 검색창에 PO 번호 입력하기
	wb.document.getElementsByTagName("INPUT")[0].focus()
	SendInput, % CurrentPONumber
	Sleep 500
	Send, {Enter}
	


		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
		

	; Serch 의 Company Name 드롭다운리스트 값을 PO Number로 바꾸기
	wb.document.getElementsByTagName("SELECT")[1].focus()
	Send, {Down}
	Send, {Enter}
	Sleep 100
	

		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
			


	; Serch 의 Company Name 드롭다운리스트 값을 PO Number로 바꾸기
	wb.document.getElementsByTagName("SELECT")[0].focus()
	Send, {Down}
	Send, {Down}
	Send, {Enter}
	



	return
}

