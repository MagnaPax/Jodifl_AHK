; #############################################################################
; Drop Style 하기

; 1. jodifl.com 에서 해당 아이템 내리기
; 2. FashionGO 에서 내리기
; 3. LA Showroom 에서 내리기
; #############################################################################



#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)


#Include %A_ScriptDir%\lib\

#Include DS_Lib.ahk



; 메소드에서 오는 driver를 포함한 값들을 받기 위한 배열 선언
Array := object()
	
; 배열 초기화 해주기
Array := []



/*
P2392	BURGUNDY	; 아이템 여러개 있음


******* NAVY 아닌 NAVY MIX 로. OFF WHITE 아닌 OFF WHITE MIX 등으로 정확한 색깔값 입력해야 됨 *******

B1413	off white MIX	; 이 색깔 하나밖에 없어서 LAS에서 UNDISPLAY 하게됨. Magento에서도 모든 색깔이 disabled 되었으므로 전체를 disable 해줌
J2927	NAVY			; LAS의 디폴트 색깔이라 사진을 바꿨음.
P2274	SAGE/IVORY		; LAS의 디폴트 색깔이라 사진을 바꿨음. Magento에 검색값 없어서 경고창 뜸
*/



;GUI Backgroud
;~ Gui, Show, w350 h150, N41 Processing, AlwaysOnTop Window
Gui, Show, w245 h160, Drop Styles, AlwaysOnTop Window
WinSet, AlwaysOnTop, On, Drop Styles

;Input Start Order Id
Gui, Add, Text, x22 y21 Cred , Drop Style #
Gui, Add, Edit, x92 y20 w100 h20 vDropStyle#, H3074 ;P2392 	;C1116-8 ;P2276-1 ;b1323 ;P2471 ;B3595-1(drop all) ;P4030 ;p2420-12(DROP ALL)             				            ;P4061 ;B3610 ;S1229 ;P2933

;Input End Order Id
Gui, Add, Text, x22 y51 CBlue , Color
Gui, Add, Edit, x92 y50 w100 h20 vStyleColor, taupe ;BURGUNDY ;navy mix ; olive mix

;Check Box For Drop All
Gui, Add, CheckBox, x22 y70 w80 h40 vDropAll, Drop All

;엔터 버튼
Gui, Add, Button, x22 y110 w200 h30 +default gClick_btn, Enter





;GUI시작 시 포커스를 DropStyle# 칸에 위치
;~ GuiControl, Focus, CustomerPO
GuiControl, Focus, DropStyle#


return



Click_btn:


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; To use the values which input on GUI
	Gui Submit, nohide
	GUI, Destroy
	
	
	; 만약 값을 입력 안 했으면 프로그램 다시 시작하기
	if(DropStyle# == ""){
		Reload		
	}
	
	; Drop All 이 체크되어 있지 않을때만 색깔 값이 없으면 다시 시작하기
	; Drop All 이 체크되어 있으면 색깔을 입력할 필요가 없으니까
	if(!DropAll){	
		if(StyleColor == "")
			Reload
	}	



	; 혹시나 있을지도 모를 아이템 번호값과 색깔 값의 양쪽 공백 없애기
	DropStyle# := Trim(DropStyle#)
	StyleColor := Trim(StyleColor)	
	StringUpper, DropStyle#, DropStyle# ; 대문자로 바꾸기
	StringUpper, StyleColor, StyleColor ; 대문자로 바꾸기


	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.AddArgument("--start-maximized") ; 창 최대화 하기	
	driver.Get("about:blank") ; 창 열기

;	MsgBox, % DropStyle# . "`n" . StyleColor




/*
	; ##################################################################
	; ###################### LAS 에서 아이템 내리기 ######################
	; ##################################################################

	; LAS 의 아이템 삭제하는 창에 들어가기
	Array := openItemListOnLAS(driver, DropStyle#, StyleColor)
	driver := Array[1] ; 배열에 있는 driver 를 driver 변수로 옮긴다
	thePageBeingUsed := Array[2] ; 현재 페이지 url 주소
;MsgBox, % "현재 페이지 url 주소 : " . thePageBeingUsed

	; 제대로 된 작업창인지 확인 후 아니면 그곳으로 이동
	driver := checkURLBeforeStartThisMethod(driver, thePageBeingUsed)

	; LAS 의 Item List 의 Search 칸에 아이템 번호 입력한 뒤 검색하기
	driver := putDropStyle#IntoTheSearchOnLAS(driver, DropStyle#)

	; 이렇게 탭 타이틀을 저장해 놓으면 다음에 SwitchToWindowByTitle 명령을 이용해 이 탭을 사용할 수 있음	
	tab_title_of_Item_Search := driver.title
	URLofItemSearchPage := driver.Url
		
;~ MsgBox, % "ITEMS 의 Search 페이지의 탭 주소`n"tab_title_of_Item_Search
;MsgBox, % "ITEMS 의 Search 페이지의 URL`n" . URLofItemSearchPage


	; 제대로 된 작업창인지 확인 후 아니면 그곳으로 이동
	driver := checkURLBeforeStartThisMethod(driver, thePageBeingUsed)
	



	; Drop All 이 체크되어 있지 않으면 정상적으로 아래 실행
	if(!DropAll){
		; 검색된 아이템 중 DropStyle# 에 딱 맞는 아이템 클릭해서 들어가기. 예를들어 P4030 으로 검색했으면 P4030, P4030-1, P4030-2 중에서 P4030만 딱 짚어서 클릭해서 들어가기
		;~ driver := clickExactSameStyle#OfDropStyle#OnLAS(driver, DropStyle#)
		Array := clickExactSameStyle#OfDropStyle#OnLAS(driver, DropStyle#)
		driver := Array[1] ; 배열에 있는 driver 를 driver 변수로 옮긴다
		thePageBeingUsed := Array[2] ; 현재 페이지 url 주소
;MsgBox, % "현재 페이지 url 주소 : " . thePageBeingUsed


		; 제대로 된 작업창인지 확인 후 아니면 그곳으로 이동
		driver := checkURLBeforeStartThisMethod(driver, thePageBeingUsed)
		
		
		; 지울 색깔이 디폴트 칸에 있는지 확인. 디폴트 칸에 지울 StyleColor 컬러가 있으면 1 없으면 0을 반환한다
		Array := isTheStyleColorInTheDefaultCell(driver, StyleColor)		
		driver := Array[1] ; 배열에 있는 driver 를 driver 변수로 옮긴다
		isTheStyleColorInTheDefaultCell := Array[2] ; 디폴트 칸에 StyleColor 컬러가 있으면 1 없으면 0 저장
;MsgBox, % "디폴트 칸에 StyleColor 컬러값인 " . StyleColor .  "가 있으면 1 없으면 0`n`n" . isTheStyleColorInTheDefaultCell


		; drop 할 색깔 찾아서 색깔 입력칸의 그 색깔 지우기
		driver := deleteTheColorOfDropStyleOnLAS(driver, StyleColor)
		
		; 색깔 입력칸 전체에 색깔 값이 들어있는지 찾기
		Array := checkAllColorCells(driver)
		driver := Array[1] ; 배열에 있는 driver 를 driver 변수로 옮긴다
		isItAllEmptyOfColorCells := Array[2] ; 색깔 칸에 컬러값이 있으면 1 하나도 없으면 0 저장
;MsgBox, % "색깔 칸에 컬러값이 하나라도 있으면 1 하나도 없으면 0`n`n" . isItAllEmptyOfColorCells
					


		; 색깔 입력칸에 색깔이 하나도 없을 때
		; 이 스타일 전체 undisplyed 시키기
		if(isItAllEmptyOfColorCells == 0){
			
			; 이 스타일 전체 undisplyed 시키기
			;~ driver := makeThisStyleundisplyed(driver, DropStyle#, tab_title_of_Item_Search)
			driver := makeThisStyleundisplyed(driver, DropStyle#, URLofItemSearchPage)
			
		} ; end끝 - if(isItAllEmptyOfColorCells == 0)
		
		
		; 색깔 입력칸에 색깔이 남아 있을 때
		; drop 한 색깔이 디폴트 칸에 있었다면 Set Default Image 에서 사진을 다른 것으로 바꿔주고 저장
		; drop 한 색깔이 디폴트 칸이 아닌 다른 칸에 있었다면 사진 바꿀 필요 없이 그냥 저장
		else if(isItAllEmptyOfColorCells == 1){
			
			; drop 한 색깔이 디폴트 칸에 있었다면 Set Default Image 에서 사진을 다른 것으로 바꿔주고 저장(위에서 이미 색깔 값을 지웠기 때문에 다시 지울 필요 없음)
			if(isTheStyleColorInTheDefaultCell){
				
				; Set Default Image 의 사진 다른 것으로 바꾸기
				driver := changeSetDefaultImage(driver)
				
			} ; end끝 - if(isTheStyleColorInTheDefaultCell)
			


			
			
			; 사진을 바꿨든 안 바꿨든 Save 버튼 클릭해서 저장하기
			driver := clickSaveButtonOnLAS(driver)

		} ; end끝 - else if(isItAllEmptyOfColorCells == 1)
		
	}


	; Drop All 이 체크되어 있으면 그냥 Undisplay 하기
	else if(DropAll){
		
		; 이 스타일 전체 undisplyed 시키기
		;~ driver := makeThisStyleundisplyed(driver, DropStyle#, tab_title_of_Item_Search)
		driver := makeThisStyleundisplyed(driver, DropStyle#, URLofItemSearchPage)
	}



*/











;~ /*

	; ######################################################################
	; ###################### Magento 에서 아이템 내리기 #####################
	; ######################################################################


	; Magento 의 아이템 삭제하는 창에 들어가기
	driver := openItemListOnMagento(driver, DropStyle#, StyleColor)
	

	; Name 에 아이템 번호 입력한 뒤 검색하기
	driver := putDropStyle#IntoTheSearchOnMagento(driver, DropStyle#)
	;~ Sleep 4000
	;~ driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
	
	
	; Drop All 이 체크되어 있지 않으면 정상적으로 아래 실행
	if(!DropAll){


		; 검색된 아이템 중 DropStyle# 와 StyleColor 가 맞는 아이템의 체크박스 체크하기. 예를들어 P4030 INK BLACK 으로 검색했으면 P4030, P4030-1, P4030-2 중에서 P4030의 INK BLACK 만 딱 짚어서 해당 체크박스 체크하기
		IsItRecursiveCall = 0
		driver := checkTheCheckBoxOfCorrectItemOnMagento(driver, DropStyle#, StyleColor, IsItRecursiveCall)
		
MsgBox, pause to disable a item

		; 위에서 체크된 아이템 상태 disable 로 만들기
		driver := disableTheStatusOfTheStyle(driver)

		; 모든 색깔이 Drop 됐는지 체크한 뒤 체크됐으면 drop 하는 메소드
		driver := isAllColorsDorppedOnMagento(driver, DropStyle#, DropAll)
		
	} ; END - if(!DropAll) 끝
	
	; Drop All 이 체크되어 있으면 모두 내리기
	else if(DropAll){
;MsgBox, Drop All 이 체크됐으므로 모두 내려야 됨
	
		; 모든 색깔이 Drop 됐는지 체크한 뒤 체크됐으면 drop 하는 메소드
		driver := isAllColorsDorppedOnMagento(driver, DropStyle#, DropAll)
	}

*/














;~ /*
	; #################################################################
	; ###################### FG 에서 아이템 내리기 ######################
	; #################################################################

	; FG 의 아이템 삭제하는 창에 들어가기
	;~ driver := styleDropFromFG(DropStyle#, StyleColor)
	driver := openItemListOnFG(driver, DropStyle#, StyleColor)
	
	; FG 의 Item List 의 Search 칸에 아이템 번호 입력한 뒤 검색하기
	driver := putDropStyle#IntoTheSearchOnFG(driver, DropStyle#)


	; 검색된 아이템 중 DropStyle# 에 딱 맞는 아이템 클릭해서 들어가기. 예를들어 P4030 으로 검색했으면 P4030, P4030-1, P4030-2 중에서 P4030만 딱 짚어서 클릭해서 들어가기
	driver := clickExactSameStyle#OfDropStyle#(driver, DropStyle#)


	; Drop All 이 체크되어 있지 않으면 정상적으로 아래 실행
	if(!DropAll){
		
		; drop 할 색깔 찾아서 X 눌러서 그 색깔 삭제하기
		driver := deleteTheColorOfDropStyle(driver, DropStyle#, StyleColor)

		; 모든 색깔이 Drop 됐는지 체크하는 메소드
		driver := isAllColorsDorpped(driver)
	}
	; Drop All 이 체크되어 있으면 Inactive 체크하기
	else if(DropAll){
		
		; Inactive 라디오 버튼 클릭
		Xpath = (//div[@class='check-circle'])[2]
		driver.FindElementByXPath(Xpath).click()
		Sleep 300	
	}
	
MsgBox, TO SAVE
	; Save 버튼 클릭해서 저장하기
	driver := clickSaveButton(driver, DropStyle#, StyleColor, DropAll)

;MsgBox, 262144, ㅇㅇ, 저장됐나?



*/







SoundPlay, %A_WinDir%\Media\Ring06.wav
MsgBox, 262144, Title, JOB FINISHED. THE STYLE '%DropStyle#% %StyleColor%' HAS BEEN DROPPED.






	
	Reload
	
	
	F5::
	Reload
	

	Esc::
	Exitapp	
	
	
