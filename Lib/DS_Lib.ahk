#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.





;~ #Include CNewBrow.ahk
#Include CNewBrow_DS.ahk


; ########################################################################################################################################################################################################################
; ############################################################################################## 공통 처리 ################################################################################################################
; ########################################################################################################################################################################################################################

; 작업해야 되는 창이 로그아웃 되거나 하지 않고 제대로 유지하고 있는지 확인하는 메소드
checkURLBeforeStartThisMethod(driver, thePageBeingUsed){

;MsgBox, % "메소드 들어옴. 리프레쉬 예정`n현재 페이지와 thePageBeingUsed 에 들어있는 주소`n`n" . driver.Url . "`n" . thePageBeingUsed


	; 현재 페이지 리프레쉬하기 
	driver.refresh()
	
	;~ Sleep 3000
	
	; 현재 페이지가 작업해야 되는 페이지라면 문제 없다는 뜻이므로 메소드 끝내고 나감
	;~ if(RegExMatch(driver.Url, thePageBeingUsed)){
	if(driver.Url == thePageBeingUsed){
;MsgBox, 제대로 된 페이지에 머물러 있음
		return driver
	}
	; 만약 현재 페이지가 작업해야 되는 페이지가 아닌데
	; 작업해야 되는 페이자가 FG라면
	else if(RegExMatch(thePageBeingUsed, "imU)fashiongo")){
MsgBox, % "문제 발생. 현재 페이지가 작업페이지가 아님. FG 로그인 다시 한 뒤 작업 페이지 다시 열어야 됨`n`n현재 페이지와 thePageBeingUsed에 들어있는 주소 비교`n`n" . driver.Url . "`n" . thePageBeingUsed
		
		; FG 로그인 하기
		driver := FG_LoginOnly(driver)
	}	
	
	; 만약 현재 페이지가 작업해야 되는 페이지가 아닌데
	; 작업해야 되는 페이자가 LAS라면
	else if(RegExMatch(thePageBeingUsed, "imU)lashowroom")){
MsgBox, % "문제 발생. 현재 페이지가 작업페이지가 아님. LAS 로그인 다시 한 뒤 작업 페이지 다시 열어야 됨`n`n현재 페이지와 thePageBeingUsed에 들어있는 주소 비교`n`n" . driver.Url . "`n" . thePageBeingUsed
		
		; LAS 로그인 하기
		driver := LAS_LoginOnly(driver)
	}
	
	; 만약 현재 페이지가 작업해야 되는 페이지가 아닌데
	; 작업해야 되는 페이자가 Magento라면
	else if(RegExMatch(thePageBeingUsed, "imU)jodifl.com")){
MsgBox, % "문제 발생. 현재 페이지가 작업페이지가 아님. Magento 로그인 다시 한 뒤 작업 페이지 다시 열어야 됨`n`n현재 페이지와 thePageBeingUsed에 들어있는 주소 비교`n`n" . driver.Url . "`n" . thePageBeingUsed
		
		; Magento 로그인 하기
		driver:= Magento_LoginOnly(driver)
	}	


	; 만약 현재 페이지가 작업해야 되는 페이지가 아니라면 위의 else 안의 각 로그인 작업을 실행한 뒤
	; 작업해야 되는 주소로 이동
	driver.Get(thePageBeingUsed)
	
	
	; 혹시 또 문제 생겼을 지 모르니까 현재 페이지 다시 확인해보기
	; 현제 페이지가 작업 페이지와 다를 경우 재귀호출 하기
	if(driver.Url != thePageBeingUsed){
		checkURLBeforeStartThisMethod(driver, thePageBeingUsed)
	}
	
	return driver
}



















; ########################################################################################################################################################################################################################
; ################################################################################################################### FG 처리 ############################################################################################
; ########################################################################################################################################################################################################################

; FG 의 아이템 삭제하는 창에 들어가기
openItemListOnFG(driver, DropStyle#, StyleColor){


	driver := FG_LoginOnly(driver) ; FG 로그인
	Sleep 500

	; Item List 페이지로 이동
	URL = https://vendoradmin.fashiongo.net/#/item/editall
	driver.Get(URL)
	Sleep 500


	; 크롬 창 열어서 FG의 Item List 주소로 이동
;	driver := goToURl_AfterLogIn_IfNeeded(driver, URL)


	; 아이템 삭제하는 창이 아니면
	if(!RegExMatch(driver.Url, "https://vendoradmin.fashiongo.net/#/item/editall")){
		
		; chrome_Profile.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;		FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\NewWindowProcessing\chrome_Profile.txt, 1
		
;MsgBox , 아이템 삭제하는 창이 아닙니다. chrome_Profile.txt 파일의 내용을 삭제했고 openItemListOnFG(DropStyle#, StyleColor) 메소드를 다시 호출합니다.
		
;		openItemListOnFG(driver, DropStyle#, StyleColor)


;		driver.close() ; closing just one tab of the browser
;		driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
;		driver.AddArgument("--start-maximized") ; 창 최대화 하기

MsgBox , 아이템 삭제하는 창이 아닙니다. 재귀 호출로 로그인 다시 합니다.

		openItemListOnFG(driver, DropStyle#, StyleColor)
		
;		driver := FG_LoginOnly(driver) ; FG 로그인
;		Sleep 500
;		driver.Get(URL)	 ; Item List 로 이동
	}
	
;MsgBox, 아이템 삭제하는 창에 제대로 들어옴
	
	
	return driver
}




; FG 의 Item List 의 Search 칸에 아이템 번호 입력한 뒤 검색하기
putDropStyle#IntoTheSearchOnFG(driver, DropStyle#){
	
	; 검색란에 DropStyle# 입력하기
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-editall/fg-item-search/div/div[2]/div/input[2]
	driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(DropStyle#).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 DropStyle# 입력 후 엔터쳐서 이동하기
	
	; 검색된 값이 없으면 에러 메세지 출력하고 프로그램 재시작하기
	Sleep 300
	if(!driver.FindElementByXPath("//*[text() = '" DropStyle# "']")){
		
		MsgBox, 262404, Drop Style ERROR, THE DROP STYLE IS NOT FOUNDED ON THIS FG PAGE. IF YOU'D LIKE TO CONTINUE THE PROCESSING, PLEASE CLICK 'YES' BUT IF YOU WANT TO RESTART THIS APP, PLEASE CLICK 'NO'`n`n지울 아이템 번호가 없습니다. 확인해 주시기 바랍니다. Yes 누르면 계속 진행하고 No 누르면 프로그램을 중지합니다.
	
		IfMsgBox, No
			Reload


	}
	
;MsgBox 제대로 입력됐남?
	
	return driver
}


; 검색된 아이템 중 DropStyle# 에 딱 맞는 아이템 클릭해서 들어가기. 예를들어 P4030 으로 검색했으면 P4030, P4030-1, P4030-2 중에서 P4030만 딱 짚어서 클릭해서 들어가기
clickExactSameStyle#OfDropStyle#(driver, DropStyle#){
	
	if(driver.FindElementByXPath("//*[text() = '" DropStyle# "']"))
		driver.FindElementByXPath("//*[text() = '" DropStyle# "']").click()
	
	
	; 아이템을 찾아서 클릭하지 못했으면 찾아서 클릭 할때까지 에러 메세지 계속 내보내기
	Sleep 500
	Loop{
		if(!RegExMatch(driver.Url, "item/detail")){
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, NO DROP COLOR ERROR, THERE IS NO DROP STYLE ON THIS PAGE OR THIS APP CAN'T FIND IT`n`nPLEASE FIND THE STYLE AND CLICK IT THEN CLICK OK BUTTON TO CONTINUE`n`n`n해당 아이템 번호를 찾지 못하였습니다. 수동으로 아이템을 클릭한 뒤 진행하시기 바랍니다.
			continue
		}
		break
	}
	
;MsgBox 제대로 클맀했나?
	Sleep 300
	return driver
}



; drop 할 색깔 찾아서 X 눌러서 그 색깔 삭제하기
deleteTheColorOfDropStyle(driver, DropStyle#, StyleColor){

	if(driver.FindElementByXPath("//div[@class='label margin-right-5' and text() = '" StyleColor "']")){
;MsgBox, 찾았음. 클릭할 예정

		driver.FindElementByXPath("//div[@class='label margin-right-5' and text() = '" StyleColor "']").click() ; 이렇게 색깔을 먼저 클릭해야 옆의 x 가 표시돼기 때문에 일단 이것 클릭 후 밑에서 x 클릭해서 삭제		
		Sleep 300
		driver.FindElementByXPath("//div[@class='label margin-right-5' and text() = '" StyleColor "']//following-sibling::i").click() ; 지우려는 색깔 옆의 X 클릭
		Sleep 2000
		driver.FindElementByXPath("//div[@class='label margin-right-5' and text() = '" StyleColor "']//parent::div//preceding-sibling::div//descendant::div[@class='check-square']").click() ; 지우려는 색깔의 체크박스 클릭(체크가 이미 되어있으니 결과적으로 체크박스 해제)
		Sleep 300
	}
	else
	{
		SoundPlay, %A_WinDir%\Media\Ring06.wav		
		MsgBox, 262404, NO DROP COLOR ERROR, THE COLOR OF THE DROP STYLE IS NOT FOUNDED ON THIS FG PAGE. IF YOU'D LIKE TO CONTINUE THE PROCESSING, PLEASE CLICK 'YES' BUT IF YOU WANT TO RESTART THIS APP, PLEASE CLICK 'NO'`n`n지울 아이템 색깔을 못 찾았습니다. 확인해 주시기 바랍니다. Yes 누르면 계속 진행하고 No 누르면 프로그램을 중지합니다.
	
		IfMsgBox, No
			Reload				
		
	}

;MsgBox, 색깔 제대로 삭제됐나?

	return driver
}


; 모든 색깔이 Drop 됐는지 체크하는 메소드
isAllColorsDorpped(driver){

	; 체크된 색깔의 체크박스가 있는지 확인하기 위한 변수
	itHasToBeActivated = 0
	
	Loop{
		
		; 화면에 있는 여러 체크박스 중에 색깔에 해당되는 체크박스만 확인하기 위한 Xpath
		Xpath := "(/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-item/div[4]/div[2]/div[2]/fg-item-edit-colors/div/ul//child::input[@type='checkbox'])" . "[" . A_Index . "]"
		
		
		; 색깔의 체크박스가 더이상 없으면 loop 끝내고 나감
		if(driver.FindElementByXPath(Xpath).isDisplayed() != 0){
;MsgBox, 체크박스가 더이상 없음. 루프 끝내고 나감. NOT DISPLAYED
			break
		}
				
		
;MsgBox, % A_Index . " 번째 체크박스가 체크 되어있으면 -1을 리턴하고 안 되어있으면 0을 리턴한다`n" . driver.FindElementByXPath(Xpath).isSelected() . "`n`n`nitHasToBeActivated : " . itHasToBeActivated
		
		
		; 체크박스가 한 개라도 체크되어 있으면 itHasToBeActivated 변수값에 1 넣어줌
		if(driver.FindElementByXPath(Xpath).isSelected() == -1){
			itHasToBeActivated = 1
;MsgBox, % "이 화면에서 체크된 체크박스가 별견되었으므로 itHasToBeActivated 변수에 1을 넣어줬음.`nitHasToBeActivated : " . itHasToBeActivated . "`n`n체크박스가 최소한 한 개는 체크되어 있기 때문에 이 아이템은 Inactive 대상이 아님. 이후의 다른 체크박스 살필 필요 없이 루프 끝내고 나감"
			break
		}

/*
		if(itHasToBeActivated == 1){
;MsgBox, 체크박스가 최소한 한 개는 체크되어 있기 때문에 이 아이템은 Inactive 대상이 아님. 이후의 다른 체크박스 살필 필요 없이 루프 끝내고 나감
			break
		}
*/		
	} ; loop 끝
	
	
	if(itHasToBeActivated == 0){
;MsgBox, 모든 체크 박스의 체크가 풀려 있으므로 이 아이템은 Inactive 되어야 됨
		
		; Inactive 라디오 버튼 클릭
		Xpath = (//div[@class='check-circle'])[2]
		driver.FindElementByXPath(Xpath).click()
		Sleep 300
		
	}	
	
	return driver
}


; Save 버튼 클릭해서 저장하기
clickSaveButton(driver, DropStyle#, StyleColor, DropAll){
	
;MsgBox, TO SAVE

	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-item/div[8]/div[2]/button
	driver.FindElementByXPath(Xpath).click()

	; Drop All 이 아닐때만 아래 창 띄워서 이미지 지우게 하기
	if(!DropAll){
		MsgBox, 262144, Delete Images, PLEASE DELETE IMAGES OF '%DropStyle#%  %StyleColor%' THEN CLICK OK BUTTON TO CONTINUE.`n`n'%DropStyle#%  %StyleColor%' 이미지를 삭제한 뒤 OK 버튼을 눌러서 진행해 주십시오.
		driver.FindElementByXPath(Xpath).click() ; 사진을 지운 뒤 저장하기 위해
	}
	
	
	return driver
}










; ########################################################################################################################################################################################################################
; ################################################################################################################### LAS 처리 ############################################################################################
; ########################################################################################################################################################################################################################


openItemListOnLAS(driver, DropStyle#, StyleColor){
	
	; Item List 의 URL 주소
	URL = https://admin.lashowroom.com/item_search.php?action=search
	
	; 크롬 창 열어서 LAS의 ITEMS-Search 주소로 이동
	driver := goToURl_AfterLogIn_IfNeeded(driver, URL)

;MsgBox, % "아이템 삭제하는 창에 들어왔나?`n`n" . driver.Url
	
	return [driver, driver.Url]
}



; LAS 의 ITEMS 의 Search 칸에 아이템 번호 입력한 뒤 검색하기
putDropStyle#IntoTheSearchOnLAS(driver, DropStyle#){
	
	; 작업해야 되는 창이 로그아웃 되거나 하지 않고 제대로 유지하고 있는지 확인하는 메소드
	;~ driver := checkURLBeforeStartThisMethod(driver, thePageBeingUsed)
	
	; 검색란에 DropStyle# 입력하기
	Xpath = //input[@id='sn_0']
	driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(DropStyle#).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 DropStyle# 입력 후 엔터쳐서 이동하기
	
	; 검색된 값이 없으면 에러 메세지 출력하고 프로그램 재시작하기
	Sleep 300
	if(!driver.FindElementByXPath("//*[@id='edit_form']/table/tbody/tr[4]/td/ul/li/div//*[@*='" DropStyle# "']")){
		
		MsgBox, 262404, Drop Style ERROR, THE DROP STYLE IS NOT FOUNDED ON THIS LAS PAGE. IF YOU'D LIKE TO CONTINUE THE PROCESSING, PLEASE CLICK 'YES' BUT IF YOU WANT TO RESTART THIS APP, PLEASE CLICK 'NO'`n`n지울 아이템 번호가 없습니다. 확인해 주시기 바랍니다. Yes 누르면 계속 진행하고 No 누르면 프로그램을 중지합니다.
	
		IfMsgBox, No
			Reload			
		
	}
	
;MsgBox 제대로 입력됐남?
	
	return driver
}


; 검색된 아이템 중 DropStyle# 에 딱 맞는 아이템 클릭해서 들어가기. 예를들어 P4030 으로 검색했으면 P4030, P4030-1, P4030-2 중에서 P4030만 딱 짚어서 클릭해서 들어가기
clickExactSameStyle#OfDropStyle#OnLAS(driver, DropStyle#){
	
	if(driver.FindElementByXPath("//*[@id='edit_form']/table/tbody/tr[4]/td/ul/li/div//*[@*='" DropStyle# "']")){
		
		; 새 창이 아닌 새 탭에서 열기 위해 Ctrl 누르고 그리로 이동하기 위해 Shift 누른다. (키보드로 하려면 두 키 동시에 누르고 클릭하면 됨)
		driver.sendKeys(driver.Keys.CONTROL)
		driver.sendKeys(driver.Keys.SHIFT)
		
		
		driver.FindElementByXPath("//*[@id='edit_form']/table/tbody/tr[4]/td/ul/li/div//*[@*='" DropStyle# "']").click()
		
		
		; 새 탭을 열기 위해 눌렀던 Ctrl 과 Shift 키 누른것 해제하기 위해
		driver.sendKeys(driver.Keys.CONTROL)
		driver.sendKeys(driver.Keys.SHIFT)


		driver.SwitchToNextWindow() ; 새로 연 탭으로 콘트롤 옮김
		
;MsgBox, % DropStyle# . " 아이템 수정 창의 주소`n" . driver.Url

	}
	
	
	; 아이템을 찾아서 클릭하지 못했으면 찾아서 클릭 하도록 안내 메세지 띄우기
	Sleep 500
	if(!RegExMatch(driver.Url, "old_item_edit")){
		SoundPlay, %A_WinDir%\Media\Ring06.wav
		MsgBox, 262144, NO DROP COLOR ERROR, THERE IS NO DROP STYLE ON THIS PAGE OR THIS APP CAN'T FIND IT`n`nPLEASE PUT THE STYLE NUMBER(%DropStyle#%) IN TO THE SEARCH FIELD AND PRESS ENTER KEY THEN CLICK OK BUTTON TO CONTINUE`n`n`n해당 아이템 번호를 찾지 못하였습니다. 스타일번호(%DropStyle#%)를 수동으로 검색한 뒤 OK 버튼을 눌러주시기 바랍니다.
		
		; 위에서 수동으로 검색한 뒤 재귀호출로 해당 아이템 수정 페이지로 넘어가기
		clickExactSameStyle#OfDropStyle#OnLAS(driver, DropStyle#)
	}
	

;MsgBox 제대로 클릭해서 아이템 에디트 탭으로 들어왔나?

	return [driver, driver.Url]
}


; 지울 색깔이 디폴트 칸에 있는지 확인
isTheStyleColorInTheDefaultCell(driver, StyleColor){
	
	; 색깔 입력칸 중 디폴트 칸에 있는 색깔 위치
	Xpath = //*[contains(text(), '(default image color name')]//parent::td//child::input
	
;MsgBox, % "디폴트 색깔 입력칸에 있는 값 : " . driver.FindElementByXPath(Xpath).Attribute("value")
	
	; 디폴트 칸에 StyleColor 컬러가 있으면 1을 리턴한다
	if(driver.FindElementByXPath(Xpath).Attribute("value") == StyleColor){
;MsgBox, % StyleColor . "와 " . driver.FindElementByXPath(Xpath).Attribute("value") . "이(가) 같다. 1을 리턴하고 메소드 끝내기"
		return [driver, 1]
	}
	
	; 디폴트 칸에 StyleColor 컬러가 없으면 0을 리턴한다
	return [driver, 0]
}


; drop 할 색깔 찾아서 색깔 입력칸의 그 색깔 지우기
deleteTheColorOfDropStyleOnLAS(driver, StyleColor){
	
	if(driver.FindElementByXPath("//input[@class='color-name-field ui-autocomplete-input' and @value= '" StyleColor "']")){
;MsgBox, 지울 색깔 찾았음. 값 지울 예정

		; 어떤 색깔을 지울건지 화면에 표시하기 위해 색깔을 일단 클릭(인간을 위한 동작이므로 삭제해도 무방)
		driver.FindElementByXPath("//input[@class='color-name-field ui-autocomplete-input' and @value= '" StyleColor "']").click
		Sleep 1000
		
		driver.FindElementByXPath("//input[@class='color-name-field ui-autocomplete-input' and @value= '" StyleColor "']").sendKeys(driver.Keys.CONTROL, "a").SendKeys(driver.Keys.Delete) ; 값을 지우기 위해 Ctrl+A 한 뒤 delete 키 눌러서 지우기
		Sleep 1000
	}
	else{
		MsgBox, 262144, NO DROP COLOR ERROR, THE DROP COLOR IS NOT NOT FOUNDED. PLEASE CHECK WHAT CORRECT COLOR IS. THE APP WILL BE RESTARTED.`n`n[%StyleColor%]`n`n지울 아이템 색깔이 없습니다. 확인해 주시기 바랍니다. 프로그램이 다시 시작합니다.
		Reload
	}
	
	return driver
}
		
		
		
; 색깔 입력칸 전체에 색깔 값이 들어있는지 찾기
checkAllColorCells(driver){
	

	; 색깔 입력칸을 돌면서 값이 있는 지 확인.
	; 색깔 입력칸은 총 15개. 값이 들어있는 색깔 입력칸이 발견되면
	Loop, 15{
		
		; 화면에 있는 여러 체크박스 중에 색깔에 해당되는 체크박스만 확인하기 위한 Xpath
		Xpath := "(//input[@class='color-name-field ui-autocomplete-input'])" . "[" . A_Index . "]"
		
		
;MsgBox, % A_Index . " 번째 색깔 입력칸의 값은 : " . driver.FindElementByXPath(Xpath).Attribute("value") . "`n`n`nitHasToBeActivated : " . itHasToBeActivated
		
		
		; 색깔 입력칸에 값이 한 개라도 있으면 1을 리턴하고 메소드 끝냄
		if(driver.FindElementByXPath(Xpath).Attribute("value")){
			itHasToBeActivated = 1
;MsgBox, % "이 화면에는 색깔 값이 최소한 한 개 이상 있으므로 1을 리턴하면서 메소드 끝냄"
			return [driver, 1]
		}
		
	} ; end끝 - Loop, 15			
	
	; 색깔 입력칸에 값이 하나도 없으면 0 리턴하고 끝냄
	return [driver, 0]
}



; 이 스타일 전체 undisplyed 시키기
;~ makeThisStyleundisplyed(driver, DropStyle#, tab_title_of_Item_Search){
makeThisStyleundisplyed(driver, DropStyle#, URLofItemSearchPage){
	
;MsgBox, 이 스타일 전체 undisplyed 시키기
			
	;~ driver.SwitchToWindowByTitle(tab_title_of_Item_Search) ; tab_title_of_Item_Search 변수에 저장되어있는 탭 타이틀의 탭을 이용하기.
	driver.Get(URLofItemSearchPage) ; ITEM-Search 창으로 넘어가기
	
	; LAS 의 Item List 의 Search 칸에 아이템 번호 입력한 뒤 검색하기
	driver := putDropStyle#IntoTheSearchOnLAS(driver, DropStyle#)	
			
	; 해당 스타일의 체크박스 체크해주기
	Xpath = //span[text()='%DropStyle#%']//following-sibling::input[@type='checkbox']
	driver.FindElementByXPath(Xpath).click()
	Sleep 300

;MsgBox, undisplayed 클릭하기

	; undisplyed 버튼 클릭하기
	Xpath = //input[@value='Undisplay']
	driver.FindElementByXPath(Xpath).click()
	Sleep 300
	
	return driver
}


; Set Default Image 의 사진 다른 것으로 바꾸기
changeSetDefaultImage(driver){
	
;MsgBox, Set Default Image 의 사진 다른 것으로 바꾸기
				
	; Set Default Image 체크 박스 중에서 체크 안 된 것 찾아서 체크한 뒤 루프 빠져나가기
	loop{
					
		; Set Default Image 체크박스 주소
		Xpath := "(//input[@type='radio' and @name='default_img_file'])" . "[" . A_Index . "]"
					
;MsgBox, % A_Index . "번째 체크박스.`n체크박스 체크 됐으면 -1 안 됐으면 0 : " . driver.FindElementByXPath(Xpath).isSelected() . "`n`n`nisDisplayed() : " . driver.FindElementByXPath(Xpath).isDisplayed()
					
		; 체크가 안 된 체크박스 찾았으면 클릭해서 체크한 뒤 루프 빠져나가기
		if(!driver.FindElementByXPath(Xpath).isSelected()){
;MsgBox, % a_index . "번째 체크박스는 체크가 안 됐으니 원래 체크박스에서 이 체크박스로 옮기기 위해 이것 체크하기"
			driver.FindElementByXPath(Xpath).click()
			break
						
		} ; end끝 - if(!driver.FindElementByXPath(Xpath).isSelected())
	} ; end끝 - loop	
	
	
	return driver
}


; Save 버튼 클릭해서 저장하기
clickSaveButtonOnLAS(driver){
;MsgBox, To Save (LAS)

	Xpath = //input[@value='Save']
	driver.FindElementByXPath(Xpath).click()
	return driver
}






; ########################################################################################################################################################################################################################
; ############################################################################################# Magento 처리 #############################################################################################################
; ########################################################################################################################################################################################################################


	; Magento 의 아이템 삭제하는 창에 들어가기
openItemListOnMagento(driver, DropStyle#, StyleColor){
	
	
	; Magento 로그인 하는 메소드
	driver:= Magento_LoginOnly(driver)
	
	
	; Catalog - Manage Product 페이지로 이동
	URL = https://www.jodifl.com/index.php/admin/catalog_product/index/key/
	URL = https://www.jodifl.com/index.php/__admin/catalog_product/index/key/
	driver.Get(URL)
	Sleep 500



	; 아이템 삭제하는 창이 아니면
	if(!RegExMatch(driver.Url, "catalog_product")){

MsgBox , Magento 아이템 삭제하는 창이 아닙니다. 재귀 호출로 로그인 다시 합니다.

		openItemListOnMagento(driver, DropStyle#, StyleColor)
	}	

	return driver
}




; Name 에 아이템 번호 입력한 뒤 검색하기
putDropStyle#IntoTheSearchOnMagento(driver, DropStyle#){	
	
	; NAME 입력칸
	Xpath = //input[@name='name']
	driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(DropStyle#).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 DropStyle# 입력 후 엔터쳐서 로그인하기	
	
	driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다

	return driver
}



; 검색된 아이템 중 DropStyle# 와 StyleColor 가 둘 다 맞는 아이템의 체크박스 체크하기. 예를들어 P4030 INK BLACK 으로 검색했으면 P4030, P4030-1, P4030-2 중에서 P4030의 INK BLACK 만 딱 짚어서 해당 체크박스 체크하기
checkTheCheckBoxOfCorrectItemOnMagento(driver, DropStyle#, StyleColor, IsItRecursiveCall){

	Xpath = //td[contains(text(), '%DropStyle#%-%StyleColor%')]//ancestor::tr//descendant::input[@type='checkbox']
	
	; 드롭하려는 아이템 정보가 있는 Xpath 가 화면에 표시됐는지 확인
	if(driver.FindElementByXPath(Xpath).isDisplayed()){		

		driver.FindElementByXPath(Xpath).click()
		Sleep 300
		
;MsgBox, %DropStyle#% 와 %StyleColor% 가 맞는 아이템의 체크박스 찾아서 클릭했음


		; 체크된 체크박스 찾은뒤 그 체크박스에 해당되는 아이템 정보와 드롭하려는 아이템 정보 비교해서 맞으면 루프 빠져나가고 안 맞으면 재귀호출해서 체크 다시하기
		Loop{
			
			Xpath := "(//input[@type='checkbox' and @name='product'])" . "[" . A_Index . "]"			
			
			; 체크된 체크박스 찾기
			if(driver.FindElementByXPath(Xpath).isSelected()){
				
				; 체크된 체크박스에 해당되는 아이템 정보 얻기
				Xpath := "(//input[@type='checkbox' and @name='product'])" . "[" . A_Index . "]" . "//ancestor::tr//descendant::td[4]"
				innerTextOfCheckedOne := driver.FindElementByXPath(Xpath).Attribute("innerText")
				
;MsgBox, % A_Index . " 번째 체크박스가 체크되었음. `n체크된 체크박스에 해당하는 아이템번호와 색깔값은  :  " . innerTextOfCheckedOne

				; 드롭하려는 스타일번호와 색깔값
				dropStyle_Color = %DropStyle#%-%StyleColor%
				
				; 드롭하려는 스타일정보와 체크된 체크박스의 스타일 정보가 맞으면 루프 빠져나감
				if dropStyle_Color in %innerTextOfCheckedOne%
				{
					break
				}
				; 체크된 것과 드롭하려는 것이 맞지 않으면 현재 페이지 리프레쉬 한 뒤 재귀호출로 다시 체크하기
				else{
					
					; 만약 재귀호출확인을 위한 변수인 IsItRecursiveCall의 값이 1이면 이미 클릭된 체크박스가 잘못되어서 재귀호출을 한 번 했다는 뜻. 이것은 아예 처음부터 색깔이 잘못 입력됐을 가능성이 높다. 예를들어 NAVY MIX를 NAVY 로 입력한것 같은 경우
					if(IsItRecursiveCall){
						MsgBox, 262404, Drop Style ERROR, THE DROP STYLE IS NOT FOUNDED ON THIS MAGENTO PAGE. IF YOU'D LIKE TO CONTINUE THE PROCESSING, PLEASE CLICK 'YES' BUT IF YOU WANT TO RESTART THIS APP, PLEASE CLICK 'NO'`n`n지울 아이템 번호가 없습니다. 확인해 주시기 바랍니다. Yes 누르면 계속 진행하고 No 누르면 프로그램을 중지합니다.
					
						IfMsgBox, No
							Reload
					}
					
					driver.refresh()
					Sleep 500
					IsItRecursiveCall = 1 ; 재귀 호출을 했다는 표시로 1로 만듬
					checkTheCheckBoxOfCorrectItemOnMagento(driver, DropStyle#, StyleColor, IsItRecursiveCall)
				}


			} ; end끝 - if(driver.FindElementByXPath(Xpath).isSelected())
		} ; Loop 끝
		
		Sleep 300
	}
	; 현재 화면에 드롭하려는 아이템 정보가 없으면 에러 메세지 출력 후 프로그램 다시 시작하기
	else{
		MsgBox, 262404, Drop Style ERROR, THE DROP STYLE IS NOT FOUNDED ON THIS MAGENTO PAGE. IF YOU'D LIKE TO CONTINUE THE PROCESSING, PLEASE CLICK 'YES' BUT IF YOU WANT TO RESTART THIS APP, PLEASE CLICK 'NO'`n`n지울 아이템 번호가 없습니다. 확인해 주시기 바랍니다. Yes 누르면 계속 진행하고 No 누르면 프로그램을 중지합니다.
	
		IfMsgBox, No
			Reload		
	}
		
;MsgBox, 체크했나?


	return driver
}



; 위에서 체크된 아이템 상태 disable 로 만들기
disableTheStatusOfTheStyle(driver){
	
;MsgBox, TO DISABLE

	
	; Actions 드롭다운 박스 상태를 Change status 로 바꾸기
	Xpath = //select[@class='required-entry select absolute-advice local-validation']
	driver.FindElementByXPath(Xpath).SendKeys("Change status").sendKeys(driver.Keys.ENTER)
	
	Sleep 300
	
	; Status 드롭다운 박스 상태를 Disabled 로 바꾸기
	Xpath = //select[@class='required-entry absolute-advice select']
	driver.FindElementByXPath(Xpath).SendKeys("Disabled").sendKeys(driver.Keys.ENTER)
	
	Sleep 300
	
	; Submit 클릭하기
	Xpath = //*[text()='Submit']
	driver.FindElementByXPath(Xpath).click()
	
;MsgBox disabled 됐나?
	
	return driver
}



; 모든 색깔이 Drop 됐는지 체크한 뒤 체크됐으면 drop 하는 메소드
isAllColorsDorppedOnMagento(driver, DropStyle#, DropAll){


	; DropStyle# 값과 꼭 맞는 Xpath의 서수값을 저장하기 위한 배열 선언
	Arr_OrdinalNumbers := object()
	
	; 배열 초기화 해주기
	Arr_OrdinalNumbers := []
	
	
	

	; 화면에 보이는 체크 박스 갯수 세기
	; 체크 박스 갯수 만큼 가로줄이 있고 줄마다 스타일 정보가 구별되어 있으니까
	#ofCheckBoxesOnThisPage = 0
	Loop{
		
			Xpath := "(//input[@type='checkbox' and @name='product'])" . "[" . A_Index . "]"
		
		; 화면에 DropStyle# 에 해당하는 Xpath 가 나타날 때마다 #ofTheXpathsOnThisPage 변수값 증가시키기
		if(driver.FindElementByXPath(Xpath).isDisplayed()){
;MsgBox, % A_Index . " 번째" . "`n`n#ofTheXpathsOnThisPage : " . #ofTheXpathsOnThisPage . "`n`nXpath : " . Xpath

			#ofCheckBoxesOnThisPage++

		}
		else{
;MsgBox, 아이템 번호에 해당하는 tr Xpath 가 더이상 없음. 루프 끝내고 나감. NOT DISPLAYED
			break			
		}
	}
	
;MsgBox, % "이 화면의 체크박스 갯수는 : " . #ofCheckBoxesOnThisPage
	
	
	; 루프 횟수 제한키 위한 변수
	; 체크박스(=화면의 줄 갯수)의 두배만큼 반복하도록 변수 값 설정하기. 한 줄에 스타일과 관련된 값이 2 개씩 있기 때문에
	repeat#ToFindOrdinaNumbers := #ofCheckBoxesOnThisPage * 2


	; DropStyle# 와 정확히 일치하는 Xpath가 몇 번째에 있는지 알아내기 위한 loop
	; 체크박스(=화면의 줄 갯수)의 두배만큼 반복하기. 한 줄에 스타일과 관련된 값이 2 개씩 있기 때문에
	Loop, %repeat#ToFindOrdinaNumbers%{
		
		loop# := A_Index
		
		; DropStyle# 값이 포함된 Xpath 주소
		basic_Xpath = //td[contains(text(), '%DropStyle#%')]
		Xpath = 	((//td[contains(text(), '%DropStyle#%')])[%find#%]//following-sibling::td)[6]
		Xpath := "(" . basic_Xpath . ")" . "[" . A_Index . "]"		
		
		; DropStyle# 값이 포함된 Xpath의 innerText 값을 valueOnThePage 변수에 넣기
		valueOnThePage := driver.FindElementByXPath(Xpath).Attribute("innerText")
;		MsgBox, % Xpath . "`n`nvalueOnThePage : " . valueOnThePage
	
	
		; valueOnThePage 변수에 담긴 글자 갯수만큼 #ofLetter 에 넣기. 이 숫자만큼만 루프 돌기 위해. 
		; 글자 맨 끝부터 하나씩 읽다가 -가 나오면 중지 글자 갯수보다 더 많은 횟수의 루프를 돌 필요가 없으니까.
		#ofLetter := StrLen(valueOnThePage)

		; #ofLetter 에 담긴 valueOnThePage 의 글자 갯수만큼만 루프 돌기
		; 글자 맨 끝부터 하나씩 왼쪽으로 읽다가 -가 나오면 그것은 스타일 번호 다음의 색깔정보를 구분하기 위한 것.
		; 예를들어 P2471-5-TAUPE 에서 맨 오른쪽의 -다음에 색깔 정보인 TAUPE 가 나온다
		; 그렇기 때문에 -가 나오면 -뒤의 색깔 값을 지운 뒤 앞의 스타일 번호 값만 sTyle#WithOutColor_OnThisPage 변수에 저장하기
		Loop, % #ofLetter{
;			MsgBox, % "A_Index : " . A_Index

			; valueOnThePage 변수에 담긴 글자를 오른쪽부터 루프 회차만큼 읽어서 OutputVar 변수에 넣기
			StringRight, OutputVar, valueOnThePage, % A_Index
;			MsgBox, % OutputVar
			
			; 맨 마지막 문자가 숫자이면 가장 처음에 나오는 색깔이 없는 스타일 번호의 기본값(?) 이므로 continue 2 로 가장 밖의 루프로 건너뛰기
			if OutputVar is number
				continue 2
			
			; valueOnThePage 변수를 오른쪽부터 읽은 뒤 저장한 OutputVar 변수에 -가 발견되면 그것은 색깔을 구분하기 위한 - 가 발견됐다는 뜻
			; - 이후의 값 즉 색깔 값을 지운뒤 앞의 스타일 번호 값만 sTyle#WithOutColor_OnThisPage 변수에 넣기
			if OutputVar contains -
			{
;				MsgBox, % " 색깔을 구분하기 위한 - 발견됨."
				StringTrimRight, sTyle#WithOutColor_OnThisPage, valueOnThePage, % A_Index
				break
			}
			#ofLetter--
		}

;		MsgBox, % "sTyle#WithOutColor_OnThisPage : " . sTyle#WithOutColor_OnThisPage

		; DropStyle# 와 sTyle#WithOutColor_OnThisPage 가 같지 않으면 이것은 화면에 검색되어 나타나기는 했지만 필요 없는 스타일번호 값이라는 뜻
		if(DropStyle# != sTyle#WithOutColor_OnThisPage){
;			MsgBox, % "DropStyle# : " . DropStyle# . "`n`nsTyle#WithOutColor_OnThisPage" . sTyle#WithOutColor_OnThisPage . "`n이렇듯 서로 같지 않음"
		}

		; ; DropStyle# 와 sTyle#WithOutColor_OnThisPage 가 같으면 이 Xpath 의 서수값을 이용해야 되기 때문에 배열에 저장하기
		else if(DropStyle# == sTyle#WithOutColor_OnThisPage){
;MsgBox, % "DropStyle# : " . DropStyle# . "`n`nsTyle#WithOutColor_OnThisPage" . sTyle#WithOutColor_OnThisPage . "`n서로 같다. 즉 의미있는 값이다." . "`n`nDropStyle# 와 꼭 일치하는 서수 번호 : " . loop#
		
			; 홀수 값이면 loop# 값을 배열에 넣는다
			If (loop# & 1){  ; if last bit is 1
;MsgBox, % "Arr_OrdinalNumbers 배열에 값 넣는다 : " . loop#
				Arr_OrdinalNumbers.Insert(loop#) ; 배열에 값 넣기				
			} ; end끝 - If (Number & 1)

		}		

	
	} ; end - Loop 끝




	; Drop All 이 아닐때 
	; 화면의 Status 를 하나씩 확인하면서 한 개라도 Enabled가 있으면 itHasToBeActivated 변수값에 1을 넣어주고 
	; 전부 다 Disabled 되어 있으면 itHasToBeActivated 변수의 기본값인 0을 유지해서 아래의 if 문에서 전체 스타일 체크 후 Disable 시켜줄 수 있도록 만들기
	if(!DropAll){
		
		; Status 에 Enable 이 있는지 확인하기 위한 변수
		itHasToBeActivated = 0

		; DropStyle# 에 해당되는 Xpath 갯수가 2개씩 찾아지기 때문에 나누기 2 해준 횟수만큼만 루프 돌면서 해당되는 곳의 Status 값이 Disabled 인지 Enabled 인지 얻기
		ordinal#ofTheArray = 1  ; 몇 번째 배열 값을 볼 지 정하기 위해
;		Loop % Arr_OrdinalNumbers.Maxindex()/2{
		Loop % Arr_OrdinalNumbers.Maxindex(){
			
; ordinal#ofTheArray 에 있는 숫자의 위치에 있는 Arr_OrdinalNumbers 배열의 값을 find# 에 넣기 예를 들어 ordinal#ofTheArray 값이 1이면 Arr_OrdinalNumbers[1] 에 들어있는 값을 find# 에 넣기
;			find# := Arr_OrdinalNumbers[ordinal#ofTheArray]
			find# := Arr_OrdinalNumbers[A_Index]
			
		;	MsgBox % "Element number " . A_Index . " is " . Arr_OrdinalNumbers[ordinal#ofTheArray]
		;	MsgBox % "Element number " . A_Index . " is " . find#

			Xpath = ((//td[contains(text(), '%DropStyle#%')])[%find#%]//following-sibling::td)[6]
;			ordinal#ofTheArray := ordinal#ofTheArray + 2 ; 1,3,5,7 이런식으로 보기위해. 왜냐면 홀수번째와 짝수번째의 값이 똑같기 때문에 두 번 확인할 필요 없으니까
			
;MsgBox % "Xpath : " . Xpath . "`n`n" . driver.FindElementByXPath(Xpath).Attribute("innerText")

			; Enable이 한 개라도 있으면 itHasToBeActivated 변수값에 1 넣어줌
			if(driver.FindElementByXPath(Xpath).Attribute("innerText") == "Enabled"){
				itHasToBeActivated = 1
;MsgBox, % "이 화면에는 최소 한 개의 Enable이 있으므로 itHasToBeActivated 변수에 1을 넣어줬음.`nitHasToBeActivated : " . itHasToBeActivated . "`n`nStatus 값에 최소한 한 개 이상의 Enabled 가 있기 때문에 이 아이템은 Drop All 대상이 아님. 이후의 다른 Status 살필 필요 없이 루프 끝내고 나감"
				break

			} ; if끝 - if(driver.FindElementByXPath(Xpath).Attribute("innerText") == "Enabled")
/*				
			if(itHasToBeActivated == 1){
MsgBox, Status 값에 최소한 한 개 이상의 Enabled 가 있기 때문에 이 아이템은 Drop All 대상이 아님. 이후의 다른 Status 살필 필요 없이 루프 끝내고 나감
				break
				
			} ; if끝 - if(itHasToBeActivated == 1)
*/
		} ; end끝 - Loop % Arr_OrdinalNumbers.Maxindex()/2
	} ; end - if(!DropAll) 끝


	; 화면의 Status가 전부 다 Disabled 되어 있거나 Drop All 이 체크되어 있으면 
	; 해당 스타일의 체크박스를 전부 체크한 뒤 Disable 시키기
	if(itHasToBeActivated == 0 || DropAll){
;MsgBox, 모든 Status 값이 Disable 이거나 Drop All이 체크 됐으므로 이 아이템 전체가 Disabled 되어야 됨.
		
		#ofRepeatTheLoop := #ofTheXpathsOnThisPage / 2
;MsgBox, % "루프를 " . #ofRepeatTheLoop . "번 반복하면서 체크박스에 체크 할 예정"
		

		; Arr_OrdinalNumbers 배열 안에는 색깔은 없고 스타일 번호만 나와있는 첫줄(?)의 서수 번호는 없음. 
		; 첫째줄의 서수번호는 배열의 첫째 값에서 -2 해준 곳에 위치하기 때문에 find# 배열에 -2 해준 값을 넣어줌. 그래야 아래 루프에서 첫째줄에 해당하는 체크박스를 처음에 체크하고 진행함
		find# := Arr_OrdinalNumbers[1] - 2 ; 
;MsgBox, % "find# : " . find#

		; Arr_OrdinalNumbers 배열 안에는 색깔은 없고 스타일 번호만 나와있는 첫줄(?)의 서수 번호는 없음. 하지만 모든 스타일을 내리려면 첫째줄의 체크박스도 체크해줘야 되기 때문에 루프는 Arr_OrdinalNumbers 배열 최대값보다 한 번 더 돌면서 첫째줄의 체크박스도 체크해줌
		Loop % Arr_OrdinalNumbers.Maxindex()+1{
			
			Xpath = (//td[contains(text(), '%DropStyle#%')])[%find#%]//preceding-sibling::td//child::input[@type='checkbox']
			driver.FindElementByXPath(Xpath).click()
			
			find# := Arr_OrdinalNumbers[A_Index] ; 배열에 저장되어 있는 서수값을 find# 변수에 넣어준다

		} ; end끝 - Loop % Arr_OrdinalNumbers.Maxindex()+1
;MsgBox, disabled 될 체크박스들 모두 클릭됐나?

		; 위에서 체크된 아이템 상태 disable 로 만들기
		driver := disableTheStatusOfTheStyle(driver)
	}


	
	return driver
}
