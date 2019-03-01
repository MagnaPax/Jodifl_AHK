#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include CommWeb.ahk
#Include ChromeGet.ahk


class FG extends CommWeb{

	
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
				
			





			
		; 백오더 된 아이템은 체크박스가 활성화 되지 않는다
		; 현재 화면에 활성화 된 체크박스가 몇 개 있는지 알아보기 위해
		; 현재 화면에 있는 모든 활성화 된 체크박스 Str_#ofCheckBoxes 배열에 저장
		UnquotedOutputVar = imU)<div\s_ngcontent-c7=""\sclass="check-square"><div
				
		FoundPos = 1
		while(FoundPos := RegExMatch(HTML_Source, UnquotedOutputVar, SubPat, FoundPos + strLen(SubPat)))
		{
;			MsgBox, % SubPat1
			Str_#ofCheckBoxes.Insert(SubPat1)
		}
			
		; 소스코드에서 읽을 때는 아이템에 있는 체크박스 갯수보다 1개가 더 많다. 아마도 Total 옆에 있는 체크박스 갯수까지 세는 것 같다. 그래서 Str_#ofCheckBoxes 배열 갯수에서 1개를 빼준다
		#ofCheckBoxes := Str_#ofCheckBoxes.Maxindex() - 1
				
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
							



							; IsItCalledMiddleOfBOItemsProcessing 변수에 값이 있다면 백오더 처리중에 호출되었다는 뜻
							; 실제 아이템을 보내는 동작이 아닌 현재 페이지에는 pre order 날짜는 없어서 인스탁으로 나오지만 실제로는 물건이 없어서 지금 보내지 못하는 아이템들 백오더로 넘어줘야 하는 처리 실행
							if(IsItCalledMiddleOfBOItemsProcessing){
								

								; 실제로 보낼 아이템인데 체크 박스가 체크되어서 백오더로 넘어가게 됐다면 체크박스 클릭해서 체크된 것 해제하기
								; 체크박스가 체크되어 있지 않다면 그냥 넘어가기
								; driver.FindElementByXPath(AllMatchedItem_Xpath).isSelected() 통해서 체크박스가 체크되었는지 안 되었는지 확인한 뒤 처리하면 얼마나 좋겠냐만은 무엇때문인지 체크 유무에 상관없이 0값만 반환한다
								; 할 수 없이 아래의 방법을 이용한다
								
								; Arr_Items_Having_NoPreOrderDates 배열에는 이 메소드를 호출한 ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems 메소드를 통해 처리된, 현재 화면에서 pre order 날짜가 없는 Style No 값들 중에서 체크박스가 체크된 Style No 값이 있다
								; StyleNo 변수에는 실제로 보낼 아이템의 Style No 값이 들어있다.
								; 만약 StyleNo 변수 안에 들이었는 값이 Arr_Items_Having_NoPreOrderDates 배열에도 있다면 그건 ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems 메소드를 통해서 이미 체크됐다는 뜻이므로 해당 체크박스를 클릭해서 체크를 해제해줘야된다
								; 값이 들어있지 않다면 클릭하지 않고 넘어간다. 값이 들어있는데 클릭한다면 체크를 해제하는 게 아니라 새롭게 체크를 하게 되는 것이므로 현재 페이지에서 보내야 될 아이템이 백오더 화면으로 넘어가게 된다.							

								Loop % Arr_Items_Having_NoPreOrderDates.Maxindex(){
									
									MsgBox % "Element number " . A_Index . " is " . Arr_Items_Having_NoPreOrderDates[A_Index]
									
									StyleNoHasToBeUnCheckIfMatchedWith := Arr_Items_Having_NoPreOrderDates[A_Index]
									
									; 만약 StyleNo 값이 Arr_Items_Having_NoPreOrderDates 배열 안에 있다면
									if StyleNo in %StyleNoHasToBeUnCheckIfMatchedWith%
									;~ if Arr_Items_Having_NoPreOrderDates[A_Index] contains StyleNo
									{
										MsgBox, % "ITEM " . StyleNo . " HAS TO BE SHIPPED BUT IT HAS BEEN CHECKED SO TO RELEASE THE CHECKING" . "`n`n" . "실제로 보낼 아이템인데 체크가 됐기 때문에 체크를 해제합니다"
										driver.FindElementByXPath(AllMatchedItem_Xpath).click()
;										break
									}

								
								}
								*/
	MsgBox, out of the array
								
								; 브레이크 실행해서 실제 아이템을 보내기 위한 아래 동작들을 실행하지 않고 밖으로 나갑니다
								break
							}



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
		
	} ; End of Method CheckItemsInTheArrays
	
	
	






	
}




