﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ ; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)


#Include %A_ScriptDir%\lib\



;~ #Include LAMBS.ahk
;~ #Include CommLAMBS.ahk
;~ #Include COM.ahk
;~ #Include JO.ahk
;~ #Include N_E.ahk ; N41의 백오더 처리 모아둔 파일


#Include FindTextFunctionONLY.ahk
#Include FG.ahk
#Include N41.ahk
#Include CommN41.ahk
#Include LAS.ahk
#Include ChromeGet.ahk
#Include MethodsForPT.ahk ; PT 프로세싱을 위한 모든 클래스, 메소드를 모아놓은 파일



;~ #Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
;~ #Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
;~ #Include [Excel]_Joe Glines'sExcelFunctions.ahk
;~ #Include [Excel]_InsertORDeleteColumns.ahk






global #ofCC_counter ; 램스에서 읽은 카드 갯수 저장하는 변수a
global URL
global IsItFromNewOrder, IsItFromExcelFile, IsItFromAllocation, isThereItemsOnOpenSo, driver
global BO_by_Styles ; 스타일별로 뽑는 옵션인지 여러곳에서 알기 위해
BO_by_Styles = 0
;~ SO#BeingUsedNow, EndSO#, CustomerPO, IsItFromNewOrder, IsItFromExcelFile, SO#FromJODIFLcom, CustMemoFromJODIFLcom, IsItFromAllocation



L_driver := new LAMBS
N_driver := new N41
F_driver := New FG
LAS_driver := New LA
CommN41_driver := New CommN41
JO_driver := New JODIFL


FGFromAll_driver := New FG_ProcessingOfPT
LASFromAll_driver := New LAS_ProcessingOfPT


Arr_CSOS := object()
Arr_CC := object()
Arr_FGInfo := object()



		;~ N41_ProcessingForPT_driver := New N41_ProcessingForPT

		;~ N41_ProcessingForPT_driver.allocateClickedItems()
		
		
		
		;~ N41_ProcessingForPT_driver.Alloc_Print_WITH_PreAuthorized(CustomerPO)
		
		;~ MsgBox


;GUI Backgroud
;~ Gui, Show, w350 h150, N41 Processing, AlwaysOnTop Window
Gui, Show, w350 h200, N41 Processing, AlwaysOnTop Window
WinSet, AlwaysOnTop, On, N41 Processing

;Input Start Order Id
Gui, Add, Text, x22 y21 Cred , Start SO #
Gui, Add, Edit, x92 y19 w100 h20 vSO#BeingUsedNow,  ;53493 ;49998 ;49993

;Input End Order Id
Gui, Add, Text, x22 y51 CBlue , End SO #
Gui, Add, Edit, x92 y49 w100 h20 vEndSO#,  ;11, 22, 33

;PO NO.
Gui, Add, Text, x22 y85 CGreen , Cust PO #`n(For Urgent Order)
Gui, Add, Edit, x92 y79 w100 h20 vCustomerPO,  ; MTR1D55EDC764

;ORDER ID
;~ Gui, Add, Text, x22 y115 CBlack , Order ID Only`n(Skip Web Processing)
;~ Gui, Add, Edit, x92 y109 w100 h20 vOrder_ID_Only, ; 54250

;Check Box For JODIFL.COM
Gui, Add, CheckBox, x25 y145 w80 h40 vJODIFLcom, JODIFL.COM

;Check Box For BO
Gui, Add, CheckBox, x120 y145 w80 h40 vBO_by_Styles, BO by Style

;Allocation
Gui, Add, CheckBox, x220 y145 w80 h40 vAllocation, Allocation



;엔터 버튼
Gui, Add, Button, x225 y19 w100 h110 +default gClick_btn, Enter


;BO 버튼
;~ Gui, Add, Button, x25 y170 w100 h20 gBO_btn, BO



;GUI시작 시 포커스를 Allocation 라디오 박스에 위치
;~ GuiControl, Focus, CustomerPO
GuiControl, Focus, Allocation


return



Click_btn:

	WinActivate, ahk_exe nvlt.exe


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; To use the values which input on GUI
	Gui Submit, nohide
	GUI, Destroy

	;~ WinClose, ahk_class Chrome_WidgetWin_1
	
	
/*	
	;~ WinClose, ahk_class Chrome_WidgetWin_1 ; 모든 크롬창 닫기
	WinClose, FashionGo Vendor Admin - Google Chrome
	WinClose, LAShowroom.com Admin (JODIFL) -- Orders Editing Page - Google Chrome
	WinClose, LAShowroom.com Admin (JODIFL) -- Orders PO Search Page - Google Chrome
*/
	
	BuyerNotes := ""
	AdditionalInfo := ""
	StaffNotes := ""
	


	; Jodifl.com 오더 처리하기
	if(JODIFLcom){		
		
		; 오더 페이지에서 정보 얻을 배열 선언
		InfoOfJodiflcom := object()
		InfoOfJodiflcom := []
		
		J_Driver := New JODIFL
		
		loop{
			
			; 변수 초기화
			SO# := ""
			ShipmentMethod := ""
			CustMemo := ""
			
			; JODIFL.COM 오더페이지에서 정보 읽어오는 메소드 호출
			InfoOfJodiflcom := J_Driver.ProcessingJodiflcom()
			
			; 리턴된 배열에서 값 추출해서 각 변수에 넣기
			SO# := InfoOfJodiflcom[1]
			ShipmentMethod := InfoOfJodiflcom[2]
			CustMemo := InfoOfJodiflcom[3]

			; UPSG 가 아닌 오더 경고창 표시
			if(ShipmentMethod != "UPSG"){
				SoundPlay, %A_WinDir%\Media\Ring02.wav
				MsgBox, 262144, No UPSG Warnnig, IT'S NOT AN UPSG ORDER
			}

/*
			; 메모 내용을 읽어서 알파벳과 숫자를 제외한 모든 것을 "" 로 바꿔서(삭제) doesMemoExist 변수에 넣기
			doesMemoExist := RegExReplace(CustMemo, "[^a-zA-Z0-9]", "")

			; 메모가 있으면 보여주기
			; 알파벳과 숫자를 제외한 모든 것을 삭제했기 때문에 if문은 메모 내용이 있을때만 실행됨
			; 하지만 doesMemoExist 변수에는 공백문자도 삭제됐기 때문에 사람에게 보여주는 내용은 원래 메모 내용이 있는 CustMemo 변수의 내용을 보여줌
			IF(doesMemoExist){
				SoundPlay, %A_WinDir%\Media\Ring02.wav
				StringUpper, CustMemo, CustMemo ; 대문자로 바꾸기
				MsgBox, 262144, CUSTOMER MEMO, %CustMemo%
			}
*/			
;			MsgBox, % "SO#`n" . SO# . "`n`nShipmentMethod`n" . ShipmentMethod . "`n`nCustMemo`n" . CustMemo

			; N41 이 안 열렸으면 경고창 띄우기
			IfWinNotExist, ahk_class FNWND3126
				MsgBox, 262144, NO N41 WARNNING, PLEASE OPEN N41 THEN CONTINUE THE PROCESSING
			
			
			IsItFromNewOrder = 1 ; 뉴오더에서 호출하는 것이기 때문에
			IsItFromExcelFile = 0 ; BO 엑셀 파일에서 호출하는 것이 아니기 때문에
			SO#FromJODIFLcom = 1 ; Jodifl.com 처리에서 호출하는 것이기 때문에 값이 1
			BasicProcessing(SO#, EndSO#, CustomerPO, IsItFromNewOrder, IsItFromExcelFile, SO#FromJODIFLcom, CustMemo, IsItFromAllocation)
		}
	}
	
	

	; 백오더 처리
	; 스타일별로 정리된 엑셀 파일에서 읽어온 뒤 처리하기
	if(BO_by_Styles){
		IsItFromAllocation = 0
		N_BO_Driver := New BO
		N_BO_Driver.BOProcess_by_Style(IsItFromAllocation)
	}




	; Allocation 처리
	; Allocation 엑셀 파일에서 읽어온 뒤 처리하기
	if(Allocation){
		IsItFromAllocation = 1
		N_BO_Driver := New BO
		N_BO_Driver.BOProcess_by_Style(IsItFromAllocation)
	}




	; 시작번호 혹은 Customer PO 번호가 입력됐을때만 BasicProcessing 함수 호출해서 진행
	if(SO#BeingUsedNow || CustomerPO){
		
		IsItFromNewOrder = 1 ; 뉴오더에서 호출 했으니까
		
		if(CustomerPO)
			IsItFromNewOrder = 0 ; Customer PO 번호가 입력됐을때는 뉴오더에서 호출 안 했을 수도 있으니까
		
		IsItFromExcelFile = 0 ; BO 엑셀 파일에서 호출하는 것이 아니기 때문에
		SO#FromJODIFLcom = 0 ; Jodifl.com 처리에서 호출하는 것이 아니기 때문에
		BasicProcessing(SO#BeingUsedNow, EndSO#, CustomerPO, IsItFromNewOrder, IsItFromExcelFile, SO#FromJODIFLcom, CustMemoFromJODIFLcom, IsItFromAllocation)
	}
	else
		Reload












; 기본 처리 준 메인 함수
; SO#BeingUsedNow : 검색하기 위해 현재 사용되고 있는 Sales Order 번호
; CustomerPO : 
; IsItFromNewOrder : 뉴오더 처리를 위한 것인지 표시됨
; IsItFromExcelFile : BO 엑셀 파일에서 Sales Order 번호를 읽어온 것인지 표시됨
; SO#FromJODIFLcom : Jodifl.com 처리에서 Sales Order 번호를 읽어온 것인지 표시됨
; IsItFromAllocation : Allocation 처리인지 아닌지 표시. N_E파일의 BO 클래스에서 호출됨
BasicProcessing(SO#BeingUsedNow, EndSO#, CustomerPO, IsItFromNewOrder, IsItFromExcelFile, SO#FromJODIFLcom, CustMemoFromJODIFLcom, IsItFromAllocation){
	
;MsgBox, BasicProcessing 메소드에 들어왔나?

		; 사용자의 마우스 이동 막음
		BlockInput, MouseMove


		N_driver := new N41
		F_driver := New FG
		LAS_driver := New LA
		
;		IsItFromNewOrder = 1
			
		
		; EndSO# 끝번호 나올때까지 계속 루프돌기
		Loop
		{			
			; N41 활성화 시키기
			WinActivate, ahk_class FNWND3126

			Sleep 150
			
			; 혹시 이전에 열린 창이 있으면 닫기
			IfWinExist, Customer P.O or S.O # or Customer
				WinClose

			Sleep 150


			; Sales Order 클릭
			N_driver.ClickSalesOrderOnTheMenuBar()


			; Sales Order 에 있는 리프레쉬 버튼 클릭하기
			N_driver.ClickRefreshButtonOnSalesOrder()


			; Customer P.O or S.O # or Customer 창 뜰 때까지 기다리기
			WinWaitActive, Customer P.O or S.O # or Customer
			Sleep 150


			; CustomerPO 입력해서 실행시켰으면 검색창에 CustomerPO 입력해서 검색하기
			if(CustomerPO){
				Send, {Tab}
				Sleep 150
;				Send, {Tab}
;				Sleep 150
				SendInput, % CustomerPO
				Send, {Enter}
				Sleep 100
			}


			; 검색창에 SO#BeingUsedNow 입력하기
			; BOProcess_by_Style 메소드에서 호출됐으면 SO#BeingUsedNow 값이 엑셀 파일에서 얻은 SO#로 바뀌어서 호출됐음
			SendInput, % SO#BeingUsedNow
			Sleep 300
			Send, {Enter}
			Sleep 150

			Loop
			{					
				; Sale Order 에 있는 고객 코드 얻기
				CustCode := N_driver.GetCustomerCode()
				
;				MsgBox, % "CustCode : " . CustCode
				
				;값이 들어있으면 루프 빠져나가기
				;값이 들어있지 않으면 계속 루프 돌기
				if(CustCode)
					break
			}


			i = 0
			Loop
			{
				; Sale Order 에 있는 Customer PO 번호 얻기
				CustomerPO := N_driver.GetCustPONumber()
				i++
								
;				MsgBox, 262144, Title, CustomerPO : %CustomerPO%
				if(CustomerPO == "CREDIT")
					MsgBox, 262144, Title, 드디어 찾았다!!! 이 Customer PO 번호 저장해놓은 파일 참조해서 반드시 바꿔야 된다
								
				;값이 들어있으면 루프 빠져나가기
				;값이 들어있지 않으면 계속 루프 돌기
				if(CustomerPO){
					break
				}
				; 쇼 오더는 Customer PO 값이 없을수도 있다. 그럴땐 반복해서 2번 확인해보기
				if(i == 2)
					break
			}			


			; Sales Order 에서 Order Type 얻기
			orderType := N_driver.getOrderType()
			
			
			; 뉴오더일때만 할인이 있는지 확인 후 있으면 Discount Rate 복붙하기
			if(IsItFromNewOrder)
				N_driver.copyDiscountRateAndPasteIfExist()
			



			; 뉴오더일때만 Customer PO에 맞는 PMT Method로 바꾸기
			if(IsItFromNewOrder){
;MsgBox, 262144, Title, CustomerPO : %CustomerPO%
				if CustomerPO contains MTR
				{
;MsgBox, 262144, Title, FG 주문임
					N_driver.changePMTMethodToFGorLAS(CustomerPO)	; PMT Method 를 FG-CC 나 LAS-CC 로 바꾸기
;MsgBox, 262144, Title, FG-CC 로 바뀌었나?
				}
				else if CustomerPO contains OP
				{
;MsgBox, 262144, Title, LAS 주문임
					N_driver.changePMTMethodToFGorLAS(CustomerPO)	; PMT Method 를 FG-CC 나 LAS-CC 로 바꾸기
;MsgBox, 262144, Title, LAS-CC 로 바뀌었나?					
				}
			}


			; Sale Order 에 있는 Memo 얻기
			Clipboard := ""
			CustMemoOnSOTab := ""
			;~ CustMemoOnSOTab := N_driver.ReachMemoOnSalesOrderTab()
			
			CustMemoOnSOTab := N_driver.MemoOnSalesOrderTab()
			
			HouseMemoOnSOTab := N_driver.HouseMemoOnSalesOrderTab()
			

			
			
			
;			MsgBox, % "CustMemoOnSOTab`n`n`n" . CustMemoOnSOTab
			
			
			
			
			; 메모 내용을 읽어서 알파벳과 숫자를 제외한 모든 것을 "" 로 바꾼뒤(삭제) doesMemoExist 변수에 넣기
			doesMemoExist := RegExReplace(CustMemoOnSOTab, "[^a-zA-Z0-9]", "")

			; 메모 내용을 읽어서 알파벳과 숫자를 제외한 모든 것을 "" 로 바꾼뒤(삭제) doesHouseMemoExist 변수에 넣기
			doesHouseMemoExist := RegExReplace(HouseMemoOnSOTab, "[^a-zA-Z0-9]", "")


			
			; Sales Order 페이지에 메모 내용이 있었으면 보여주기
			; 알파벳과 숫자를 제외한 모든 것을 삭제했기 때문에 if문은 메모 내용이 있을때만 실행됨
			; 하지만 doesMemoExist 변수에는 공백문자도 삭제됐기 때문에 사람에게 보여주는 내용은 원래 메모 내용이 있는 CustMemoOnSOTab 변수의 내용을 보여줌
			
/*			
			; 하지만 쇼 오더는 SO 작성한 사람이 자기 이름 넣거나 Billing Addr 넣는등 크게 중요치 않은 정보가 너무 많아서 여기서 표시하면 너무 성가시다
			; 쇼 오더의 Customer Memo 내용은 일단 여기서는 표시하지 말고 픽티켓 뽑기 직전에 한 번만 보여주자
			;~ if CustomerPO not contains 2000
			if SO#BeingUsedNow not contains 200
*/



			; Sales Order 탭에서 메모 여는건 그냥 뉴오더일때만 하자
			if(IsItFromNewOrder)
			{
				if(doesMemoExist){
					
					; 사용자의 마우스 이동 허용
					BlockInput, MouseMoveOff
					
					StringUpper, CustMemoOnSOTab, CustMemoOnSOTab ; 대문자로 바꾸기
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, Customer Memo, %CustMemoOnSOTab%
					
					; 사용자의 마우스 이동 막음
					BlockInput, MouseMove					
					
				}				
			}
			
			; House Memo 가 있으면 띄우기
			if(doesHouseMemoExist){
					
				; 사용자의 마우스 이동 허용
				BlockInput, MouseMoveOff
					
				StringUpper, HouseMemoOnSOTab, HouseMemoOnSOTab ; 대문자로 바꾸기
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 262144, House Memo, %HouseMemoOnSOTab%
					
				; 사용자의 마우스 이동 막음
				BlockInput, MouseMove					
					
			}				



			
			; 왼쪽 메뉴바에 있는 Customer 클릭하기
			N_driver.ClickCustomerMarkOnTheLeftBar()			
			
			
			; 왼쪽 메뉴바에 있는 SO Manager 클릭하기
			N_driver.ClickSOManagerOnTheLeftBar()
			
			
			; SO Manager 에 있는 Customer 표시 찾기. 고객 코드 입력하기 위함
			N_driver.FindCustomerMarkToFillInTheBlank()
			
			
			; 검색창에 고객 코드 입력 후 엔터쳐서 검색하기		
			Sleep 300
			Send, ^a
			Sleep 150
			Send, % CustCode
			Sleep 150
			SendInput, {Enter}
			Sleep 150
			
			
			Sleep 3000
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}			
			while (A_cursor = "Wait")
				Sleep 1000




			; #############################################################################################
			; ################ 왼쪽은 6번,  오른쪽은 9번, 오른쪽 아래는 9번 확대해줘야 작동함 #################
			; #############################################################################################



			; ################ 이번 주문 처리 않고 다음 주문으로 넘기기 ################			

			; 마지막 배송일 읽은뒤 만약 최근에 내보냈다면
			; 뉴오더일때는 경고창 띄우고 백오더일때는 그냥 제끼기						
			Sleep 1500			
			isItSupposedToBeSkipped := isItRecentShippedOutOrder() ; 마지막 배송일 읽기
			
			; 뉴오더가 아닐때는 최근 주문이면 그냥 제끼기
			if(!IsItFromNewOrder){
				
				; isItSupposedToBeSkipped 변수 값이 0일때는 마지막 배송일이 최근이라는 뜻이니 제끼기
				if(!isItSupposedToBeSkipped){
;					MsgBox, 262144, Title, 이 오더는 배송된지 얼마 안 됐으니 제끼기
					return
				}
			}
			; 뉴오더일때는 마지막 배송날짜가 최근이면 경고창 띄우기
			else if(IsItFromNewOrder){
				
				; isItSupposedToBeSkipped 변수 값이 0일때는 마지막 배송일이 최근이라는 뜻이니 경고창 띄우기
				if(!isItSupposedToBeSkipped){
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					;~ MsgBox, 262144, RECENT SHIP OUT ORDER, 이 오더는 배송된지 얼마 안 된 주문.
					MsgBox, 262144, RECENT SHIP OUT ORDER, THE LAST INVOICE DATE OF THIS CUSTOMER WAS RECENTLY, LESS THAN 2 BUSINESS DAYS.`n`n마지맥 배송일이 2일 이내
					
				}			
			}


			; 펜딩 오더가 있는지 확인키 위한 변수.
			isTherePendingOrder = 0
			
			; 펜딩 오더가 있는지 확인하기
			; 펜딩 오더 있으면 isTherePendingOrder 값은 0
			; 펜딩 오더 없으면 isTherePendingOrder 값은 1
			isTherePendingOrder := N_driver.checkPickTicketSectionToFindIfPendingOrderExists()
			
			
			; 뉴오더가 아닐때는 펜딩된 픽티켓이 있다면 그냥 제끼기
			if(!IsItFromNewOrder){
				
				; 펜딩된 Pick Ticket 이 있다면 이번 주문 제끼기
				if(!isTherePendingOrder){
					return
				}
			}

			
			; TIFFANY ONLY 주문이면 그냥 다음 주문으로 자동으로 넘어가기
			; 뉴오더가 아닐때 실행. 뉴오더는 정보를 업데이트해야되는 경우가 많기 때문에
			if(!IsItFromNewOrder){
				
				; SO Manager 화면에서 priority 번호 얻기
				priority# := N_driver.getPriorityOnSOManager()
				
				; priority 번호가 1번이면 더이상 진행하지 않고 리턴해서 다음 주문으로 넘어가기
				if(priority# == 1){
					return
				}
			}
			
			
			; po box 주소인지 확인			
			addrs := N_driver.getADDr()

			; PO BOX 주소면 SHIP VIA 바꾸도록 안내
			if addrs contains box
			{
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 262144, PO BOX, IT'S SHIP TO PO BOX. SHIP VIA HAS TO BE CHANGED TO FEDEX.
			}
			
			
			; Open SO 에 아이템이 있는지 확인하기
			; 수동으로 처리하더라도 추가할 아이템이 있는지 확인하기 filter를 사용해서 쉽게 나타내주면 좋으니까
			; isThereItemsOnOpenSo 변수는 global 로 선언되어서 아무곳에서나 사용 가능함
			isThereItemsOnOpenSo = 0
			;~ isThereItemsOnOpenSo := N_driver.checkOpenSoIfThereAreItemsShipOut()
			
			
			

			
			; ################ Allocation 주문 중에서 자동으로 뽑기 ################



			; 사용자의 마우스 이동 허용
			BlockInput, MouseMoveOff



;~ /*
; #############################################################
; 2019년 10월 달라스 쇼 오더를 주의할 필요가 없게 되면 비활성화 시키기
if(orderType == "10DAL19"){
	SoundPlay, %A_WinDir%\Media\Ring02.wav
	MsgBox, 262160, Title, 이거 2019년 10월 달라스 쇼. 픽티켓 만들기 전 메모 확인
}
*/
/*
; 2019년 8월 매직 쇼 오더를 주의할 필요가 없게 되면 비활성화 시키기
if(orderType == "08MAG19"){
	SoundPlay, %A_WinDir%\Media\Ring02.wav
	MsgBox, 262160, Title, 이거 2019년 8월 매직 쇼. 픽티켓 만들기 전 메모 확인
}
*/
;~ /*
; 2019년 10월 아틀란타 쇼 오더를 주의할 필요가 없게 되면 비활성화 시키기
if(orderType == "10ATL19"){
	SoundPlay, %A_WinDir%\Media\Ring02.wav
	MsgBox, 262160, Title, 이거 2019년 10월 아틀란타 쇼. 픽티켓 만들기 전 메모 확인
}
*/
/*
if SO#BeingUsedNow contains 3000
{
    MsgBox, 262160, Title, 이거 2019년 8월 아틀란타 쇼. 픽티켓 만들기 전 메모 확인
}
; ############################################################
*/


/*
; ######################################################################################################################################################
if(orderType != "06DAL2019"){ ; Order Type 이 2019년 6월의 달라스 쇼가 아닌것을 자동뽑기 확인하기. 정보 일일히 수동으로 확인해야 됨
	if(orderType != "06ATL18"){ ; Order Type 이 2018년 6월의 아틀란타 쇼가 아닌것을 자동뽑기 확인하기. 이 오더는 여름에는 겨울 옷 보내면 안됨. 시간이 지나면 이거 지우기
; ######################################################################################################################################################
*/

;MsgBox, % "isItSupposedToBeSkipped : " . isItSupposedToBeSkipped

		; 변수 값이 2일때만 자동으로 픽티켓 뽑을지 말지 진행하기
		if(isItSupposedToBeSkipped == 2){
			
		
			; Allocation 주문일때만 
			if(IsItFromAllocation){
				
				
				; 펜딩 오더가 있는지 확인키 위한 변수.
				;~ isTherePendingOrder = 0
				
				
				; priority 번호가 2인 경우만
				if(priority# == 2){
					
					;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
					;~ MsgBox, 262144, priority # 2, priority# : %priority#% `n`n Click Ok To continue`n`npriority 번호가 2. 안전을 위해 잠시 멈춤. 계속하려면 Ok 클릭
						


					; 펜딩 오더가 있는지 확인하기
					; 펜딩 오더 있으면 isTherePendingOrder 값은 0
					; 펜딩 오더 없으면 isTherePendingOrder 값은 1
					;~ isTherePendingOrder := N_driver.checkPickTicketSectionToFindIfPendingOrderExists()
					
					
					
					; Sales Order 페이지에 메모 내용이 있었으면 다시 보여주기. 고객 메모나 하우스 메모 둘 중에 하나만 있어도 보여주기
					; 알파벳과 숫자를 제외한 모든 것을 삭제했기 때문에 if문은 메모 내용이 있을때만 실행됨
					; 하지만 doesMemoExist 변수에는 공백문자도 삭제됐기 때문에 사람에게 보여주는 내용은 원래 메모 내용이 있는 CustMemoOnSOTab 변수의 내용을 보여줌
					if(doesMemoExist || doesHouseMemoExist){
						
;MsgBox, 262144, Memo, 메모 내용이 있어서 소리발생 후 메세지 나올 예정
						SoundPlay, %A_WinDir%\Media\Ring06.wav
						
						; 고객 메모가 있으면 고객 메모만 보여주고
						if(doesMemoExist){
							MsgBox, 262144, Memo, %CustMemoOnSOTab%
						}
						; 하우스 메모가 있으면 고객 메모와 하우스 메모 같이 보여주기
						if(doesHouseMemoExist){
							MsgBox, 262144, Memo, Customer Memo`n`n%CustMemoOnSOTab%`n`n`n`nHouse Memo`n`n%HouseMemoOnSOTab%			
						}				
					}					
					
					; 뉴오더가 아닐때 주문화면 한 번 확인하게끔 환기시키기 위해 효과음 발생
					; 뉴오더일때는 어차피 주문마다 정지하기 때문에 여기서 효과음 들려줄 필요 없음
					; 컴퓨터 2대를 사용해서 어차피 매번 멈추기 때문에 이것 사용할 필요가 없어서 비활성화 시키려고했는데 음...
					else if(!IsItFromNewOrder){
						SoundPlay, %A_WinDir%\Media\Ring02.wav			
					}
					

					; 컴퓨터 2대를 사용하기 때문에 priority 번호가 2라도 한 번 멈추고 확인케 하기 위해
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, priority # 2, priority# : %priority#% `n`n Click Ok To continue`n`npriority 번호가 2. 안전을 위해 잠시 멈춤. 계속하려면 Ok 클릭														


					
					
						
					
					; 펜딩된 Pick Ticket 이 없을때만 (isTherePendingOrder 변수에 값이 있을때만)
					if(isTherePendingOrder){							
							
						; Open SO 에 아이템이 있는지 확인하기
						isThereItemsOnOpenSo := N_driver.checkOpenSoIfThereAreItemsShipOut()
							
							
						; 아이템이 없는 경우에만 (아이템이 있으면 추가해서 뽑아야 하기 때문에)
						if(isThereItemsOnOpenSo){
							
;							MsgBox, 아이템이 없습니다(자동으로 픽티켓 뽑아도 됩니다)
							
							; 인쇄 위한 객체 생성
							N41_ProcessingForPT_driver := New N41_ProcessingForPT


							; 카드 정보 읽기
;							ccInfo := N_driver.checkCC()
							ccInfo = 0
								
							; 카드가 있으면
							if(ccInfo){
								
								; Sales Order 탭으로 가서 PMT Method 를 Credit Card로 바꾼뒤 SO Manager 탭으로 다시 돌아오기
								
								; pre-authorized 받고 프린트하기
								N41_ProcessingForPT_driver.Alloc_Print_WITH_PreAuthorized(CustomerPO)
								
							}
								
							; 카드가 없으면
							if(!ccInfo){
								; CBS 등 돈 받지 말고 프린트 하기
								N41_ProcessingForPT_driver.Alloc_Print_WITHOUT_PreAuthorized(CustomerPO)
								
							}						


							
							; 픽티켓도 뽑았으니 계속 진행하지 않고 여기서 리턴하기
							return
							
						}

					}
;					else
;						MsgBox, 펜딩오더 있음(자동으로 픽티켓 뽑지 않음)
				}
			} ; 끝end - if(IsItFromAllocation)
		}

/*
; ###################################################################
	} ; 2019년 7월 넘어서 아틀란타 쇼인지 아닌지 비교할 필요가 없으면 이거 비활성화 하기
} ; 2019년 7월 넘어서 달라스 쇼에서 온 주문 정보를 수동으로 수정할 필요가 없어지면 이거 비활성화 하기
; ###################################################################
*/




			; 뉴오더일때는 펜딩된 픽티켓이 있다면 경고창 띄우기
			; 펜딩 오더가 있는지 없는지 확인하는 checkPickTicketSectionToFindIfPendingOrderExists 함수는 위에서 호출해서 이미 isTherePendingOrder 변수에 결과값 저장했다
			if(IsItFromNewOrder){
				
				; 펜딩된 Pick Ticket 이 있다면 경고창 띄우기
				if(!isTherePendingOrder){
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 262144, WARNNING, !!!! WARNNING !!!!`n`nCHECK PENDING ORDERS
				}								
			}
			
			
			; 필요 없는 문자가 들어있을 경우를 대비해 CustomerPO 값 정리해주기
			CustomerPO := Trim(CustomerPO)
			CustomerPO := RegExReplace(CustomerPO, "[^a-zA-Z0-9]", "")

			
;			IsItFromNewOrder = 1
			
			; BO 처리하는 엑셀에서 호출된 것이 아닐때 실행
			; CustomerPO 가 FG 일때 FG 오더 처리
			if(!IsItFromExcelFile){
;MsgBox, 이제 이게 실행되나?
					
				if(RegExMatch(CustomerPO, "imU)MTR")){
					FGFromAll_driver := New FG_ProcessingOfPT
					FGFromAll_driver.ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile)
					;~ ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile)
				}			
				; CustomerPO 가 LAS 일때
				else if(RegExMatch(CustomerPO, "imU)OP")){
					LASFromAll_driver := New LAS_ProcessingOfPT
					LASFromAll_driver.ProcessingLASOrder(CustomerPO, LAS_driver, N_driver)
					;~ ProcessingLASOrder(CustomerPO, LAS_driver, N_driver)
				}			

				;~ MsgBox, out of FG processing method
			}
			


			
			; BO 처리하는 엑셀에서 호출된 것일때 실행
			; 메세지 나타내기 전에 노래 나오게
			if(IsItFromExcelFile){
				SoundPlay, %A_WinDir%\Media\Ring06.wav
			}
			
			
			
			; Sales Order 페이지에 메모 내용이 있었으면 다시 보여주기. 고객 메모나 하우스 메모 둘 중에 하나만 있어도 보여주기
			; 알파벳과 숫자를 제외한 모든 것을 삭제했기 때문에 if문은 메모 내용이 있을때만 실행됨
			; 하지만 doesMemoExist 변수에는 공백문자도 삭제됐기 때문에 사람에게 보여주는 내용은 원래 메모 내용이 있는 CustMemoOnSOTab 변수의 내용을 보여줌
			if(doesMemoExist || doesHouseMemoExist){
				
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				
				; 고객 메모가 있으면 고객 메모만 보여주고
				if(doesMemoExist){
					MsgBox, 262144, Memo, %CustMemoOnSOTab%
				}
				; 하우스 메모가 있으면 고객 메모와 하우스 메모 같이 보여주기
				if(doesHouseMemoExist){
					MsgBox, 262144, Memo, Customer Memo`n`n%CustMemoOnSOTab%`n`n`n`nHouse Memo`n`n%HouseMemoOnSOTab%			
				}				
			}


; ####################################################################################################################################################################################################
; ####################################################################################################################################################################################################
; ####################################################################################################################################################################################################
; ####################################################################################################################################################################################################
; ####################################################################################################################################################################################################

/*
			; Pick Ticket 뽑기
			; Allocation 처리가 아닐 때 실행
			; 1사분면에 있는 Pick Ticket 버튼을 누른 뒤 인쇄 진행
			if(!IsItFromAllocation){

				; Yes 는 pre-authorized 받고 인쇄
				; No 는 CBS등 돈 받지 않고 인쇄
				; Cancel 는 인쇄하지 말고 다음으로 넘어가기
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 266243, OPTIONS, Yes : PRINT P.T. WITH pre-authorize BY N41`n`nNo : PRINT P.T. WITH no pre-authorize BY N41`n`nCancel : NO PRINTING`n`n%CustMemoFromJODIFLcom%
				
				
				Print_WITH_PreAuthorized_AfterAllocateItems()
				
				N41_ProcessingForPT_driver := New N41_ProcessingForPT
				
				; pre-authorized 받고 프린트하기
				IfMsgBox, Yes
				{
					N41_ProcessingForPT_driver.OpenSO_Print_WITH_PreAuthorized(CustomerPO, IsItFromNewOrder)
				}
				; CBS 등 돈 받지 말고 프린트 하기
				IfMsgBox, No
				{
					; 스타일별로 뽑을때는 뉴오더가 아니라고 표시해서 무조건 주문창 열게끔
					if(BO_by_Styles){
						IsItFromNewOrder = 0
						
						; pre-authorized 안 받고 프린트하기
						N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)
;			MsgBox, 262144, Title, <MAIN> 스타일별로 뽑았음. 이 메세지 창 닫으면 리턴함. 여기서 에러나나 **1**
						return
					}
						
					; 스타일별로 뽑을때가 아닐때는 뉴오더라고 표시해보자
					else{						
						IsItFromNewOrder = 1
						; pre-authorized 안 받고 프린트하기
						N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)
					}

					; pre-authorized 안 받고 프린트하기
					;~ N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)
					


				}

			}
*/

;~ /*
			; Pick Ticket 뽑기
			; Allocation 처리가 아닐 때 실행
			; 1사분면에 있는 Pick Ticket 버튼을 누른 뒤 인쇄 진행
			if(!IsItFromAllocation){

				; Yes 는 pre-authorized 받고 인쇄
				; No 는 CBS등 돈 받지 않고 인쇄
				; Cancel 는 인쇄하지 말고 다음으로 넘어가기
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 266243, OPTIONS, Yes : PRINT P.T. WITH pre-authorize BY N41`n`nNo : PRINT P.T. WITH no pre-authorize BY N41`n`nCancel : NO PRINTING`n`n%CustMemoFromJODIFLcom%				
				
				N41_ProcessingForPT_driver := New N41_ProcessingForPT
				
				; pre-authorized 받고 프린트하기
				IfMsgBox, Yes
				{
					if(BO_by_Styles){
N41_ProcessingForPT_driver.OpenSO_Print_WITH_PreAuthorized(CustomerPO, IsItFromNewOrder)
;~ MsgBox, 262144, Title, 스타일 별로 pre-authorized 받고 프린트 하려고 함
;~ allocateClickedItems()

/*
WOULD YOU LIKE TO CHANGE PMT Method of Sales Order AND LEAVE MESSAGE THAT 'FG PA' or 'LAS PA' ON THE HOUSE MEMO?
*/
					}
					
					; 일반적인 뉴오더일때는 allocate 하지 않고 프린트
					else
						N41_ProcessingForPT_driver.OpenSO_Print_WITH_PreAuthorized(CustomerPO, IsItFromNewOrder)
					
					
				}
				; CBS 등 돈 받지 말고 프린트 하기
				IfMsgBox, No
				{
					; 스타일별로 뽑을때는 뉴오더가 아니라고 표시해서 무조건 주문창 열게끔
					if(BO_by_Styles){
						IsItFromNewOrder = 0
						
						; pre-authorized 안 받고 프린트하기
;~ MsgBox, 262144, Title, 스타일 별로 pre-authorized '안' 받고 프린트 하려고 함						
						N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)
;			MsgBox, 262144, Title, <MAIN> 스타일별로 뽑았음. 이 메세지 창 닫으면 리턴함. 여기서 에러나나 **1**
						return
					}
						
					; 스타일별로 뽑을때가 아닐때는 뉴오더라고 표시해보자
					else{						
						IsItFromNewOrder = 1
						; pre-authorized 안 받고 프린트하기
						N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)
					}

					; pre-authorized 안 받고 프린트하기
					;~ N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)
					


				}

			}			
*/			
			
			
			; Pick Ticket 뽑기
			; Allocation 처리일때 실행. 3사분면에 있는 Pick Ticket 버튼을 누른 뒤 인쇄 진행
			if(IsItFromAllocation){
				MsgBox, 266243, OPTIONS, Yes : PRINT P.T. WITH pre-authorize BY N41`n`nNo : PRINT P.T. WITH no pre-authorize BY N41`n`nCancel : NO PRINTING`n`n%CustMemoFromJODIFLcom%
				
				N41_ProcessingForPT_driver := New N41_ProcessingForPT
				
				; pre-authorized 받고 프린트하기
				IfMsgBox, Yes
				{
					N41_ProcessingForPT_driver.Alloc_Print_WITH_PreAuthorized(CustomerPO)
					
					;~ PrintByPayment(IsItFromAllocation)
				}
				; CBS 등 돈 받지 말고 프린트 하기
				IfMsgBox, No
				{
					N41_ProcessingForPT_driver.Alloc_Print_WITHOUT_PreAuthorized(CustomerPO)
	;				MsgBox, 262144, Title, WAIT TO CONTINUE NEXT ORDER
				}

			}
			
			
driver.quit()
			
			
			; Allocation 에서 불려온거면 여기서 그냥 리턴하기
			if(IsItFromAllocation){
				;~ MsgBox, 262144, Title, 뽑았으니까 리턴해서 끝내자. BOProcess_by_Style 함수로 다시 돌아가기
				return
			}
			
	

			
			; BO 처리하는 엑셀에서 호출된 것일때 실행
			; 아래 처리하지 않고 호출한 BOProcess_by_Style 메소드로 돌아가기
			if(IsItFromExcelFile){
				return
			}
			
			
			
			; Jodifl.com 처리하는 메소드에서 호출할 때 실행
			; 아래 처리하지 않고 호출한 곳으로 돌아가기
			if(SO#FromJODIFLcom){
				return
			}
			
			
			

			; SO 번호를 증가시키기
			if(SO#BeingUsedNow){
				SO#BeingUsedNow++
			}
			
			
			; 시작 SO 번호와 끝 SO 번호가 같으면 중단하고 프로그램 다시 시작하기
			if(SO#BeingUsedNow == EndSO#+1){
				MsgBox, 262144, Title, IT'S END OF SO NUMBER, RELOAD THE PROGRAM
				Reload
			}
			
			; SO#BeingUsedNow 에 값이 있으면 변수값 초기화 해준 뒤 루프 돌기
			if(SO#BeingUsedNow){

				; 변수값 초기화
				Clipboard := ""
				CustCode := ""
				CustomerPO := ""
				
				continue
			}
			
			
			; CustomerPO 입력해서 실행시켰으면 루프 돌 필요 없이 리로드 해주기
			if(CustomerPO){
				WinClose, FashionGo Vendor Admin - Google Chrome
				Reload
			}
			
			

		}

	
	return	
	
} ; BasicProcessing 메소드 끝

































; 마지막 배송 날짜가 너무 가까운 날짜인지 아닌지 알아보는 메소드
isItRecentShippedOutOrder(){

	
	; 마지막 배송 날짜(Last Inv. Dt) 읽어오기
	N_driver := new N41

	i = 0
	Loop
	{					
		; SO Manager 에 있는 Last Inv. Date 얻기
		lastInvDate := N_driver.getLastInvDateOnSOManager()
		i++
					
;		MsgBox, % "lastInvDate : " . lastInvDate

		;값이 들어있으면 루프 빠져나가기
		;값이 들어있지 않으면 계속 루프 돌기
		if(lastInvDate){
			break
		}
		; 처음 배송되는 주문은 lastInvDate 값이 없을수도 있다. 그럴땐 반복해서 2번 확인해보기
		if(i == 2)
			break
	}
	
	; 마지막 배송 날짜가 없으면 제끼지 않고 자동으로 처리할 지 살펴보기 위해 2 리턴하기
	if(!lastInvDate){
		return 2
	}
					
	

	; 마지막 배송 날짜 읽어오기
	; #### 실제로는 N41에서 읽어와야 됨 #########
	;~ lastInvDate = 12/30/2018 00:00:00
	;~ lastInvDate = 05/16/2018 00:00:00
	;~ lastInvDate := ""


	; 알파벳과 숫자만 저장 (알페벳과 숫자 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
	lastInvDate := RegExReplace(lastInvDate, "[^a-zA-Z0-9]", "")


	; 오늘 날짜
	todaysDate = %A_MM%%A_DD%%A_YYYY%  ; ############ 실전에는 이걸 써야됨 ################
	;~ todaysDate = 05212018


	; lastInvDate 변수에서 뒤에 있는 시간(00:00:00 이렇게 표시됨)을 제외한 년월일만 뽑아서 변수에 다시 넣기
	StringLeft, lastInvDate, lastInvDate, 8


;	MsgBox, % "마지막 배송 날짜 : "lastInvDate . "`n" . "            오늘 날짜 : " . todaysDate






	; 마지막 배송날짜가 없다면 수동으로 처리하기 위해 0을 리턴
	IF(lastInvDate == "")
	{
;		MsgBox, 마지막으로 배송된 날짜가 없음
		return 0
	}


	;~ StringLeft, lastInvDate, lastInvDate, 8


	; 연도를 변수에 저장
	yearOfLastInvDate := SubStr(lastInvDate, 5, 4)
	yearOfToday := SubStr(todaysDate, 5, 4)

	dateOfLastInvDate := SubStr(lastInvDate, 3, 2)
	dateOfToday := SubStr(todaysDate, 3, 2)

	monthOfLastInvDate := SubStr(lastInvDate, 1, 2)
	monthOfToday := SubStr(todaysDate, 1, 2)

	
	; 오늘 요일 변수에 넣기
	; 주일이면 1 월요일이면 2 화요일이면 3 수요일이면 4 목요일이면 5 금요일이면 6
	whatsTheDayToday := A_WDay    ; ########### 실전에서는 이것 써야됨 ####################


;	MsgBox, % "yearOfLastInvDate : " . yearOfLastInvDate . "`n" . "yearOfToday : " . yearOfToday . "`n`n" . "dateOfLastInvDate : " . dateOfLastInvDate . "`n" . "dateOfToday : " . dateOfToday . "`n`n" . "monthOfLastInvDate : " . monthOfLastInvDate . "`n" . "monthOfToday : " . monthOfToday


/*
	; 마지막 배송 날짜의 연도와 오늘날짜의 연도가 같지 않으면 자동으로 제끼지 말고 수동으로 처리하기 위해 1 리턴. 2를 리턴하면 자동으로 처리할 지 살펴본다
	if(yearOfLastInvDate != yearOfToday){
;		SoundPlay, %A_WinDir%\Media\Ring02.wav
;		MsgBox, 262144, Title, 연도가 같지 않음
		;~ return 0
		return 1
	}
*/
/*
	; 마지막 배송 날짜의 월과 오늘날짜의 월이 같지 않으면 자동으로 제끼지 말고 수동으로 처리하기 위해 1 리턴. 2를 리턴하면 자동으로 처리할 지 살펴본다
	if(monthOfLastInvDate != monthOfToday){
		return 1

	}
*/	



	; 마지막 배송 날짜의 월과 오늘날짜의 월이 같지 않으면 자동으로 제끼지 말고 수동으로 처리하기 위해 1 리턴. 2를 리턴하면 자동으로 처리할 지 살펴본다
	if(monthOfLastInvDate != monthOfToday){

		; 이번달이 1월이면 이번달을 13으로 바꾸기
		; 1월의 지난달은 12월인데 1에서 1을 빼면 12가 아니라 0이 되니까
		if(monthOfToday == 01){
			monthOfToday = 13
		}
		
		; 마지막 배송일이 지난달인지 확인
		if(monthOfLastInvDate == monthOfToday-1){
			
;	MsgBox, 262144, Title, 마지막 배송일이 이번달 아닌 지난달
			
			;지난달의 마지막날이 31일인지 아닌지
			;~ if(monthOfToday == 01||03||05||07||08||10||12){
			if monthOfLastInvDate contains 01,03,05,07,08,10,12
			{
				dateOfToday := dateOfToday + 31
;	MsgBox, % "지난달은 마지막 날이 31일로 끝난다`n`n`n" . "monthOfLastInvDate : " . monthOfLastInvDate . "`n`ndateOfToday : " . dateOfToday
			}
			; 마지막날이 31이 아닌 달이라면
			else{
				; 지난달이 2월이라면
				;~ if(monthOfToday == 02){
				if(monthOfLastInvDate == 02){
					dateOfToday := dateOfToday + 28
;	MsgBox, % "지난달은 2월이다.`n`n`n" . "monthOfLastInvDate : " . monthOfLastInvDate . "`n`ndateOfToday : " . dateOfToday
				}
				; 이번달의 마지막날이 31일인 달도 아니고 2월도 아니라면
				else{
					dateOfToday := dateOfToday + 30
;	MsgBox, % "지난달은 마지막 날이 30일로 끝난다`n`n`n" . "monthOfLastInvDate : " . monthOfLastInvDate . "`n`ndateOfToday : " . dateOfToday
				}
			}
			
;	MsgBox, % "dateOfToday : " . dateOfToday . "`n`ndateOfLastInvDate : " . dateOfLastInvDate
			
		}
		; 지난달이 아니면 더 오래전에 마지막으로 배송했으므로 제끼지 않으면서도 수동이 아닌 자동으로 처리할 지 묻기 위해 2를 리턴
		else{
			return 2
			MsgBox, 제끼지 않고 자동처리 확인하기
		}
		
		;~ return 1
		
	}
	







	; 오늘이 월요일이나 화요일일 때 
	if(whatsTheDayToday == 2 || whatsTheDayToday == 3){
;		MsgBox, 오늘은 월요일 혹은 화요일
		
		; 마지막 배송 날짜가 지난 주 금요일 보다 최근일 때 주문 처리 않고 제끼기 위해 0 리턴
		if(dateOfLastInvDate >= dateOfToday - 4){
			;~ return 1
;			MsgBox, 주문 처리 않고 제끼기 위해 0을 리턴합니다
			return 0
		}
	}	
	; 그 이외의 요일일 때
	else{
;		MsgBox, 오늘은 월요일이나 화요일이 아니다
		
		; 마지막 배송 날짜가 이번 주의 이틀 전보다 최근일 때 주문 처리 않고 제끼기 위해 0 리턴
		if(dateOfLastInvDate >= dateOfToday - 2){ ; 2일 전까지는 너무 최근이라 안됨. 3일 전 주문부터는 내보낼 수 있음
		;~ if(dateOfLastInvDate >= dateOfToday - 3){ ; 3일 전까지는 너무 최근이라 안됨. 4일 전 주문부터는 내보낼 수 있음
			;~ return 1
;			MsgBox 0을 리턴합니다
			return 0
		}
	}	


	; 마지막 배송 날짜가 작년이나 지난 달이 아니고 이틀 이내의 최근도 아닐때
	; 이때는 제끼지 않으면서도 수동이 아닌 자동으로 처리할 지 묻기 위해 2를 리턴
	;~ return 0
;	MsgBox 2 을 리턴합니다
	return 2
}

























































































; FG 오더 처리
ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile){
	
	BuyerNotes := ""
	AdditionalInfo := ""
	StaffNotes := ""
	

/*
	; 크롬창이 열려있으면 닫을지 묻고 대답 없으면 5초 이내 닫기
	IfWinExist, ahk_class Chrome_WidgetWin_1
	{
		MsgBox, 4, , ALL CHROME BROWSERS WILL BE CLOSED.`nWOULD YOU LIKE TO CONTINUE?`n`nTHIS WINDOW IS CLOSED IN 5 SECONDS THEN CLOSE ALL CHROME WINDOWS, 3
		IfMsgBox No
		{
			MsgBox, 262144, WAIT TO CONTINUE, CLICK OK TO CONTINUE
		}
		
		WinClose, ahk_class Chrome_WidgetWin_1
		
	}
*/	
	
	if(RegExMatch(CustomerPO, "imU)MTR")){
		
;		MsgBox, % "CustomerPO : " . CustomerPO
		
		; FG 페이지에서 정보 읽어서 저장하기
		; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서
		; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
		Arr_FGInfo := F_driver.GettingInfoFromCurrentPage(CustomerPO, IsItFromNewOrder, IsItFromExcelFile)


		Arr_BillingAdd := Arr_FGInfo[1].Clone()
		Arr_ShippingAdd := Arr_FGInfo[2].Clone()
		Arr_CC := Arr_FGInfo[3].Clone()
		Arr_Memo := Arr_FGInfo[4].Clone()
		ShippingMethodStatus := Arr_FGInfo[5]

		
		BuyerNotes := Arr_Memo[1]
		AdditionalInfo := Arr_Memo[2]
		StaffNotes := Arr_Memo[3]
		CC# := Arr_CC[2]
		
		

		; 필요 없는 문자가 들어있을 경우를 대비해 메모들 값 정리해주기
		BuyerNotes := Trim(BuyerNotes)
		AdditionalInfo := Trim(AdditionalInfo)
		StaffNotes := Trim(StaffNotes)
		BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")
		AdditionalInfo := RegExReplace(AdditionalInfo, "[^a-zA-Z0-9 ]", "")
		StaffNotes := RegExReplace(StaffNotes, "[^a-zA-Z0-9 ]", "")
					
		
		; UPS Ground 값은 3이다. 3이 아니면 
;		MsgBox, % ShippingMethodStatus
		if(ShippingMethodStatus != 3)
		{
			SoundPlay, %A_WinDir%\Media\Ring02.wav
			MsgBox, 262144, UPS STATUS, IT IS NOT UPS GROUND SHIPMENT`n`n**급행**`n`nOK TO CONTINUE
		}
		
		
		; 메모가 있을때만 창 키워서 표시하기
		if(BuyerNotes || AdditionalInfo || StaffNotes){
			
			SoundPlay, %A_WinDir%\Media\Ring02.wav ; Ring03 이 이상하면 Ring02 써보기
			;~ SoundPlay, %A_WinDir%\Media\Ring03.wav
			MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
		}
		else{

			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
		}

		; No 눌렀으면 다시 시작
		IfMsgBox, No
		{
			N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
			return
;			Reload
		}
		
		
		; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
		N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
		
		


		SoundPlay, %A_WinDir%\Media\Ring06.wav
		MsgBox, 262144, Title, CHANGE SCREEN TO SO Manager
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
		return
;		Reload


	}
	
	return
	
} ; ProcessingFGOrder 메소드 끝
















; LASHOWROOM 오더 처리
ProcessingLASOrder(CustomerPO, LAS_driver, N_driver){

	; LASHOWROOM 오더 처리
	if(RegExMatch(CustomerPO, "imU)OP")){
		
		BuyerNotes := ""
		
		; LAS 페이지에서 정보 읽어서 저장하기
		; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서 
		; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
		Arr_FGInfo := LAS_driver.GetInfoFromLASPage(CustomerPO)


		Arr_BillingAdd := Arr_FGInfo[1].Clone()
		Arr_ShippingAdd := Arr_FGInfo[2].Clone()
		Arr_CC := Arr_FGInfo[3].Clone()
		Arr_Memo := Arr_FGInfo[4].Clone()

		
		BuyerNotes := Arr_Memo[1]
	;	AdditionalInfo := Arr_Memo[2] ; 이 정보는 없음
	;	StaffNotes := Arr_Memo[3] ; 이 정보는 없음
	;	CC# := Arr_CC[2] ; 이 정보는 없음


		; 필요 없는 문자가 들어있을 경우를 대비해 메모값 정리해주기
		BuyerNotes := Trim(BuyerNotes)
		BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")

		
		; 메모가 있을때만 창 키워서 표시하기
		;~ if(BuyerNotes || AdditionalInfo || StaffNotes){
		if BuyerNotes not in None
		{
			;~ SoundPlay, %A_WinDir%\Media\Ring02.wav ; Ring03 이 이상하면 Ring02 써보기
			SoundPlay, %A_WinDir%\Media\Ring03.wav
			MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
		}
		else{	

			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
		}



		; No 눌렀으면 다시 시작
		IfMsgBox, No
		{
;			Reload
			return
		}
		
		; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
		N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
		
		; N41에 카드 정보가 있는지 확인하기 위해 카드 정보 입력창 열기
		N_driver.OpenRegisterCreditCard()

		Sleep 2000
		MsgBox, 262144, Memo, PLEASE CLICK Ok TO RESTART THE APPLICATION
		
		WinClose, Credit Card Management
		
		N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기

;		Reload
		return
	}

} ; ProcessingLASOrder 메소드 끝
















; JODIFL WEB 오더 처리
if(RegExMatch(CustomerPO, "imU)JOD")){
	
	; Credit Sales Orders Small 탭에서 CustomerPO 검색한 뒤 열기
	L_driver.SearchPONumber(CustomerPO)
	
	; 램스에서 카드 정보 아닌 고객 메모 등 읽기
;	Arr_CSOS_Memo := L_driver.getInfoFromCSOS()
	

	; N41 창을 최소화하고 램스를 왼쪽모니터로 옮겨오기
	N41_wintitle := " N41"
	WinMinimize, %N41_wintitle%
	WinMove, LAMBS, , 0, 0
	WinMaximize, LAMBS

	
	; LAMBS 에서 카드 정보 읽어서 배열에 저장하기
	; 1~5 까지는 카드정보, 6~10까지는 주소정보 담겨있음
	Arr_CC_integration := L_driver.ReadingCCInfoFromLAMBS()
	
	
	; Arr_CC_integration에는 5개의 카드 값과 5개의 주소 값이 들어있지만 카드 갯수만큼만 루프 돌기
	;~ loop, 5{
	Loop, %#ofCC_counter%
	{		
		Arr_%A_Index%_CC := Arr_CC_integration[A_Index].Clone() ; 카드 정보는 배열의 1부터 넣으면 되지만
		Arr_%A_Index%_Billing := Arr_CC_integration[A_Index+5].Clone() ; 주소 정보는 6부터 있으니 인덱스에 5를 더한 6부터 시작한다
	}


	; cc정보가 없을수도 있지만 확인을 위해 변수에 값 넣기
	Name := Arr_1_CC[1]
	CC#_1 := Arr_1_CC[2]
	CC#_2 := Arr_2_CC[2]
	CC#_3 := Arr_3_CC[2]
	CC#_4 := Arr_4_CC[2]
	
;	MsgBox, % "Arr_1_CC[1]" . Arr_1_CC[1]
	
	CommN41.BasicN41Processing()
	
	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 4100, Memo, %BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`nREADY TO UPDATE CUSTOMER INFO`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.
	
	; No 눌렀으면 다시 시작
	IfMsgBox, No
		Reload


	; ## 고객 이름이 같은지 확인하기 ##	
	CommN41.ClickCustomerMasterTab()
		
	; To get Contact Name
	Text:="|<Contact>*149$41.S001005W00600O0QTSwRw1YmN9Bc2BYlm3E4P9YY6lAaHN9gwCAXTCC"
	if ok:=FindText(284,611,150000,150000,0,0,Text)
	{
		CoordMode, Mouse
		X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		MouseMove, (X+W)+80, Y+H//2
		Click
		Sleep 100
		Send, ^a^c
		Sleep 100
			
		;~ if Arr_CC[1] not contains %Clipboard%
		Name := % Arr_1_CC[1]
		; N41에 있는 이름이 웹이나 램스에서 가져온 값과 맞지 않으면 경고창 띄우기
		if Name not contains %Clipboard%
		{
			MsgBox, 4100, Alert, CONTACT NAME OF N41 IS NOT MATCHED WITH THE NAME ON FG`n`n`n`nWOULD YOU LIKE TO CONTINUE TO CHANGE ADDRESS INFO?`nIF YOU CLICK No, IT WILL RESTART THE APPLICATION.
				
			IfMsgBox, No
			{
				Reload
			}
			
			; Yes 눌렀으면 Contact에 값 입력하기
			IfMsgBox, Yes
			{
				MouseMove, (X+W)+80, Y+H//2
				Click				
				Send, % Name
			}
		}
			
	}
	
	

	; 입력할지 말지 결정하기 위해 카드 정보 입력창 열기
	N_driver.OpenRegisterCreditCard()
	
	MsgBox, 4100, Memo, CREDIT CARD NUMBER OF FG IS : `n%CC#_1%`n%CC#_2%`n%CC#_3%`n%CC#_4%`n`n`n`nWOULD YOU LIKE TO TRANSFER CC INFO TO N41?`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.

	; Yes 눌렀으면 N41 에 카드 정보 입력하기
	IfMsgBox, Yes
	{
		; 램스에 저장된 카드 갯수만큼만 루프 돌아서 카드정보 N41 에 입력하기
		Loop, %#ofCC_counter%
			N_driver.PutCCInfoInN41(Arr_%A_Index%_CC, Arr_%A_Index%_Billing)
		
		N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기		
		
		Reload
	}
	
	; No 눌렀으면 CC 창 닫고 어플 다시 시작하기
	IfMsgBox, No
	{
		WinClose, Credit Card Management
		
		N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
		
		Reload
	}

}

Exitapp








Esc::
Exitapp

^o::
URL = https://vendoradmin.fashiongo.net/#/home
CommWeb.OpenNewBrowser(URL)
return

^5::
MsgBox
N41_login_wintitle := "ahk_exe nvlt.exe"
WinWaitActive, N41_login_wintitle
WinMaximize, N41_login_wintitle
return

!z::
SendInput, ( available_allocate_qty > 0 )  
return

!x::
/*
SendInput, PLZ ADD ITEMS %A_MM%/%A_DD%
loop, 2
{
	send, ^s
}
*/
SendInput, PLZ ADD ITEMS ON PICK {#} %Clipboard%
Sleep 150
Send, {Enter}
return

^r::
run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " )
return

F5::
Reload


GuiClose:
ExitApp 


















































F12::

N41_ProcessingForPT_driver := New N41_ProcessingForPT

; pre-authorized 받고 프린트하기
IsItFromNewOrder = 0
N41_ProcessingForPT_driver.OpenSO_Print_WITH_PreAuthorized(CustomerPO, IsItFromNewOrder)

;~ IsItFromAllocation = 0
;~ PrintByPayment(IsItFromAllocation)
return



; pre-authorized 받고 인쇄하기
PrintByPayment(IsItFromAllocation){


	CN41_driver := New CommN41
	
	; Create Pick Ticket 버튼 클릭하기
	; Allocation 오더에서 온 것이 아닐 때는 오른쪽 위 Open So 에 있는 버튼 누르기
	if(!IsItFromAllocation){
		CN41_driver.ClickCreatePickTicketButton()
	}
	; Allocation 오더에서 온 것일때는 왼쪽 밑의 Open Allocation 에 있는 버튼 누르기
	else if(IsItFromAllocation){
		CN41_driver.ClickCreatePickTicketButtonOnOpenAllocation()
	}
	
/*	
	; 체크박스 클릭 하지 않고 픽티켓 버튼 눌렀을 때 나오는 경고창 나오면
	; 엔터쳐서 창 없애고 재귀호출해서 
	IfWinExist, SO Manager
	{
		Send, {Enter}
		MsgBox, 262144, Title, PLEASE CHECK ITEMS THAT SHIP OUT`nTHEN CLICK THE OK
		PrintByPayment(IsItFromAllocation)
	}
*/

	WinMinimize, Pick Ticket Processing
	WinMinimize, N41 Processing
	
	
	; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
	Sleep 5000	
	
	
	; Merge 확인 창
	IfWinActive, SO Manager
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		WinWaitActive, SO Manager
		IfWinActive, SO Manager
		{
			Sleep 500
			Send, {Enter}
			Sleep 1000			
			
			WinWaitActive, Pick Ticket ; Allocation 경고창
			IfWinActive, Pick Ticket
			{
				Sleep 500
				Send, {Left}
				Sleep 500
				Send, {Enter}
				Sleep 700
				
				
				WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
				IfWinActive, Pick Ticket
				{
					Sleep 500
					Send, {Enter}
					Sleep 800
					
					FromClickingPreAuthorizedButton_To_PrintOutPickTicket()
					return
				}				
			}			
		}
	}


	WinWaitActive, Pick Ticket ; Allocation 경고창
	IfWinActive, Pick Ticket
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		
		WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
		IfWinActive, Pick Ticket
		{
			Sleep 500
			Send, {Enter}
			Sleep 800
			
			; Pick Ticket 창이 또 나오면 이전에 에러 메세지 창이 나왔을 것
			IfWinActive, Pick Ticket
			{
				MsgBox, MAYBE 'Warehouse is required!' ERROR HAS BEEN OCCURED`n`nWAREHOUSE INFO ON Sales Order OF THIS ORDER HAS TO BE MODIFIED.`n`nIF OK BUTTON ON Pick Ticket WINDOW ON N41, ALL INFO WILL BE SET AS DEFAULT.
				return
			}
			
			FromClickingPreAuthorizedButton_To_PrintOutPickTicket()
		}
	}
	
	return
}










; CBS 위한 처리
; pre authorized 버튼 누르지 않음
F11::

N41_ProcessingForPT_driver := New N41_ProcessingForPT

; CBS 등 돈 받지 말고 프린트 하기
IsItFromNewOrder = 0
N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)


;~ IsItFromAllocation = 0
;~ CBS_Print(IsItFromAllocation)
return


CBS_Print(IsItFromAllocation){
	
	CN41_driver := New CommN41

	; Create Pick Ticket 버튼 클릭하기
	; Allocation 오더에서 온 것이 아닐 때는 오른쪽 위 Open So 에 있는 버튼 누르기
	if(!IsItFromAllocation){
		CN41_driver.ClickCreatePickTicketButton()
	}
	; Allocation 오더에서 온 것일때는 왼쪽 밑의 Open Allocation 에 있는 버튼 누르기
	else if(IsItFromAllocation){
		CN41_driver.ClickCreatePickTicketButtonOnOpenAllocation()
	}
		
	
	
/*	
	; 체크박스 클릭 하지 않고 픽티켓 버튼 눌렀을 때 나오는 경고창 나오면
	; 엔터쳐서 창 없애고 재귀호출해서 
	IfWinExist, SO Manager
	{
		Send, {Enter}
		MsgBox, 262144, Title, PLEASE CHECK ITEMS THAT SHIP OUT`nTHEN CLICK THE OK
		CBS_Print(IsItFromAllocation)
	}	
*/

	WinMinimize, Pick Ticket Processing
	WinMinimize, N41 Processing
	
	
	; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
	Sleep 5000	
	
	
	; Merge 확인 창
	IfWinActive, SO Manager
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		WinWaitActive, SO Manager
		IfWinActive, SO Manager
		{
			Sleep 500
			Send, {Enter}
			Sleep 700			
			
			WinWaitActive, Pick Ticket ; Allocation 경고창
			IfWinActive, Pick Ticket
			{
				Sleep 500
				Send, {Left}
				Sleep 500
				Send, {Enter}
				Sleep 700
				
				
				WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
				IfWinActive, Pick Ticket
				{
					Sleep 500
					Send, {Enter}
					Sleep 800

					PrintOut()
					return
				}				
			}			
		}
	}


	WinWaitActive, Pick Ticket ; Allocation 경고창
	IfWinActive, Pick Ticket
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		
		WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
		IfWinActive, Pick Ticket
		{
			Sleep 500
			Send, {Enter}
			Sleep 800			
			
			; Pick Ticket 창이 또 나오면 이전에 에러 메세지 창이 나왔을 것
			IfWinActive, Pick Ticket
			{
				MsgBox, MAYBE 'Warehouse is required!' ERROR HAS BEEN OCCURED`n`nWAREHOUSE INFO ON Sales Order OF THIS ORDER HAS TO BE MODIFIED.`n`nIF OK BUTTON ON Pick Ticket WINDOW ON N41, ALL INFO WILL BE SET AS DEFAULT.
				return
			}			
			
			PrintOut()
		}
	}

	return	
}
































^1::
;~ SetKeyDelay, 300
;~ SetKeyDelay 50,200
SetKeyDelay, 1000
;~ SetKeyDelay 300,200

Data = %Clipboard%

StringReplace, Data, Data, ', , All
StringReplace, Data, Data, -, , All
StringReplace, Data, Data, (, , All
StringReplace, Data, Data, ), , All
Data := Trim(Data)
StringUpper, Data, Data ; 대문자로 바꾸기

;~ StringLeft, Data, Data, 20  ; 왼쪽부터 20개 읽어서 저장하기

Send, %Data%
return




^2::
SetKeyDelay, 1000
Data = %Clipboard%

;~ RegExMatch(Data, "imU)(\d*)\.", SubPat)
;~ Data := SubPat1

Data := Trim(Data)
Send, %Data%
return




^3::
SetKeyDelay, 1000

Data = %Clipboard%

Data := RegExReplace(Data, "[^0-9]", "") ;숫자만 저장

StringReplace, Data, Data, ', , All
StringReplace, Data, Data, -, , All
StringReplace, Data, Data, (, , All
StringReplace, Data, Data, ), , All
StringReplace, Data, Data, %A_SPACE%, , All
StringReplace, Data, Data, `n, , All
StringReplace, Data, Data, `r, , All
StringUpper, Data, Data ; 대문자로 바꾸기
Data := Trim(Data)


;~ StringLeft, Data, Data, 20  ; 왼쪽부터 20개 읽어서 저장하기

Send, %Data%
return




!F2::
/*
run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)
Sleep 3000
driver := ChromeGet()

MsgBox, PAUSE
*/		

	WinActivate, FashionGo Vendor Admin - Google Chrome

	WinMinimize, N41 Processing
	
	; 만약 현재 페이지가 FG 페이지라면
	if(RegExMatch(driver.Url, "imU)fashiongo")){

		; Xpath 들
		TheBlankOfShippingFee_Xpath = //input[@class='ng-valid ng-dirty ng-touched']
		TheBlankOfShippingFee_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[3]/div/div[8]/div[2]/div/input
		
		SaveButton_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[5]/div/button
		AuthorizeButton_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li[2]/span[2]/div/div[1]/div/button[1]
		AuthorizeButton_Xpath = //*[contains(text(), 'Authorize')]
		OkButtonInAuthorizeWindow_Xpath = //*[@id='okButton']
		total#OfQty_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tfoot/tr/td[3] ; 전체 아이템 갯수
		;~ subTotlaPrice_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[3]/div/div[7]/div[2] ; 제품 가격 합계
		subTotlaPrice_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[3]/div/div[3]/div[2] ; 제품 가격 합계
		PaymentStatus_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[1]/ul/li[1]/span[2]/div/span[1] ; Payment Status
		
		
		total#OfQty := driver.FindElementByXPath(total#OfQty_Xpath).Attribute("innerText")
		subTotlaPrice := driver.FindElementByXPath(subTotlaPrice_Xpath).Attribute("innerText")
		
		; 알파벳, 숫자, 소숫점(.)만 저장 (알페벳, 숫자, 소숫점(.) 을 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
		subTotlaPrice := RegExReplace(subTotlaPrice, "[^a-zA-Z0-9.]", "")
		

;~ MsgBox, 262144, Title, totlaPrice : %totlaPrice%
		
		shippingFee = 0
		Loop{
			; 수량에 맞는 금액 입력하기
		;	SoundPlay, %A_WinDir%\Media\Ring06.wav
			;~ if #ofCheckBoxes between 1 and 4
			;~ if total#OfQty between 1 and 24
			if subTotlaPrice between 1 and 399
			{
				driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("40").sendKeys(driver.Keys.TAB)
		;		MsgBox, 262144, Title, #ofCheckBoxes : %#ofCheckBoxes%`nPUT IN $50
			}
			;~ else if #ofCheckBoxes between 5 and 10
			else if subTotlaPrice between 399.1 and 899
			{
				driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("60").sendKeys(driver.Keys.TAB)
		;		MsgBox, 262144, Title, #ofCheckBoxes : %#ofCheckBoxes%`nPUT IN $70
			}
			;~ else if #ofCheckBoxes between 11 and 100
			else if subTotlaPrice between 899.1 and 1300
			{
				driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("130").sendKeys(driver.Keys.TAB)
		;		MsgBox, 262144, Title, #ofCheckBoxes : %#ofCheckBoxes%`nPUT IN $90
			}
			else
			{
				driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("150").sendKeys(driver.Keys.TAB)
				MsgBox, 262144, Title, totlaPrice : $%subTotlaPrice%`nPUT IN $170
			}
			
			; Save 버튼 클릭 후 Authorize 버튼 클릭하기
			Sleep 500
			driver.FindElementByXPath(SaveButton_Xpath).click()
			Sleep 500
			driver.FindElementByXPath(SaveButton_Xpath).click()
			Sleep 500
			
			; 배송비가 입력 안 되는 에러를 줄이기 위해 화면 리프레쉬 후 화면을 아래 끝으로 내려보기
			; 페이지 리프레쉬 해보기
			driver.refresh()
			Sleep 1000			
			driver.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);")
			Sleep 500
			
			; 배송료가 제대로 저장됐는지 확인 후 금액이 들어있으면 루프 빠져나가기
			shippingFee := driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).Attribute("value")
			if(shippingFee != 0)
				break
			else
				shippingFee = 0
		}
		
;~ driver.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);")
;MsgBox, 262144, Title, shippingFee : %shippingFee%`n배송료 제대로 입력됐나?

		
		
		; 이미 다른 벤더에 의해 Decline 났으면 배송료만 입력한 뒤 이 핫키 빠져나가기
		theResult := driver.FindElementByXPath(PaymentStatus_Xpath).Attribute("innerText")
		if theResult contains Pending
		{
			SoundPlay, %A_WinDir%\Media\Ring02.wav
			MsgBox, 262160, Title, 디클라인!!`nIT'S DECLINED!!!`nCLICK OK TO CONTINUE
			
			; 디클라인 난 것 넘기지 않기 위해 화면 위로 옮기기
			;~ driver.FindElementByXPath(PaymentStatus_Xpath).click()			
			; 상태 확인하기 쉽게 하기 위해 화면 위로 올리기
			driver.ExecuteScript("window.scrollTo(document.body.scrollHeight,0);")						

			
			; N41 동작시키기 위한 메세지 창들 활성화 시키기 (마우스로 움직이기 귀찮아서)
			WinActivate, OPTIONS
			WinActivate, Memo
			WinActivate, UPS STATUS
			WinActivate, NOT APPROVED
			WinActivate, NOT PRE-AUTHORIZED
			
			return
		}		

		
		
;MsgBox, 262144, Title, totlaPrice : %totlaPrice%`n`n`n`n$1 ~ $399 = $60`n$400 ~ $899 = $80`n$900 ~ $1300 = $150`n`n배송료 제대로 입력됐나?
		
		
		; Save 버튼 클릭 후 Authorize 버튼 클릭하기
;		Sleep 3000
;		driver.FindElementByXPath(SaveButton_Xpath).click()
;		Sleep 1000
		
		; Authorize 버튼 생길때까지 기다린 뒤 클릭한 뒤 확인창이 나오면 OK 버튼 클릭하기
		Loop{
			Sleep 200
			if(driver.FindElementByXPath(AuthorizeButton_Xpath).isDisplayed()){
				Sleep 100
	;			MsgBox, DISPLAYED
				driver.FindElementByXPath(AuthorizeButton_Xpath).click()
				Sleep 150
				
				; 확인창 나오면 ok 버튼 클릭
				if(driver.FindElementByXPath(OkButtonInAuthorizeWindow_Xpath).isDisplayed()){
					Sleep 100
					driver.FindElementByXPath(OkButtonInAuthorizeWindow_Xpath).click()
					break
				}
				
				break
			}
		}
		
		
		
		Sleep 3000
		
		
		
		; 제대로 결제됐는지 확인
		; 1초마다 제대로 결제됐는지 확인 후 제대로 결제됐으면 루프 빠져나오기		
		
		Loop, 10
		{
			
			Sleep 1000
			
			payment_Status := driver.FindElementByXPath(PaymentStatus_Xpath).Attribute("innerText")
			
;~ MsgBox, % "Payment Status : " . payment_Status
		
			if payment_Status contains Authorized
			{
;~ MsgBox, escape from the loop
				break
			}
		}
		
		; 만약 결제가 안됐으면 메세지 띄우기
		theResult := driver.FindElementByXPath(PaymentStatus_Xpath).Attribute("innerText")
		if theResult contains Pending
		{
			SoundPlay, %A_WinDir%\Media\Ring02.wav
			MsgBox, 262144, Title, IT'S A PENDING ORDER`nCLICK OK TO CONTINUE
			;~ break
		}
		; 위에서 Authorize 버튼 안 눌러서 아직까지 Payment Status 가 Ready 이면
		else if theResult contains Ready
		{
			; Authorize 버튼 생길때까지 기다린 뒤 클릭한 뒤 확인창이 나오면 OK 버튼 클릭하기
			Loop{
				Sleep 200
				if(driver.FindElementByXPath(AuthorizeButton_Xpath).isDisplayed()){
					Sleep 100
		;			MsgBox, DISPLAYED
					driver.FindElementByXPath(AuthorizeButton_Xpath).click()
					Sleep 150
					
					; 확인창 나오면 ok 버튼 클릭
					if(driver.FindElementByXPath(OkButtonInAuthorizeWindow_Xpath).isDisplayed()){
						Sleep 100
						driver.FindElementByXPath(OkButtonInAuthorizeWindow_Xpath).click()
						break
					}
					
					break
				}
			}

		}
		
		
		;~ MsgBox, % driver.FindElementByXPath(PaymentStatus_Xpath).Attribute("outerHTML")
		
		
		; Pending 이 아닐때만 화면 아래로 내리기
		if theResult not contains Pending
		{
			; 금액 확인하기 쉽게 하기 위해 화면 아래로 내리기
			driver.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);")
			
			; 배송료 얼마 입력됐는지 띄워주기
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			shippingFee := driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).Attribute("value")
			MsgBox, 262144, Title, %shippingFee%`n`n배송료 금액 숫자가 있다면 배송료 제대로 입력된 것
		}


		
		;~ IfWinExist, OPTIONS
			;~ WinActivate, OPTIONS
			
			
		; N41 동작시키기 위한 메세지 창들 활성화 시키기 (마우스로 움직이기 귀찮아서)
		WinActivate, OPTIONS
		WinActivate, Memo
		WinActivate, UPS STATUS
		WinActivate, NOT APPROVED
		WinActivate, NOT PRE-AUTHORIZED

		SoundPlay, %A_WinDir%\Media\Ring06.wav
		
		return
		
	} ; if ends - 만약 현재 페이지가 FG 페이지라면
	
	; 만약 현재 페이지가 LAS 페이지라면
	else if(RegExMatch(driver.Url, "imU)lashowroom")){
		
		Xpath = //*[@id="orderedit_form"]/div[1]/div[9]/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]
		#ofItems := driver.FindElementByXPath(Xpath).Attribute("innerText")
		
;MsgBox, % "#ofItems : " . #ofItems
		
		
		; Xpath 들
		TheBlankOfShippingFee_Xpath = //*[@id="ship_charge"] ; 배송비 입력칸
		updateGrandTotal_Xpath = //*[@id="update_grandtotal"] ; 배송비 입력 후 업데이트 버튼
		preAuthorize_Xpath = //*[@id="authorize_payment"] ; PRE AUTHORIZE 버튼
		pre_AuthorizePayment_Xpath = //button[@value='Authorize Payment Now'] ; Pre-Authorize Payment 버튼. Xpath 가 계속 바뀌어서 이렇게 value 값으로 찾았다
		transactionsResult_Xpath = /html/body/div[2]/div[3]/table/tbody/tr[2]/td[6]/span ; 카드 결제 결과		
		subtotalPrice_Xpath = //*[@id="orderedit_form"]/div[1]/div[9]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]
		
		
		subtotalPrice := driver.FindElementByXPath(subtotalPrice_Xpath).Attribute("innerText")
		
		; 알파벳, 숫자, 소숫점(.)만 저장 (알페벳, 숫자, 소숫점(.) 을 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
		subtotalPrice := RegExReplace(subtotalPrice, "[^a-zA-Z0-9.]", "")
		
;~ MsgBox, % "subtotalPrice : ||" . subtotalPrice . "||"
		
		
		; 수량에 맞는 금액 입력하기
		;~ if #ofItems between 1 and 4
		;~ if #ofItems between 1 and 24
		if subtotalPrice between 1 and 399
		{
			driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("40")
	;		MsgBox, 262144, Title, #ofItems : %#ofItems%`nPUT IN $50
		}
		;~ else if #ofItems between 5 and 10
		else if totlaPrice between 399.1 and 899
		{
			driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("60")
	;		MsgBox, 262144, Title, #ofItems : %#ofItems%`nPUT IN $70
		}
		;~ else if #ofItems between 11 and 100
		else if totlaPrice between 899.1 and 1300
		{
			driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("130")
	;		MsgBox, 262144, Title, #ofItems : %#ofItems%`nPUT IN $90
		}
		else{
			driver.FindElementByXPath(TheBlankOfShippingFee_Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys("150")
			MsgBox, 262144, Title, subtotalPrice : $%subtotalPrice%`nPUT IN $170
		}
	
	
;MsgBox, 262144, Title, subtotalPrice : %subtotalPrice%`n`n`n`n$1 ~ $399 = $60`n$400 ~ $899 = $80`n$900 ~ $1300 = $150`n`n배송료 제대로 입력됐나?

		
		
		; 배송비 입력 후 업데이트 버튼 클릭
		driver.FindElementByXPath(updateGrandTotal_Xpath).click()
		
		; PRE AUTHORIZE 버튼 클릭
		driver.FindElementByXPath(preAuthorize_Xpath).click()		
		
		Sleep 3000
		
		; Pre-Authorize Payment 버튼 클릭
		driver.FindElementByXPath(pre_AuthorizePayment_Xpath).click()
		


		; 제대로 결제됐는지 확인
		; 1초마다 결제 결과 페이지로 넘어갔는지 확인 후 넘어갔으면 루프 빠져나온다
		Loop, 10
		{
			
			Sleep 1000
		
			if(RegExMatch(driver.Url, "imU)order_transaction_view")){
				Sleep 500
				break
			}
		}
		
		
;~ MsgBox, % driver.FindElementByXPath(transactionsResult_Xpath).Attribute("value")	
;~ MsgBox, % driver.FindElementByXPath(transactionsResult_Xpath).Attribute("textContent")
;~ MsgBox, % driver.FindElementByXPath(transactionsResult_Xpath).Attribute("innerText")
;~ MsgBox, % driver.FindElementByXPath(transactionsResult_Xpath).Attribute("innerTEXT")
		
		
		; 만약 결제가 안됐으면 메세지 띄우기
		Result := driver.FindElementByXPath(transactionsResult_Xpath).Attribute("innerText")
;MsgBox, 262144, Title, Result : %Result%
		if Result contains Failed
		{
			SoundPlay, %A_WinDir%\Media\Ring02.wav
			MsgBox, 262144, Title, IT'S A PENDING ORDER`nCLICK OK TO CONTINUE
			;~ break
		}
		; 결제 잘 처리됐으면 제대로 처리됐다는 효과음 플레이하기
		else
			SoundPlay, %A_WinDir%\Media\Ring06.wav
		
		
		; N41 동작시키기 위한 메세지 창들 활성화 시키기 (마우스로 움직이기 귀찮아서)
		WinActivate, OPTIONS
		WinActivate, Memo
		WinActivate, UPS STATUS
		WinActivate, NOT APPROVED
		


;MsgBox, 262144, Title, 잘 처리됐남?
	
		
		return
	}
	
	return
	
	
	
	
!F3::
	;~ driver.close() ; closing just one tab of the browser
	driver.quit()
	
	return
	
		


^+v::
	;~ MsgBox, ctrl shift v
	
	eMail_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[1]/ul/li[2]/span[3]
	sHippingPhone#_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[2]/ul/li[5]/span[1]
	
	eMail := driver.FindElementByXPath(eMail_Xpath).Attribute("innerText")
	sHippingPhone# := driver.FindElementByXPath(sHippingPhone#_Xpath).Attribute("innerText")
	
	;숫자, 대쉬문자(-), 괄호열기 괄호닫기만 저장 (숫자, 대쉬문자(-), 괄호열기문자 괄호닫기문자를 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
	;~ sHippingPhone# := RegExReplace(sHippingPhone#, "[^a-zA-Z0-9.]", "")
	sHippingPhone# := RegExReplace(sHippingPhone#, "[^0-9-()]", "")
	
;	MsgBox, % "eMail : " . eMail . "`n`nsHippingPhone# : ||" . sHippingPhone# . "||"

	N_driver := new N41
	CustCode := N_driver.clickPriority() ; Customer Master 에 있는 Priority 클릭하기
	
	; Priority 에 2 입력하기
	Send, 2
	
	
	
	; Country 로 이동
	Loop,11
	{
		Sleep 150
		Send, {Tab}
	}
	
	; Country 에 USA 입력하기
	Send, USA
	
	
	
	
	; E-MAIL 로 이동
	Loop,6
	{
		Sleep 150
		Send, {Tab}
	}
	
	; E-MAIL 에 email 값 입력하기
	Send, %eMail%





	
	; phone 로 이동
	Loop,2
	{
		Sleep 150
		Send, {Tab}
	}
	
	; phone 에 phone 값 입력하기
	Send, %sHippingPhone#%		
	
	
	
	
	
	return





^+o::


	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible

	
	; 14 열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(14), 1)	

return



^+k::


	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible
	
	; 2 열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(2), 1)
	
	Sleep 200
	
	; 3 열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(3), 1)	
	
	Xl.Rows(1).EntireRow.Insert

return


+F12::
Send, 999
return


^+!s::
Loop{
    
    Sleep 500
    
    ; N41 활성화 시키기
    WinActivate, ahk_class FNWND3126
	
	N_driver := new N41

	; Customer Master Tab 클릭하기
	N_driver.ClickCustomerMasterTab()	

	; Sales Order 클릭
	N_driver.ClickSalesOrderOnTheMenuBar()
	

    MsgBox, 262148, title, Would you like to stop this program?, 3
    
    IfMsgBox, Yes
        break
    
    IfMsgBox, Timeout
        continue
}

return





; ### 티파니 고객 찾아서 삭제하기 ###
^#!t::

Xl := ComObjActive("Excel.Application")	; 열려있는 엑셀 창 사용하기

; Constants
xlValues := -4163
xlWhole := 1


TiffanyOnlyCustomerList := ["AHN AND AHN COLLECTI", "ALLURE", "AMBIANCE-PLACERVILLE", "AUBURN UNIVERSITY", "BETSEYS BOUTIQUE SHO", "BETTY BE GOOD BOUTIQ", "BOHO BLU", "BOOP DE DOOP", "CHARLI ROSE-CO", "COBOS BOUTIQUE", "COBOS-STARKVILLE", "ENTOURAGE", "FINNLEYS", "FINNLEYS 2", "FLAUNT BOUTIQUE-ROUN", "FOCUSED HOLDINGS", "JOSIES BOUTIQUE", "JUST JEWELRY", "K-S WHOLESALE", "KIKILARUE", "MAGNOLIA BTQ FRANKLI", "ROOLEE", "SHOP DRESS UP", "STYLEFOX", "THE COLLECTION DBA D", "THE GIRLIE BOUTIQUE", "THE PINK LILY BOUTIQ", "THREADS BOUTIQUE", "TY ALEXANDERS", "UNIQUE BOUTIQUE-SUTT", "UNIVERSITY BOOK STOR", "WAKEFIELDS-INC", "WREN AND IVORY LLC", "MY STORY", "B&B DEP STORE SOUTH", "ANGELS FRAGRANCES", "ANGELS FRAGRANCES", "THEWILLOWTREE", "SHOP-STYLE-YOUR-SENS", "GINGERS COLLECTION", "PERFECTLY PRISCILLA", "ALISONS LETTERLAND-G", "AVENUE HILL"]


; 배열(TiffanyOnlyCustomerList)에 들어있는 값의 갯수만큼만 반복
Loop % TiffanyOnlyCustomerList.Maxindex(){	

	FindThis := TiffanyOnlyCustomerList[A_Index]  ; Look for this.	
	
	Loop{
		
		; Determine the total range. Get the last cell in that range.
		LastRow := Xl.Sheets(1).UsedRange.Rows.Count  ; Get the number of rows in the used range.
		MyRange := Xl.Sheets(1).Range("A1:Z" LastRow)  ; Get the range of all cells in column A-Z in the used range.
		LastCell := MyRange.Cells(MyRange.Cells.Count)  ; The last cell in the used range.

		; Find the first cell.
		FoundCell := MyRange.Find(FindThis, LastCell, xlValues, xlWhole)  ; LookIn:=xlValues, LookAt:=xlWhole
		FirstAddr := FoundCell.Address  ; The while-loop below will exit when it reaches this cell.	
			
		row# := RegExReplace(FirstAddr, "\D")	; 찾은 단어가 있는 셀 위치에서 숫자만 찾아서 배열에 저장
		
		; row# 에 값이 없으면, 즉 해당 키워드를 찾지 못해서 FirstAddr 변수에서 숫자를 추출하지 못했으면 루프 빠져나가기
		if(!row#)
			break

;MsgBox, % "The first found cell in the range is " FirstAddr " with a value of '" FoundCell.Value "'."

		; 찾은 티파니 고객의 row 삭제하기
		#ofRowToReadAddedOne :=  "A" . row#
;MsgBox, % #ofRowToReadAddedOne
		Xl.Sheets(1).Range(#ofRowToReadAddedOne).EntireRow.Delete
	}

;MsgBox, % FindThis . " 이(가) 이 엑셀파일에 없거나 모두 찾아서 삭제했습니다. 다음으로 넘어갑니다."
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가

}

MsgBox, 262144, Title, 티파니 고객 모두 지웠음
return






;## 엑셀에서 중복 자료 삭제하기 ##
^#!d::

	WinMinimize, N41 Processing

	MsgBox, 262144, Title, 중복된 고객 코드를 지운 뒤 아이템 번호별로 다시 정렬합니다. 시작 전에 정렬할 열(Columns) 값을 확인 후 바꿔야 됩니다.`n`nallocate 된 파일은 바꿀 필요 없습니다.

	; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
	IfWinNotExist, ahk_class XLMAIN
	{
		loop{
			MsgBox, 262144, No Excel file Warning, PLEASE OPEN AN ORDER EXCEL FILE
			IfWinExist, ahk_class XLMAIN
				break
		}
	}	
	
	
	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible
	
	

	
		; 중복값 지우기 위해 일단
		; 고객 코드가 있는 2열(Columns)로 정렬
		xl.cells.sort(xl.columns(2), 1)
		;~ xl.cells.sort(xl.columns(7), 1)	; BO by Style 에서는 7열에 고객 코드가 있다

		;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
		XL_Handle(XL,1) ;get handle to Excel Application
		i := XL_Last_Row(XL)
		
;		MsgBox 고객코드로 정렬 했습니다. 마지막 행 번호는 %i%

		
		locOfCheckColumns := "B"		; B에 고객 코드가 있다
		;~ locOfCheckColumns := "G"		; BO by Style 에서는 G에 고객 코드가 있다
		
		j = 1
		Loop{

			k := j + 1

			#ofRowToRead := locOfCheckColumns . j
			#ofRowToReadAddedOne :=  locOfCheckColumns . k
			
			
			var1 := Xl.Range(#ofRowToRead).Value
			var2 := Xl.Range(#ofRowToReadAddedOne).Value
			
;			MsgBox, % "j : " . j . "`nk : " . k . "`n`n" . "var1 : " . var1 . "`nvar2 : " . var2
			
			
			if(var1 == ""){
				break
			}
						

			; 만약 지금 얻은 SO# 값이 이전 SO# 값을 저장하고 있는 previousNumber 값과 같다면 
			; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
			IfEqual, var1, %var2%
			{
				
;			MsgBox, % var1 . "`n" . var2 . "`n`n" . "delete" . "`n`n" . "i : " . i . "`nj : " . j
;			MsgBox, % #ofRowToReadAddedOne . "를 지웁니다"
				Xl.Sheets(1).Range(#ofRowToReadAddedOne).EntireRow.Delete
				
				continue
			}
			
			IfNotEqual, var1, %var2%
			{
				j++
				
				continue
			}
			
			

		}
		
		;~ MsgBox, 262144, Title, IT'S END
;		return
		
		
		
		
;MsgBox 스타일 별로 정렬한 뒤 20개 넘는 아이템들은 빈 줄 삽입하기. 

	; 20개 넘는 아이템만 빈줄 삽입해서 나누는 메소드 호출하기
	insertAEmptyRowForItemsMoreThan20()
	
		
	; 20개 넘는 아이템만 빈줄 삽입해서 나누는 메소드
	insertAEmptyRowForItemsMoreThan20(){
		
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible

		j = 1 ; 기준 스타일 번호의 줄 저장 변수
		k = 1 ; 비교할 스타일 번호의 줄 저장할 변수
		#ofStyle = 0 ; 같은 스타일이 몇개인지 세기 위해
		
		
		; 색깔이 있는 4열(Columns)로 정렬
		xl.cells.sort(xl.columns(4), 1)
		; 스타일 넘버가 있는 3열 정렬
		xl.cells.sort(xl.columns(3), 1)
		
		
		locOfStyle# := "C"
		
		Loop{

			#ofRowToRead := locOfStyle# . j
			#ofRowToReadAddedOne :=  locOfStyle# . k
			
			
			standStyle# := Xl.Range(#ofRowToRead).Value
			Style#ToBeCompared := Xl.Range(#ofRowToReadAddedOne).Value
			
			;~ MsgBox, % standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k
			
			; 스타일 번호가 없는 빈칸이면 루프 끝내고 나가기
			if(standStyle# == ""){
				break
			}
						

			; standStyle# 변수와 Style#ToBeCompared 변수 값이 같으면
			; 다음 아이템 비교하기위해 루프 계속 진행하기
			IfEqual, standStyle#, %Style#ToBeCompared%
			{
;				MsgBox, % "아이템이 같음`n`n" . standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k

				; 다음줄로 넘어가기 위해
				k := k + 1
				
				; 같은 아이템이 몇개인지 확인하기 위해
				#ofStyle++ 
				
				continue
			}
			
			; 기준 스타일 번호와 지금 스타일 번호가 다르면 
			else IfNotEqual, standStyle#, %Style#ToBeCompared%
			{
				
;				MsgBox, % "아이템이 다름`n`n" . standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k
				
				; 20개 넘는 아이템만 빈줄 삽입해서 나누기
				;~ if(#ofStyle >= 2){ ; #################################################### 아이템이 1개든 2개든 스타일번호가 바뀔때마다 빈줄 넣고싶을때는 이것 사용하기 ####################################################
				if(#ofStyle >= 20){
					
;					MsgBox, % standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k
					
					; 10 개 넘는 아이템의 끝을 표시해주기 위해 엑셀에 빈줄 넣기
					Xl.Rows(k).EntireRow.Insert
					
					; 처음을 표시해주기 위해 빈줄 삽입해야하는데 만약 처음 표시할 빈줄 이전에 빈줄이 있다면 빈줄 삽입하지 않기. 만약 빈줄을 삽입하게 되면 빈줄이 연달아 2개가 되니까
					#ofRowToRead := "C" . j-1
					standStyle# := Xl.Range(#ofRowToRead).Value
					
					; 앞선 줄이 빈줄이 아닐때만 
					if(standStyle# != ""){
						
						; 처음을 다른 아이템과 나누기 위한 빈줄 삽입
						Xl.Rows(j).EntireRow.Insert
						
						; 10개 넘는 아이템의 처음과 끝을 표시해주기 위해 처음과 맨끝의 빈줄 2줄 입력했기 때문에 2 증가
						j := k + 2
						k := k + 2
						#ofStyle = 1						
;						MsgBox, % "j : " . j . "`n" . "k : " . k
						continue
					}
					
					; 20개 넘는 아이템이지만 첫줄은 삽입 않고 마지막 줄만 삽입했기 때문에 1 증가
					j := k + 1
					k := k + 1
					#ofStyle = 1
;					MsgBox, % "j : " . j . "`n" . "k : " . k
					continue
				}
				
				; 스타일 번호가 같은 갯수가 20개가 안되면 다음으로 넘어가기
				; k는 이미 스타일 번호가 다른 다음줄이기 때문에 기준 스타일 번호 줄을 k로 하기
				j := k
				#ofStyle = 1
;				MsgBox, % "다음의 j k 값`n`n" . "j : " . j . "`n" . "k : " . k
				continue
			}
			
			
			;~ GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가

		}
		

;		SoundPlay, %A_WinDir%\Media\Ring06.wav
;		MsgBox, 262144, Title, THE ITEMS HAVE BEEN DEVIDED.

		return		
		
		
		
		
	} ; method ends - insertAEmptyRowForItemsMoreThan20()
			
		
		
		
		
		
		

/*
	;엑셀 값의 끝 row 번호 알아낸 후 lastLine#OfRow 에 값 넣기
	XL_Handle(XL,1) ;get handle to Excel Application
	;~ i := XL_Last_Row(XL)
	lastLine#OfRow := XL_Last_Row(XL)
	;~ GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
MsgBox, % "lastLine#OfRow : " . lastLine#OfRow	
*/	

MsgBox, 4100, Wintitle, 끝

return











; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지 
FromClickingPreAuthorizedButton_To_PrintOutPickTicket(){
	
		CN41_driver := New CommN41
	
		; pre authorized 버튼 클릭
		Text:="|<pre-authorize Button>*205$16.001zzbzyTztzzY0SE1tzzbzyTztzzc01zzy"
		if ok:=FindText(718,129,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			
			Sleep 1000
			
			; Pre-Authorized 통과 됐거나 Declined 됐을 때
			WinWaitActive Credit Card Processing, , 4
			IfWinActive, Credit Card Processing
			{
				Sleep 500
				Send, {Enter}
				Sleep 500
				
				; Print 버튼 클릭
				Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
				if ok:=FindText(359,129,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, X+W//2, Y+H//2
					Click					
					
					Sleep 1000
					
					; 프린트 창 최대화 하기
					WinWaitActive, Pick Ticket Print
					WinMaximize, Pick Ticket Print
;~ /*					
					; 안에 있는 프린트 버튼 클릭
					Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
					if ok:=FindText(199,44,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, Y+H//2
						Click
						
						Sleep 500
						
						; 에러창 나오면 프로그램 다시 시작하기
						IfWinActive, Microsoft Visual C++ Runtime Library
						{						
							MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
							return
;							Reload
						}
						
						Send, {Down}
						Sleep 200
						Send, {Down}
						Sleep 200
						Send, {Enter} ; Print Now 눌러서 인쇄하기
						
						Sleep 3000
						WinActivate, Pick Ticket Print
						WinClose, Pick Ticket Print ; 프린트 창 닫기
						Sleep 700						
						
						;~ CommN41.runN41() ; N31 활성화 한 뒤 
						;~ CommN41.OpenSOManager() ; SO Manager 탭 열고 끝내기						
						;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
						
						; SO MANAGER 탭 누르고 끝내기						
						CN41_driver.ClickREfresh()
						
						Send, {Enter} ; 리프레쉬 버튼 누른 뒤 
						Sleep 700
						
						result := CN41_driver.DoesThisPickTicketApproved() ; Approved 됐는지 화면에서 찾아본 뒤 찾았으면 1을 리턴하고 못 찾았으면 0을 리턴
						if(result == 0){					
							MsgBox, IS THIS PICK TICKET GOT APPROVED?							
						}						
						
						; SO Manager 탭 클릭해서 pick ticket 탭에서 나오기
						CN41_driver.OpenSOManager()
						
						; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
						; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
						CN41_driver.ClickREfreshButtonOnSOManager()
						

;						Reload
						return
						
					}
*/									
					
					
				}
							
				
			}
				
			; CC 가 없어서 업데이트 할거냐고 물을 때
			IfWinActive, Pick Ticket
			{
;				SoundPlay, %A_WinDir%\Media\Ring06.wav
;				MsgBox, 262144, Title, CC update`n`nCHECK THE SHIP VIA`n`nCLICK OK TO CONTINUE
				
				Sleep 300				

				Send, {Right}
				Sleep 200
				Send, {Enter}
				Sleep 500			


				
				; Print 버튼 클릭
				Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
				if ok:=FindText(359,129,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, X+W//2, Y+H//2
					Click					
					
					Sleep 1000
					
					; 프린트 창 최대화 하기
					WinWaitActive, Pick Ticket Print
					WinMaximize, Pick Ticket Print
;~ /*					
					; 안에 있는 프린트 버튼 클릭
					Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
					if ok:=FindText(199,44,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, Y+H//2
						Click
						
						Sleep 500
						
						; 에러창 나오면 프로그램 다시 시작하기
						IfWinActive, Microsoft Visual C++ Runtime Library
						{						
							MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
;							Reload
							return
						}
						
						Send, {Down}
						Sleep 200
						Send, {Down}
						Sleep 200
						Send, {Enter} ; Print Now 눌러서 인쇄하기
						
						Sleep 3000
						WinActivate, Pick Ticket Print
						WinClose, Pick Ticket Print ; 프린트 창 닫기
						Sleep 700						
						
						;~ CommN41.runN41() ; N31 활성화 한 뒤 
						;~ CommN41.OpenSOManager() ; SO Manager 탭 열고 끝내기						
						;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
						
						SoundPlay, %A_WinDir%\Media\Ring06.wav
						MsgBox, 262144, Title, NO CC INFO ON THIS CUSTOMER`n`n`n`nCHECK THE SHIP VIA`n`nCLICK OK TO CONTINUE

						; SO MANAGER 탭 누르고 끝내기
						CN41_driver := New CommN41
						CN41_driver.OpenSOManager()
						
						
						; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
						; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
						CN41_driver.ClickREfreshButtonOnSOManager()						
						
;						Reload
						return
						
					}
				}

				;~ Reload
				
			}

		}

	return
}








; 프린트 하기
PrintOut(){	
	
	; Print 버튼 클릭
	Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
	if ok:=FindText(359,129,150000,150000,0,0,Text)
	{
		CoordMode, Mouse
		X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		MouseMove, X+W//2, Y+H//2
		Click					
					
		Sleep 1000
					
		; 프린트 창 최대화 하기
		WinWaitActive, Pick Ticket Print
		WinMaximize, Pick Ticket Print
				
		; 안에 있는 프린트 버튼 클릭
		Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
		if ok:=FindText(199,44,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
						
			Sleep 500
					
			; 에러창 나오면 프로그램 다시 시작하기
			IfWinActive, Microsoft Visual C++ Runtime Library
			{						
				MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
;				Reload
				return
			}
						
			Send, {Down}
			Sleep 200
			Send, {Down}
			Sleep 200
			Send, {Enter} ; Print Now 눌러서 인쇄하기
						
			Sleep 3000
			WinActivate, Pick Ticket Print
			WinClose, Pick Ticket Print ; 프린트 창 닫기
			Sleep 700						
					
			;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262148, CBS_ORDER, CBS, CALL FOR CC, or FG PA`n`nCLICK OK TO CONTINUE.`nIF NO CLICKED, IT'S GOING TO PROCESS FOR FG PA
			
			; 하우스 메모에 FG PA 입력하기
			IfMsgBox, No
			{				
				CommN41_driver := New CommN41
				CommN41_driver.PutMemoIntoHouseMemoOnPickTicket()
				Send, FG PA
				Sleep 100
				Send, ^s
			}			

			
			; SO MANAGER 탭 누르고 끝내기
			CN41_driver := New CommN41
			CN41_driver.OpenSOManager()

			; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
			; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
			CN41_driver.ClickREfreshButtonOnSOManager()
			
;			Reload
			return
		}

	}	

	return	
}
