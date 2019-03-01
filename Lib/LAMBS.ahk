#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include MethodsForAll.ahk
#Include CommLAMBS.ahk




; Create Sales Orders Small 에서 정보 읽어오기
class LAMBS extends CommLAMBS
{
	
	; 램스의 CC 창에서 CC 정보 읽기
	getCCInfoFromCCWindowOfLAMBS(){
		
		driver := New MethodsForAll
		
		windowtitle := "Credit Card ( P999131 )"
		
		; CC Type 저장
		ControlGetText, CCType, WindowsForms10.EDIT.app.0.378734a17, %windowtitle%
		
		; Name on Card 저장
		ControlGetText, CCName, WindowsForms10.EDIT.app.0.378734a13, %windowtitle%
		StringUpper, CCName, CCName ; 대문자로 바꾸기
		
		; CC번호 저장
		ControlGetText, CCNumbers, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%

		; CVS 저장
		ControlGetText, CVV, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
		
		; 만료일 저장
		ControlGetText, ExpDate, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
		
		ExpDate := driver.Refine_ExpDate(ExpDate)

		; ADD 1 저장
		ControlGetText, ADD1, WindowsForms10.EDIT.app.0.378734a9, %windowtitle%
		StringUpper, ADD1, ADD1 ; 대문자로 바꾸기

		; ADD 2 저장
		ControlGetText, ADD2, WindowsForms10.EDIT.app.0.378734a8, %windowtitle%
		StringUpper, ADD2, ADD2 ; 대문자로 바꾸기

		; CITY 저장
		ControlGetText, CITY, WindowsForms10.EDIT.app.0.378734a7, %windowtitle%

		; STATE 저장
		ControlGetText, STATE, WindowsForms10.EDIT.app.0.378734a6, %windowtitle%

		; ZIP 저장
		ControlGetText, ZIP, WindowsForms10.EDIT.app.0.378734a5, %windowtitle%

		; COUNTRY 저장
		ControlGetText, COUNTRY, WindowsForms10.EDIT.app.0.378734a4, %windowtitle%


;		MsgBox, % CCType . "`n`n" . CCName . "`n`n" . CCNumbers . "`n`n" . CVV . "`n`n" . ExpDate . "`n`n" . ADD1 . "`n`n" . ADD2 . "`n`n" . CITY . "`n`n" . STATE . "`n`n" . ZIP . "`n`n" . COUNTRY
		
		Array := [CCType, CCName, CCNumbers, CVV, ExpDate, ADD1, ADD2, CITY, STATE, ZIP, COUNTRY]
		Sleep 1000
		
		return Array		
	}
	
	
	
	; CC 창 기다렸다가 아무데나 클릭해서 한 칸 내려서 다음 CC 정보 읽도록 처리해주기
	ToMoveNextCCInfo(){
		
		WinActivate, Credit Card ( P999131 ) ; 아래 DllCall 이 화면에 상대적이기 때문에 활성화 해주기

		; 아무데나 클릭해서 화살표 내려주면 다음 cc로 넘어가기 때문에 
		Text:="|<Default>*152$49.yDrkV2EzlY20MV828G10IEY1490U88G0W4yT0Y90F2E84G4U8V843x2E4FY212V82DXx11D7l4"

		if ok:=FindText(3020,595,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Send, {Down}
		}

		Sleep 1000
			
		return
	}

	
	
	
	; 램스에서 cc 정보 얻는 통합버전
	ReadingCCInfoFromLAMBS(){
		

		; 마우스 커서는 사용자의 물리적 마우스 이동에 반응하지 않습니다
		BlockInput, MouseMove
		BlockInput, MouseMoveOff
		
		
		CoordMode, mouse, relative
		
		;LAMBS활성화 후 화면 초기화
		CommLAMBS.Start()
		LAMBS.OpenCreateSalesOrdersSmallTab()
		Sleep 1000

		; LAMBS에 있는 인보이스 밸런스 값 읽어오기
	;	GetInvoiceBalanceOnLAMBS()


		;CC 버튼 찾고 클릭
		LAMBS.FindCCButtonAndClickIt()
		WinWaitActive, Credit Card ( P999131 ) ; CC 창 뜰 때까지 기다리기
		
		
		i = 1 ; i 값이 증가한 만큼 배열 갯수가 만들어 졌다. 만들어진 배열 갯수가 곧 LABMS에 저장되어 있는 카드의 갯수
		loop{
			
			Array%i% := LAMBS.getCCInfoFromCCWindowOfLAMBS() ; LAMBS의 cc 창에서 정보 얻은 후 배열에 저장
			
			; 카드 정보가 들어있는 배열의 세번째 값에 아무것도 들어있지 않으면 LAMBS예는 CC값이 없다는 뜻이고 루프를 빠져나간다
			if(Array%i%[3] == ""){
				MsgBox, No CC info in LAMBS
				WinClose, Credit Card ( P999131 )
				break
			}
			
			; 이전 카드 번호와 같은 카드 번호가 들어있으면 중복된 정보가 들어있다는 뜻이므로 중복된 마지막 배열을 삭제 후 루프 중단
			if(Array%i%[3] == previousCCNum){
				Array%i%.remove() ; previousCCNum 와 같은 값이 들어있는 배열은 중복된 것이니 지워주기
				--i ; if 들어오기 전에 값을 한 번 읽어줬기 때문에 i 값 줄이기
				WinClose, Credit Card ( P999131 )
				break
			}
		
			; CC 창 기다렸다가 아무데나 클릭해서 한 칸 내려서 다음 CC 정보 읽도록 처리해주기
			LAMBS.ToMoveNextCCInfo()
			
			previousCCNum := Array%i%[3]
			
			i++ ; 배열이 몇 개 만들어 졌는지 세기 위해
		}
		
		; 카드 갯수 몇개인지 저장
		#ofCC_counter := i
		
;		MsgBox % "A number of Arry(CC) is " #ofCC_counter
		
		
		
		; 주소정보 옮기는 처리 위한 배열 선언하기
		1_BillingAdd := object()
		2_BillingAdd := object()
		3_BillingAdd := object()
		4_BillingAdd := object()
		5_BillingAdd := object()
		
		arrayCounter4RemoveAddressInfo = 1 ; 카드 정보 배열에 있는 주소값 제거하는데 사용할 카운터 변수
		
		; 카드 갯수만큼만 루프 돌기
		Loop, %#ofCC_counter%
		{
			; CCType 값을 축약어로 바꾸기
			CCType := Array%a_index%[1]
			
			if CCType contains visa
				CCType = v
			else if CCType contains master
				CCType = m
			else if CCType contains american
				CCType = a
			else if CCType contains discover
				CCType = d
			else if CCType contains amex
				CCType = a
			
			Array%a_index%[1] := CCType
			
;			MsgBox, % a_index . " Abriviated CC Type : " . Array%a_index%[1]
			
			
			; 카드 정보 순서 변경하기
			Arr_Temp := Array%a_index%.Clone()
			
			Array%a_index%[1] := Arr_Temp[2] ; Name
			Array%a_index%[2] := Arr_Temp[3] ; CC #
			Array%a_index%[3] := Arr_Temp[1] ; CC Type
			Array%a_index%[4] := Arr_Temp[5] ; Exp
			Array%a_index%[5] := Arr_Temp[4] ; CVV
			
			
			
			
			; 6번부터 11번까지 있는 주소정보 _BillingAdd 배열에 넣기
			ADD1_or_ADD1_and_ADD2 := Array%a_index%[6] ; ADD1 (or ADD1+ADD2)
			ADD2 := Array%a_index%[7] ; ADD2
			CITY := Array%a_index%[8] ; CITY
			STATE := Array%a_index%[9] ; STATE
			ZIP := Array%a_index%[10] ; ZIP
			Country := Array%a_index%[11] ; USA
			
			; Array 배열 6번부터 11번까지 남아있는 주소정보 삭제하기
			loop, 6
			{
				Array%arrayCounter4RemoveAddressInfo%.Remove()
				Sleep 100
			}
			
			
			; 각 변수에 옮겨놓은 주소값 _BillingAdd 배열에 넣기
			%A_Index%_BillingAdd.Insert(ADD1_or_ADD1_and_ADD2)
			%A_Index%_BillingAdd.Insert(ADD2)
			%A_Index%_BillingAdd.Insert(CITY)
			%A_Index%_BillingAdd.Insert(STATE)
			%A_Index%_BillingAdd.Insert(ZIP)
			%A_Index%_BillingAdd.Insert(Country)
			
			; 주소값 제거하기 위한 카운터 값 증가
			arrayCounter4RemoveAddressInfo++
		}


/*
		;~ /* 카드정보 확인
		Loop % Array1.Maxindex(){
			MsgBox % "카드1 Element number " . A_Index . " is " . Array1[A_Index]
		}		

		;~ /* 주소 확인
		Loop % 1_BillingAdd.Maxindex(){
			MsgBox % "주소1 Element number " . A_Index . " is " . 1_BillingAdd[A_Index]
		}
*/			


		; 한 고객이 많아봤자 2,3개 카드 갖고 있으니 넉넉하게 5개 카드 관련 정보 리턴하기
		Array := [Array1, Array2, Array3, Array4, Array5, 1_BillingAdd, 2_BillingAdd, 3_BillingAdd, 4_BillingAdd, 5_BillingAdd]		
		
		return, Array
	}
	
	
	
	; 램스에 있는 메모 등 정보 얻기
	getInfoFromCSOS(){

		;~ LAMBS.BasicLAMBSProcessing()
		
		CommLAMBS.Start()
		
		; CSOS 탭 활성화 하기
		CommLAMBS.ActiveCreateSalesOrdersSmallTab()
				
		wintitle = LAMBS
		
		; Customer PO 1 클릭하기
		DllCall("SetCursorPos", int, 59-8, int, 262-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		Click
		
		
		; Company Name 얻기
		DllCall("SetCursorPos", int, 87-8, int, 375-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, CompanyName, %control%, %wintitle%

		
		; Customer 밑에 있는 메모 POSourceOrMemo 변수에 저장하기
		DllCall("SetCursorPos", int, 83-8, int, 406-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, POSourceOrMemo, %control%, %wintitle%
		POSourceOrMemo := RegExReplace(POSourceOrMemo, "imU)FASHIONGO", "")
		POSourceOrMemo := RegExReplace(POSourceOrMemo, "imU)LASHOWROOM", "")
		POSourceOrMemo := RegExReplace(POSourceOrMemo, "imU)WEB", "")
		

		; 창이 제대로 열렸는지 비교하기 위한 Order Id 얻기
		DllCall("SetCursorPos", int, 217-8, int, 293-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, OrderId, %control%, %wintitle%
		
		; OrderId 값은 빈 값이 없다. 만약 아무런 값도 없으면 메소드 재귀호출
		if(OrderId == "")
			LAMBS.getInfoFromCSOS()		


		; Customer Memo 값 CustomerMemoOnLAMBS 변수에 저장하기
		DllCall("SetCursorPos", int, 67-8, int, 466-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).	
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, CustomerMemoOnLAMBS, %control%, %wintitle%
		StringUpper, CustomerMemoOnLAMBS, CustomerMemoOnLAMBS ; Staff only notes 대문자로 바꾸기		
		

		; Sales Order Memo 값 SalesOrderMemoONLAMBS 변수에 저장하기
		DllCall("SetCursorPos", int, 69-8, int, 563-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, SalesOrderMemoONLAMBS, %control%, %wintitle%
		SalesOrderMemoONLAMBS := RegExReplace(SalesOrderMemoONLAMBS, "imU)Handling\sFee:\s.0.00", "")
		StringUpper, SalesOrderMemoONLAMBS, SalesOrderMemoONLAMBS ; Staff only notes 대문자로 바꾸기
		
		
		; Shipping Add 의 PhoneNumber 값 읽기
		DllCall("SetCursorPos", int, 583-8, int, 609-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, PhoneNumber, %control%, %wintitle%
		
		
		; PO 번호 CurrentPONumber 변수에 저장하기	
		DllCall("SetCursorPos", int, 994-8, int, 378-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, CurrentPONumber, %control%, %wintitle%
		
		
		; Customer PO 2 클릭하기
		DllCall("SetCursorPos", int, 156-8, int, 257-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		Click


		; Email 주소 Email 에 저장하기	
		DllCall("SetCursorPos", int, 1011-8, int, 399-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, Email, %control%, %wintitle%


		; Customer PO 1 클릭하기
		DllCall("SetCursorPos", int, 59-8, int, 262-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		Click
		
		
		; Billing ADD1 값 ADD1 변수에 저장하기
		DllCall("SetCursorPos", int, 777-8, int, 538-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, ADD1, %control%, %wintitle%


		; Billing ADD2 값 ADD2 변수에 저장하기
		DllCall("SetCursorPos", int, 777-8, int, 538-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, ADD2, %control%, %wintitle%
		
			
		
		
		;~ Array := [CompanyName, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, PhoneNumber, CurrentPONumber, Email, ADD1, ADD2, CITY, STATE, ZIP, COUNTRY]
		
		Array := [CompanyName, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, PhoneNumber, CurrentPONumber, Email]		
		
		
		;~ MsgBox, % CompanyName "`n" . POSourceOrMemo . "`n" . CustomerMemoOnLAMBS . "`n" . SalesOrderMemoONLAMBS . "`n" . PhoneNumber . "`n" . CurrentPONumber . "`n" . Email
		;~ MsgBox, % Array[1]
		
		return, Array
		
		
	}
	
	
	
	; Credit Sales Orders Small 탭에서 CustomerPO 검색한 뒤 열기	
	SearchPONumber(PO_Number){
		
		;상태바에서 알트키, 방향키 등 눌러서 Create Sales Orders Small Tab열기
		CommLAMBS.ActiveCreateSalesOrdersSmallTab()
		
		;~ OpenCreateSalesOrdersSmallTab()
		
		Send, {F3}
		WinWaitActive, Find

		; 커서 상태가 작업처리중이면 끝날때까지 기다리기
		while (A_cursor = "Wait")
			Sleep 500

		Sleep, 500
		
		ControlSetText, WindowsForms10.EDIT.app.0.378734a4, Cust PO NO, Find
		;~ ControlSetText, WindowsForms10.EDIT.app.0.378734a4, {Enter}, Find
		
		ControlClick, WindowsForms10.Window.8.app.0.378734a7, Find, ,l
		Send, {Down}
		Send, {Down}
		Send, {Enter}
		
		
		Send, {Tab}
		Send, %PO_Number%
		Send, {Enter}
		Send, {Enter}



		Sleep 2000 ; 위의 동작으로 검색한 뒤에는 이정도 시간을 둬야됨





		Text:="|<CUST PO NO>*126$63.D0083kQ0l3W8010F4E68WUF7S2910t8A2990F88591UF882910Z8A28t0S884d1UF182110X8++M90E4E4MWDDCC20Q0V3Y"

		if ok:=FindText(697,417,150000,150000,0,0,Text)
		{
			;~ MsgBox, found it
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5

	
			; PO Number 입력해서 Order ID 찾기 위한 위치로 이동해서 더블 클릭하기
			MouseClick, left, X, Y+20, 2
			
			; PO Number 입력
;			Send, % PO_Number
			
			Sleep, 700
			
			; 찾은 결과 더블 클릭해서 페이지 열기
			MouseMove, X, Y+42
			; 현재 마우스 위치에 더블클릭 합니다 (아래 세 줄 모두 현재 위치에서 더블 클릭하는 것 안 먹어서 그냥 다 해봄)
			Sleep 200
			MouseClick, left
			MouseClick, left, , , 2
			Click 2



			; 3초 동안 Find 창 닫히기 기다리기
			WinWaitClose, Find, , 3
			
			; Find 안 닫혔으면, 즉 더블 클릭해서 주문 페이지로 넘어가지 않았으면 Find 창 닫고 처음부터 다시 시작하기
			IfWinExist, Find
			{
				WinClose, Find
				LAMBS.SearchPONumber(PO_Number)
			}
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 500

			Sleep, 500
		}		

		return
	}
	
	
	
	
	
	
	
	
	
	
	
	
}


