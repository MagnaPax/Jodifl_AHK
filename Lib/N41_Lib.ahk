#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include <N41>
;~ #Include N41.ahk
;~ #Include FindText_02FunctionONLY.ahk
#Include <FindText_02FunctionONLY>


;~ MsgBox
;~ findTheStyleOnSalesOrderTab(Style#, StyleColor, OrderQty)


; SO 번호 받아서 Sales Order 창 열기
openSO(SO#){
	

	
	N_driver := new N41
	
	; 사용자의 마우스 이동 막음
;	BlockInput, MouseMove

	; N41 활성화 시키기
	WinActivate, ahk_class FNWND3126
	
	Sleep 150


	; 혹시 이전에 열린 창이 있으면 닫기
	IfWinExist, Customer P.O or S.O #
		WinClose

	Sleep 150


	; Sales Order 클릭
	N_driver.ClickSalesOrderOnTheMenuBar()


	; Sales Order 에 있는 리프레쉬 버튼 클릭하기
	N_driver.ClickRefreshButtonOnSalesOrder()


	; Customer P.O or S.O # 창 뜰 때까지 기다리기
	WinWaitActive, Customer P.O or S.O #
	Sleep 150

	; 열린 Customer P.O or S.O # 창에 SO 번호 입력하기
	SendInput, % SO#
	Sleep 150			
	Send, {Enter}
	Sleep 150


	; 고객 코드 얻기
	; 특별한 의미는 없고 창이 열린 뒤 안정화 됐는지 확인하기 위해.
	Loop
	{					
		; Sale Order 에 있는 고객 코드 얻기
		CustCode := N_driver.GetCustomerCode()
				
;		MsgBox, % "CustCode : " . CustCode
				
		;값이 들어있으면 루프 빠져나가기
		;값이 들어있지 않으면 계속 루프 돌기
		if(CustCode)
			break
	}
	
	Sleep 300
	
	return
}










; Desc. 의 X 와 Y 위치 얻기
getXandYcoordinatesOf_Desc(){
	
	t1:=A_TickCount
	Text:="|<Desc.>*176$26.w0008U002AtssXHGH8rr42B0R0WHnFD3bXa"
	
	X_Arr := object()
	Y_Arr := object()
	
	X_Arr := []
	Y_Arr := []
	
	; 가장 처음의 Desc. 찾기
	if (ok:=02_FindText(232-150000//2, 398-150000//2, 150000, 150000, 0, 0, Text))
	{
		CoordMode, Mouse
		X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
		; Click, %X%, %Y%
		MouseMove, X, Y
;		MsgBox, % "X : " . X . "`nY : " . Y . "`n`nW : " . W . "`n" . "H : " . H . "`n`n1"
		X_Arr.Insert(X)
		Y_Arr.Insert(Y)
		
		preX := X
		preY := Y
		Sleep 100
			  
		; 처음 밑의 두 번째 Desc.부터 7개 찾기(첫 번째 포함 8개 찾게 되는 셈)
		; 화면에 Desc. 가 2개만 있으면 2개만 읽고 5개만 있으면 5개만 읽는다. 1개만 있으면 위의 가장 처음만 읽고 아래는 실행되지 않는다
		Loop, 7{
				
		if (ok:=02_FindText(X-20, Y, 150000, 150000, 0, 0, Text))
		{
			CoordMode, Mouse
			X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
			; Click, %X%, %Y%
			MouseMove, X, Y

;	        MsgBox, % "X : " . X . "`nY : " . Y . "`n`nW : " . W . "`n" . "H : " . H . "`n`n" . A_Index+1
			X_Arr.Insert(X)
			Y_Arr.Insert(Y)
			Sleep 100
		}
				

		} ; end끝 - Loop, 7
	} ; end끝 - if (ok:=02_FindText(232-150000//2, 398-150000//2, 150000, 150000, 0, 0, Text)	
	
	
	return [X_Arr, Y_Arr]
}
















/*
; 스타일 번호, 색깔, 수량이 맞는 곳의 Desc. X와 Y 좌표 얻기
; DropAll 체크됐으면 번호만 찾아서 모두 다 내리기
findTheStyleOnSalesOrderTab(Style#, StyleColor, OrderQty, DropAll){
MsgBox, 262144, Title, pause
	
	t1:=A_TickCount

	Text:="|<Desc.>*176$26.w0008U002AtssXHGH8rr42B0R0WHnFD3bXa"

	; ############ ########################################################################################################################
	; 이렇게 10번 돌지 말고 화면 Desc 찾은 뒤 이전 값과 똑같은 값이 있다면 멈추고 나갈까? 중복되게 주문할 수도 있나?
	Loop, 10{
		
	  
		; 가장 처음의 Desc. 찾기
		if (ok:=02_FindText(232-150000//2, 398-150000//2, 150000, 150000, 0, 0, Text))
		{
			CoordMode, Mouse
			X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
			; Click, %X%, %Y%
			MouseMove, X, Y
			;~ subFunctionOf_findTheStyleOnSalesOrderTab_ToDropIfItIsMatched(X, Y, Style#, StyleColor, OrderQty, DropAll)
;			MsgBox, % "X : " . X . "`nY : " . Y . "`n`nW : " . W . "`n" . "H : " . H . "`n`n1"
			Sleep 500
			  
			; 처음 밑의 두 번째 Desc.부터 7개 찾기(첫 번째 포함 8개 찾게 되는 셈)
			Loop, 7{
				
			if (ok:=02_FindText(X-20, Y, 150000, 150000, 0, 0, Text))
			{
				CoordMode, Mouse
				X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
				; Click, %X%, %Y%
				MouseMove, X, Y
;		        MsgBox, % "X : " . X . "`nY : " . Y . "`n`nW : " . W . "`n" . "H : " . H . "`n`n" . A_Index+1
			}
				

			} ; end끝 - Loop, 7
		} ; end끝 - if (ok:=02_FindText(232-150000//2, 398-150000//2, 150000, 150000, 0, 0, Text)


		; 못 찾았으면 화살표 클릭하기
		loop, 8{
			DownArrow_Text:="|<>*199$16.UUW368CsUT20s810U02"

			if (ok:=02_FindText(1734-150000//2, 983-150000//2, 150000, 150000, 0, 0, DownArrow_Text))
			{
				CoordMode, Mouse
				;~ X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
				X_DownArrow:=ok.1.1, Y_DownArrow:=ok.1.2, W_DownArrow:=ok.1.3, H_DownArrow:=ok.1.4, Comment_DownArrow:=ok.1.5, X_DownArrow+=W//2, Y_DownArrow+=H//2

;			    MsgBox, 262144, Title, 화살표 찾았음 클릭 예정
				;~ Click, %X%, %Y%
				Click, %X_DownArrow%, %Y_DownArrow%
				Sleep 150
			}
		} ; end끝 - loop, 8
	} ; end끝 - Loop, 10
		
	
	
	MsgBox, 262144, NO ITEM ERROR, NO ITEM ON THIS SCREEN FOUND`n`n스타일 : %Style#%`n색깔 : %StyleColor%`n이 현재 화면에 없습니다.
	
	
	
	
	
	
	
	return
}
*/



; Desc 위치를 기준 삼아서 스타일번호, 색깔, 수량을 읽은 뒤 드롭해야 되는 스타일이면 드롭하기
; X, Y : Desc 위치
checkValuesAndDropItems(X, Y, Style#, StyleColor, OrderQty, DropAll){
	
;MsgBox, 262144, title, 메소드 시작
	
	; 변수 초기화
	Clipboard := ""
	Style#_onTheScreen := ""
	Color_onTheScreen := ""
	OrderQTY_onTheScreen := ""
	theItemFound := ""
	
	
	; Style 값 얻기
	MouseMove, X+50, Y-18
	Sleep 150

	; Style값 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
	Send, {RButton}
	Loop, 4
	{
		Sleep 150
		Send, {Down}
		;~ Sleep 150
	  }
	Send, {Enter}
	Sleep 150

	Style#_onTheScreen := Clipboard	
	Sleep 150
	  
	  
	  
	; Color 값 얻기
	MouseMove, X+150, Y-18
	Sleep 150

	; Style값 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
	Send, {RButton}
	Loop, 4
	{
		Sleep 150
		Send, {Down}
		;~ Sleep 150
	}
	Send, {Enter}
	Sleep 150

	Color_onTheScreen := Clipboard	
	Sleep 150  




	; Order QTY 값 얻기
	MouseMove, X+780, Y
	Sleep 150

	; Style값 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
	Send, {RButton}
	Loop, 4
	{
		Sleep 150
		Send, {Down}
		;~ Sleep 150
	}
	Send, {Enter}
	Sleep 150

	OrderQTY_onTheScreen := Clipboard	
	Sleep 150





	; 값을 못 읽은게 하나라도 있으면 재귀호출해서 다시 읽기
	if(!Style#_onTheScreen||!Color_onTheScreen||!OrderQTY_onTheScreen){
;		MsgBox, 262144, no data error, 값이 없는 변수가 있기 때문에 재귀호출합니다.
		checkValuesAndDropItems(X, Y, Style#, StyleColor, OrderQty, DropAll)
	}
	  
	
	; 화면에서 읽어온 값과 인자로 받은 값이 모두 같다면 아이템 캔슬하기		
	if(Style# == Style#_onTheScreen && StyleColor == Color_onTheScreen && OrderQty == OrderQTY_onTheScreen){
		MsgBox, 262144, correct drop item, 스타일번호, 색깔, 수량이 모두 같습니다. 이 아이템을 드롭합니다.
		dropThisStyle(X, Y)
		theItemFound = 1 ; 아이템이 발견된 것 표시
	}
	
	; Drop All 선택했다면 스타일만 입력값만 맞으면 아이템 캔슬하기
	else if(DropAll && Style# == Style#_onTheScreen){
		MsgBox, 262144, correct drop item, Drop All 이 선택되었음. 아이템을 드롭합니다.
		dropThisStyle(X, Y)
	}	
	

	return [Style#_onTheScreen, Color_onTheScreen, OrderQTY_onTheScreen]
}
	



; 인자로 받은 X, Y 를 기준삼아서 아이템 드롭하기
dropThisStyle(X, Y){
	
	MouseMove, X+1150, Y-18
	Sleep 150

	Click
	Send, Out of Stock{Enter}
	
	Sleep 300
	
	
	
	return
}




; 아래쪽 화살표 8번 클릭하기
clickDownArrow8Times(){
	
	Text:="|<>*199$16.UUW368CsUT20s810U02"
	
	; 화살표 8번 클릭하기
	loop, 8{
		
		if (ok:=02_FindText(232-150000//2, 398-150000//2, 150000, 150000, 0, 0, Text))
		{
			CoordMode, Mouse
			X:=ok.1.1, Y:=ok.1.2, W:=ok.1.3, H:=ok.1.4, Comment:=ok.1.5, X+=W//2, Y+=H//2
			Click, %X%, %Y%
			;MouseMove, X, Y
			Sleep 300
		}
		else
			return 0  ; 화살표를 못 찾았으면 0 리턴함
	} ; end끝 - loop, 8	
	
	
	return 1 
}







