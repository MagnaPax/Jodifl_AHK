#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;	ControlGetText, SubTotal, WindowsForms10.EDIT.app.0.378734a401, LAMBS
;	ControlGetText, SubTotal, , LAMBS
;	ControlClick, x41 y265, LAMBS
;	70829

/*
	;Invoice 입력하기 후 엔터
	ControlSetText, WindowsForms10.EDIT.app.0.378734a512, %Invoice_No%
	ControlSend, WindowsForms10.EDIT.app.0.378734a512, {Enter}
*/



1stCommonLAMBSProcessing(Invoice_No, Invoice_Wts, No_of_Boxes, ApplyCredit, Consolidation, CustomerUPSAccount, NextDay, 2ndDay, 3rdDay, Qty_of_Items){
	
	
	
;	No_of_Boxes := ExtractWtsofBoxes(Invoice_Wts)

	; Invoice_Wts값에 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 Wts_of_Boxes에 저장
	StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.

	; Wts_of_Boxes를 루프 돌려서 들어있는 값 개수만큼 No_of_Boxes에 저장
	Loop, %Wts_of_Boxes0%
	{
		
	;	this_Wts := Wts_of_Boxes%a_index%
	;	MsgBox, The Weight Box number %a_index% is %this_Wts%
		No_of_Boxes := % A_Index
	}
	
	
	;LAMBS Window 활성화 후 Hide All 클릭해서 메뉴 바 없애기
	Start()
	
	;LAMBS Invoice 1 탭 클릭
	Start_Invoice_1()
	
	;창 크기 최대화
	WinMaximize, LAMBS

	;Create Invoice로 돌아가기
	OpenCreateInvoiceTab()

	sleep 1000


	;New & Clear 버튼 클릭
	MouseClick, l, 60, 124, 2, 
	sleep 1000
	
	;Order ID 라디오 버튼 선택하기
	MouseClick, l, 37, 205, 2, 
	sleep 500
	
	;Invoice 입력하기
	MouseClick, l, 209, 205, 2, 
	SendInput, %Invoice_No%
	SendInput, {Enter}
	sleep 1000
	
	
	; [예외처리] Invoice 번호 잘못 입력하여 경고창 나왔을 때 엔터 누른 뒤 다시 시작하기
	IfWinExist, Confirm
	{
		SendInput, {Enter}
		reload
	}
	
	;Security 버튼 있는지 찾기
	;그림이 찾아졌으면 클릭
	;비밀번호 입력칸에 1 입력 후 엔터
	FindSecurityButtonAndClickItThenInputNumber1()


	; Invoice 1 탭에서 시작하기 위해 클릭
	MouseClick, l, 45, 263, 1

	;Customer 클릭해서 정보 갱신하기
	MouseClick, l, 306, 297
	MouseClick, l, 270, 297
	SendInput, {Enter}
	sleep 1000


	;무게 입력하기
	MouseClickDrag, l, 591, 476, 363, 476
	sleep 500
	SendInput, %Invoice_Wts%
	SendInput, {Enter}


	;상자 갯수 입력하기
	MouseClickDrag, l, 591, 500, 363, 500
	sleep 500
	SendInput, %No_of_Boxes%
	SendInput, {Enter}


	;인보이스 메모 처리
	;맨 아랫줄 끝으로 가서 TRACKING # 입력
	Clipboard := 
;	MsgBox, Clipboard`n%Clipboard%
	SendInput, {down 50}
	SendInput, {End}
	Send, {SPACE}TRACKING{#}{SPACE}


	;인보이스 메모 읽어오기
	MouseMove, 384, 540
	MouseGetPos, , , , control
	ControlGetText, InvoiceMemoOnLAMBS, %control%, LAMBS
	
	
	;Customer Memo 읽어오기
	MouseMove, 358, 624
	MouseGetPos, , , , control
	ControlGetText, CustomerMemoOnLAMBS, %control%, LAMBS



	jpgLocation = %A_ScripDir%PICTURES\count.png
	QtyofItemsOnLAMBS := PicSearchAndReturnItsVal(jpgLocation)
/*	
	; LAMBS의 아이템 개수 값 읽기
	MouseMove, 50, 596
	MouseGetPos, , , , control
	ControlGetText, QtyofItemsOnLAMBS, %control%, LAMBS
	MsgBox, % QtyofItemsOnLAMBS
*/

	if(Qty_of_Items != QtyofItemsOnLAMBS){
	
		;아이템 갯수 확인 위해 Details Edit열기
		SendInput, !d

		;Details Edit창이 없어질때까지 대기
		CheckTheWindowPresentAndWaitUntillItClose("Details Edit")
		Sleep 500
		
		;저장 후 Status창이 나타나면 없어질때까지 대기
		IfWinExist, Status
			CheckTheWindowPresentAndWaitUntillItClose("Status")


		Sleep 500

		;Security 버튼 있는지 찾기
		;그림이 찾아졌으면 클릭
		;비밀번호 입력칸에 1 입력 후 엔터
		FindSecurityButtonAndClickItThenInputNumber1()
		
		
		;아이템을 추가할 것인지 묻고 Yes 눌렀으면 Sales Orders 버튼 클릭
		MsgBox, 260, Add Items to, Would you like to transfer Items from Sales Order?
		IfMsgBox, Yes
		{
			MouseClick, l, 232, 388
			
			;Transfer from Sales Order창이 없어질때까지 대기
			CheckTheWindowPresentAndWaitUntillItClose("Transfer from Sales Order")
			
			;Security 버튼 있는지 찾기
			;그림이 찾아졌으면 클릭
			;비밀번호 입력칸에 1 입력 후 엔터
			FindSecurityButtonAndClickItThenInputNumber1()
		}	
	}
*/


/*
		;아이템 갯수 확인 위해 Details Edit열기
		SendInput, !d

		;Details Edit창이 없어질때까지 대기
		CheckTheWindowPresentAndWaitUntillItClose("Details Edit")
		Sleep 500
		
		;저장 후 Status창이 나타나면 없어질때까지 대기
		IfWinExist, Status
			CheckTheWindowPresentAndWaitUntillItClose("Status")


		Sleep 500

		;Security 버튼 있는지 찾기
		;그림이 찾아졌으면 클릭
		;비밀번호 입력칸에 1 입력 후 엔터
		FindSecurityButtonAndClickItThenInputNumber1()
		
		
		;아이템을 추가할 것인지 묻고 Yes 눌렀으면 Sales Orders 버튼 클릭
		MsgBox, 260, Add Items to, Would you like to transfer Items from Sales Order?
		IfMsgBox, Yes
		{
			MouseClick, l, 232, 388
			
			;Transfer from Sales Order창이 없어질때까지 대기
			CheckTheWindowPresentAndWaitUntillItClose("Transfer from Sales Order")
			
			;Security 버튼 있는지 찾기
			;그림이 찾아졌으면 클릭
			;비밀번호 입력칸에 1 입력 후 엔터
			FindSecurityButtonAndClickItThenInputNumber1()
		}			
*/

	WinRestore, LAMBS

	
	
	
	
	
	
	
	
	
	
	
	
	return No_of_Boxes
}