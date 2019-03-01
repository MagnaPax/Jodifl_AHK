#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk



	ApplyCreditFunction(Invoice_No){
		
		CoordMode, mouse, Relative
		
		
		
		;LAMBS 화면 초기화
		Start()
		
		;상태바에서 알트키, 방향키 등 눌러서 Cash Receipts 찾아가기
		Send, {Alt}
		Sleep 200
		
		Send, {Right}
		Sleep 200
		
		Send, {Right}
		Sleep 200
		
		Send, {Right}
		Sleep 200
		
		Send, {Down}
		Sleep 200
		
		Send, {Down}
		Sleep 200
		
		Send, {Right}
		Sleep 200
		
		Send, {Enter}
		Sleep 2000


		;New&Clear 버튼 클릭
		MouseClick, l, 60, 120
		sleep 1500

		; OrderID 선택하기
		MouseClick, l, 145, 295

		; Invoice_No 붙여넣고 Payment Method 상태 RCredit로 바꾸기
		DragAndPast(190, 319, 30, 319, Invoice_No)
		Sleep 200
		
		Send, {Enter}
		Sleep 200
		
		Send, {Enter}
		Sleep 200
		
		Send, {Enter}
		Sleep 200
		
		Send, {Enter}
		Sleep 200
		
		Send, RCredit
		Sleep 200
		
		Send, {Enter}
		Sleep 200


		; 주황색 [Credit] 체크박스 체크하기
		MouseClick, l, 665, 892
		Sleep 200
		
		MouseClick, l, 665, 873
		Sleep 200		


/*
		; 파란색 [Invoice] 체크박스 체크하기
		; 제일 아래 3번째 체크박스 위치라고 추정되는 곳
		MouseClick, l, 665, 324
		Sleep 200
		
		; 파란색 [Invoice] 체크박스 체크하기
		; 두 번째 체크박스 위치라고 추정됨
		MouseClick, l, 665, 304
		Sleep 200		

		; 이게 제일 위에 있는 체크박스
		; 파란색 [Invoice] 체크박스 체크하기
		MouseClick, l, 665, 284
		Sleep 200
*/

		; 파랑색 [Credit] 체크박스 체크하기
		Loop, 3{
			
			Clipboard := 
				
			; 맨 위의 Order Id 값 얻기
			MouseMove, 830, 276
			SendInput, #q
			Sleep 100
			MouseMove, 884, 295
			Sleep 100
			SendInput, #q
			Sleep 100
				
			; Capture2Text 창 닫기
			IfWinExist, Capture2Text - OCR Text
				WinClose


			OrderId := % Clipboard
				
				
			; 파란색 창의 OrderId 값과 Invoice_No가 같으면 첫 번째 체크박스 클릭
			if(OrderId == Invoice_No){
				WinActivate, LAMBS
				MouseClick, l, 665, 284
				break
			}
			else{
				WinActivate, LAMBS

				Clipboard := 
					
				; 두 번째 칸의 Order Id 값 얻기
				MouseMove, 830, 296
				SendInput, #q
				Sleep 100
				MouseMove, 884, 315
				Sleep 100
				SendInput, #q
				Sleep 100
					
				; Capture2Text 창 닫기
				IfWinExist, Capture2Text - OCR Text
					WinClose


				OrderId := % Clipboard
				
				; 파란색 창의 OrderId 값과 Invoice_No가 같으면 두 번째 체크박스 클릭
				if(OrderId == Invoice_No){
					WinActivate, LAMBS
					MouseClick, l, 665, 304
					break
				}
					
				; 세 번째 칸의 Order Id 값 얻기
				MouseMove, 830, 316
				SendInput, #q
				Sleep 100
				MouseMove, 884, 335
				Sleep 100
				SendInput, #q
				Sleep 100
					
				; Capture2Text 창 닫기
				IfWinExist, Capture2Text - OCR Text
					WinClose


				OrderId := % Clipboard
				
				; 파란색 창의 OrderId 값과 Invoice_No가 같으면 세 번째 체크박스 클릭
				if(OrderId == Invoice_No){
					WinActivate, LAMBS
					MouseClick, l, 665, 324
					break					
				}				
			}			
		}
		
;		MsgBox, %OrderId%     %Invoice_No%
		/*
		IfNotEqual, %OrderId%, %Invoice_No%
		{
			MsgBox, Order Id 와 Invoice_No가 맞지 않습니다
			Reload
		}
		*/
		if(%OrderId% != %Invoice_No%){
			MsgBox, Order Id 와 Invoice_No가 맞지 않습니다
			Reload			
		}
	
	
		Sleep 200


		
		; 저장 하기
		Send, {F9}
		Sleep 1000

		;Save successfully 나오길 기다렸다 엔터
		Sleep 2000
		Send, {Enter}


	;Create Invoice로 돌아가기
	OpenCreateInvoiceTab()
	
	; security 버튼 풀기
	FindSecurityButtonAndClickItThenInputNumber1()	

/*
	; refresh 버튼 더블 클릭하기
	ControlClick, x123 y125, LAMBS -  Garment Manufacturer & Wholesale Software
	Sleep 200
	ControlClick, x123 y125, LAMBS -  Garment Manufacturer & Wholesale Software
*/

	; refresh 버튼 더블 클릭하기
	MouseClick, l, 123, 125, , D
	
	
	; security 버튼 풀기
	FindSecurityButtonAndClickItThenInputNumber1()		

		
		
		
		
		
		
	return
	}
