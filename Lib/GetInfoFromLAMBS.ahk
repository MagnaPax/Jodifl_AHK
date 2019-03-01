#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk



	
	;LAMBS에서 고객 정보 찾아서 저장
	GetInfoFromLAMBS(){
		
		; 마우스를 LAMBS화면 기준으로 세팅
		CoordMode, mouse, relative
		
		; 마우스 커서는 사용자의 물리적 마우스 이동에 반응하지 않습니다
		BlockInput, MouseMove
		
		
		

;		WinActivate, LAMBS
		
		Start()
		Start_Invoice_1()		
		

		;MouseMove, 220, 596
		DllCall("SetCursorPos", int, 60-8, int, 601-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 60, 601
		Sleep 100
		MouseGetPos, , , , control
		ControlGetText, CompanyName, %control%, LAMBS

/*
		MouseClick, l, 189, 601
		Send, ^a^c
		;Send, {Esc}
		SendInput, {Enter}
		Sleep 700
		CompanyName := Clipboard
*/		

		;MouseMove, 200, 626
		DllCall("SetCursorPos", int, 310-8, int, 625-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 310, 625
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, Attention, %control%, LAMBS
/*		
		MouseClick, l, 81, 626
		Send, ^a^c
		Sleep 700
		Attention := Clipboard		
*/		
		DllCall("SetCursorPos", int, 300-8, int, 646-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 300, 646
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, Address1, %control%, LAMBS

		DllCall("SetCursorPos", int, 200-8, int, 669-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 200, 669
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, Address2, %control%, LAMBS

		DllCall("SetCursorPos", int, 100-8, int, 691-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 100, 691
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, ZipCode, %control%, LAMBS

		DllCall("SetCursorPos", int, 317-8, int, 690-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 317, 690
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, State, %control%, LAMBS
 		
		DllCall("SetCursorPos", int, 200-8, int, 690-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 200, 690
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, City, %control%, LAMBS

		DllCall("SetCursorPos", int, 156-8, int, 712-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 156, 712
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, Phone, %control%, LAMBS
		
		DllCall("SetCursorPos", int, 857-8, int, 313-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 857, 313
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, SubTotal, %control%, LAMBS
		
		
		;Billing Add1
		DllCall("SetCursorPos", int, 201-8, int, 494-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 201, 494
		Sleep 100		
		MouseGetPos, , , , control
		ControlGetText, BillingAdd1, %control%, LAMBS
/*
		;Billing Zip Code
		DllCall("SetCursorPos", int, 62-8, int, 531-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 62, 531
		MouseGetPos, , , , control
		ControlGetText, BillingZip, %control%, LAMBS	
*/
		MouseClick, l, 62, 531
		Send, ^a^c
		Sleep 700
		BillingZip := Clipboard
		
		
		
;Country, United States

	
		; Invoice 2 탭 클릭하기
		MouseClick, l, 101, 266, LAMBS
		Sleep 100
/*		
		DllCall("SetCursorPos", int, 80-8, int, 742-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;		MouseMove, 80, 742
		MouseGetPos, , , , control
		ControlGetText, Email, %control%, LAMBS
*/		
		MouseClick, l, 80, 742
		Send, ^a^c
		Sleep 700
		Email := Clipboard
				
		

		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
		DeleteAdd2InAdd1(Address1)
		
		; Invoice 1 탭 클릭하기
		MouseClick,l , 41, 265



		; 사용자에게 마우스 커서를 이동하도록 허용합니다.
		BlockInput, MouseMoveOff



;		MsgBox, %CompanyName%`n%Attention%`n%Address1%`n%Address2%`n%ZipCode%`n%City%`n%Phone%`n%SubTotal%
;		MsgBox, 이름 : %CompanyName%`n`n수령인 : %Attention%`n`n주소1 : %Address1%`n`n주소2 : %Address2%`n`n우편번호 : %ZipCode%`n`n도시명 : %City%`n`n전번 : %Phone%`n`n가격 : %SubTotal%`n`n이멜 : %Email%`n`n가격 : %SubTotal%`n`n`n청구소주소 :  %BillingAdd1%`n`n`n청구소ZipCode:  %BillingZip%
	

		
		




	return
	}
	
	
	
	
	
	
	
	
	
	
	
	
	

	;웹에서 온 주문이 아니기 때문에 LAMBS에서 고객 정보 찾아서 저장
	1GetInfoFromLAMBS(){

;		WinActivate, LAMBS
		
		; Hide All 눌러서 메뉴바 없애기
;		ControlClick, x213 y65, LAMBS
		; Invoice 1 탭 클릭하기
;		ControlClick, x41 y265, LAMBS

		ControlGetText, CompanyName, WindowsForms10.EDIT.app.0.378734a458, LAMBS
		ControlGetText, Attention, WindowsForms10.EDIT.app.0.378734a450, LAMBS
		ControlGetText, Address1, WindowsForms10.EDIT.app.0.378734a457, LAMBS
		ControlGetText, Address2, WindowsForms10.STATIC.app.0.378734a409, LAMBS
		ControlGetText, ZipCode, WindowsForms10.EDIT.app.0.378734a455, LAMBS
		ControlGetText, City, WindowsForms10.EDIT.app.0.378734a485, LAMBS
		ControlGetText, Phone, WindowsForms10.EDIT.app.0.378734a452, LAMBS
		ControlGetText, SubTotal, WindowsForms10.EDIT.app.0.378734a401, LAMBS
		
		; Invoice 2 탭 클릭하기
;		ControlClick, x101 y266, LAMBS
		
		ControlGetText, Email, WindowsForms10.EDIT.app.0.378734a484, LAMBS

		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; Invoice 1 탭 클릭하기
;		ControlClick, x41 y265, LAMBS
		
;		MsgBox, %CompanyName%`n%Attention%`n%Address1%`n%Address2%`n%ZipCode%`n%City%`n%Phone%`n%SubTotal%
		MsgBox, 이름: %CompanyName%`n수령인:%Attention%`n주소1:%Address1%`n주소2%Address2%`n우편번호: %ZipCode%`n도시명: %City%`n전번: %Phone%`n가격: %SubTotal%`n이멜: %Email%		

		

	
	return
	}
	
	
	;웹에서 온 주문이 아니기 때문에 LAMBS에서 고객 정보 찾아서 저장
	11GetInfoFromLAMBS(){
		
		;화면 초기화
		Start()
		Start_Invoice_1()


		CompanyName := ClickAndCtrlAll(218, 598)
		Attention := ClickAndCtrlAll(273, 621) ;DragAndCopy(513, 621, 225, 621) ;fromX, fromY, toX, toY)
		Address1 := ClickAndCtrlAll(273, 642) ;DragAndCopy(513, 642, 225, 642) ;fromX, fromY, toX, toY)
		Address2 := ClickAndCtrlAll(273, 665) ;DragAndCopy(513, 665, 225, 665) ;fromX, fromY, toX, toY)
		ZipCode := ClickAndCtrlAll(61, 687) ;DragAndCopy(301, 687, 225, 687) ;fromX, fromY, toX, toY)
		City := ClickAndCtrlAll(152, 688)
		;State := ClickAndCtrlAll(501, 691)
		Phone := ClickAndCtrlAll(92, 710)
		SubTotal := ClickAndCtrlAll(814, 310)

		; Invoice 2 탭으로 이동하기
		MouseClick, l, 101, 266, 1
		Email := ClickAndCtrlAll(273, 739)

		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; 다시 Invoice 1 탭으로 돌려놓기
		MouseClick, l, 41, 265, 1
		
		MsgBox, 이름: %CompanyName%`n수령인:%Attention%`n주소1:%Address1%`n주소2%Address2%`n우편번호: %ZipCode%`n도시명: %City%`n전번: %Phone%`n가격: %SubTotal%`n이멜: %Email%		


		
	return
	}
		
	
	