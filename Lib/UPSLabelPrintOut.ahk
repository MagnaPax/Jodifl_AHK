#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk

/*
1UPSLabelPrintOut(){
	CoordMode, mouse, Relative
	
	WinActivate, UPS WorldShip - Remote Workstation
	
	;F10 Click
	MouseMove, 575, 627, 50
	click			
	Send, !{tab}


	;F2, F3 Click
;	Send, !{tab}	
	MouseMove, 37, 84, 50
	click
	Send, !{tab}	
	Sleep 5000
	Send, !{tab}	
	
	MouseMove, 127, 250, 50
	MouseClick, l, 127, 250, , , d
	Send, {enter}
	Click
	Sleep 100
	
	Click
	
	MouseMove, 127, 264, 50
	Click
	Sleep 100	
	Click			


	return
}
*/


22UPSLabelPrintOut(){
	
	windowtitle = UPS WorldShip - Remote Workstation
;	windowtitle = UPS WorldShip



	; F10버튼 클릭해서 UPS 라볠 인쇄하기
	;Control_SnedButton(F10, windowtitle) ;이걸로 한 번 해보자
	
	ControlSend, Button10, {Enter}, %windowtitle%
	Sleep 1000



	; UPS WorldShip 에 경고창 나오면 없애기
	Loop, 5
	{
		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip
		Sleep 200
	}
	
	
	


;	IfWinExist, UPS WorldShip
;		WinClose

	
/*
;; 이것도 됐었는데	ControlClick, Button10, %windowtitle%
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;	Sleep 3000

	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	if(WinWait, UPS WorldShip, , 1){
		WinActivate
		ignoreUPSWorldShip()
	}
	if(WinWait, UPS WorldShip, , 1){
		WinActivate
		ignoreUPSWorldShip()
	}
	
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	ControlSend, Button1, {Enter}, UPS WorldShip
*/

/*	
;	WinWait, UPS WorldShip, , 1
	Sleep 1000
	ignoreUPSWorldShip()
;	WinWait, UPS WorldShip, , 1
	Sleep 1000
	ignoreUPSWorldShip()	
;	WinWait, UPS WorldShip, , 1
	Sleep 1000
	ignoreUPSWorldShip()
*/
/*	
	; F3 눌러서 History로 넘어가기
	Sleep 1500
	Control_SnedButton("F3", windowtitle)
	

	; F3 눌러서 History로 넘어가기 
	Sleep 1000
	Control_SnedButton(F3, windowtitle)
	
	; History 버튼 클릭해서 넘어가기
	Sleep 1000
	ControlClick, Afx:RibbonBar:11e0000:8:10005:101, %windowtitle%	
	
	; F3 클릭해서 History로 넘어가기
	ControlClick, x36 y87, %windowtitle%	
	Sleep 3000	
	
*/





; 	이것도 작동됨 ******************** ControlSend, , {F3}, UPS WorldShip - Remote Workstation
/*
	windowtitle = UPS WorldShip - Remote Workstation
	ControlSend, , {F3}, %windowtitle%
	Sleep 2000
	
	;UPS WorldShip 경고창이 뜨면 비활성화로 Yes 버튼 누르기
	ControlClick, Button1, %windowtitle%
;	ControlSend, Button1, {Enter}, %windowtitle%
*/






	; F3 클릭해서 History로 넘어가기	
	Loop, 3
	{
		ControlSend, , {F3}, UPS WorldShip - Remote Workstation
		Sleep 200
		key = F3
		Control_SnedButton(Key, windowtitle)
		Sleep 200
		Control_SnedButton(F3, windowtitle)
		Sleep 200

		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip
	}
	









/*
	ControlSend, , {F3}, %windowtitle%
	
	
	ControlSend, , {Right}, %windowtitle%
	Sleep 200
	ControlSend, , {Down}, %windowtitle%
	Sleep 200	
	ControlSend, , {Right}, %windowtitle%
*/	

/*
	; F3 눌러서 History로 넘어가기
;	key = F3
	Control_SnedButton(F3, windowtitle)
	Sleep 2000
*/	
/*
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	WinWait, UPS WorldShip, , 1
	ignoreUPSWorldShip()
*/	
	
/*	
	Control_InputText(TextLoc, Contents, windowtitle){
	
	Control_SnedButton(Right, windowtitle) 
	Sleep 200
	Control_SnedButton(Down, windowtitle)
	Sleep 200
	Control_SnedButton(Right, windowtitle)
	Sleep 200
*/




/*
	;UPS WorldShip 경고창이 뜨면 비활성화로 그냥 무시하고 엔터치기
	WinWait, UPS WorldShip, , 3
	ignoreUPSWorldShip()
	WinWait, UPS WorldShip, , 3
	ignoreUPSWorldShip()
	WinWait, UPS WorldShip, , 3
	ignoreUPSWorldShip()
	WinWait, UPS WorldShip, , 3
	ignoreUPSWorldShip()
	WinWait, UPS WorldShip, , 3
	ignoreUPSWorldShip()
	WinWait, UPS WorldShip, , 3
	ignoreUPSWorldShip()
*/	
;	WinActivate, UPS WorldShip - Remote Workstation
;	SendInput, F3
	
	
/*

*/	

/*
	;History 화면의 Weekday's Pending Pickup 클릭해서 활성화하기
	ControlClick, Edit8, %windowtitle%
	Sleep 500	
*/


/*	
	ControlSend, , {Up}, %windowtitle%
	ControlSend, , {Up}, %windowtitle%

	ControlSend, , {Right}, %windowtitle%
	ControlSend, , {Down}, %windowtitle%
	ControlSend, , {Right}, %windowtitle%
	ControlSend, , {Down}, %windowtitle%
*/	

/*	
	Send, {up}
	Send, {up}
	
*/	

	Sleep 3000
	
	
	;쪼그만 화살표 두 개 눌러서 배송 내역 목록표 확장
	ControlClick, x50 y246, %windowtitle%
	Sleep 200
	ControlClick, x70 y263, %windowtitle%
	Sleep 200
	
	
/*	
	ControlSetText, Edit8, Enter, %windowtitle%
	ControlSetText, Edit8, Down, %windowtitle%
	ControlSetText, Edit8, Enter, %windowtitle%
	ControlSetText, Edit8, Down, %windowtitle%
*/	
/*
	; Tracking Number 얻기 위해 목룍표 이동하기
	Control_SnedButton("Right", windowtitle)
	Control_SnedButton("Down", windowtitle)
	Control_SnedButton("Right", windowtitle)
	Control_SnedButton("Down", windowtitle)
*/	
/*
	Send, {Right}
	Send, {Down}
	Send, {Right}
	Send, {Down}
*/	



	; Tracking Number 넣기 위한 GUI
	Gui, Show, w330 h88, Tracking Number

	Gui, Add, Text, x12 y25 w60 h30 +Center, Tracking Number
	Gui, Add, Edit, x72 y28 w180 h20 vTrackingNumber
;	Gui, Add. Button, x92 y28 w30 h30, ASDF

	Gui, Add, Button, x260 y28 w60 h20, Enter_


	;GUI시작 시 포커스를 Invoice_No 입력칸에 위치
	GuiControl, Focus, TrackingNumber
	
	; 창이 종료될때까지 대기
	WinWaitClose, Tracking Number
	
	;Enter 버튼 누르면 Tracking Number GUI창 종료
	ButtonEnter_:
	WinClose, Tracking Number
	
	; GUI에서 입력받은 값을 사용하기 위해
	Gui, Submit, nohide
;	GUI, TrackingNumber:Destroy
	GUI, Hide
	

	
;	WinClose, Tracking Number

	
;	GuiClose:
;	ExitApp



	;UPS 배송 창으로 넘어감으로 초기화 하면서 함수 종료
;	Control_SnedButton("F2", windowtitle)
	Loop, 3
		Control_SnedButton(F2, windowtitle)

	
	
;	Enter::
;	WinClose, Tracking Number
	

;	GuiClose:
;	ExitApp		

	

	return
}



;A_ScreenWidth = 1920 화면 가로넓이
;A_ScreenHeight = 1082 화면 높이


UPSLabelPrintOut(){
	
	; 문자열 비교가 대소문자를 구분하도록
	StringCaseSense, On
	
	; UPS 창 항상위에 설정
	Winset, AlwaysOnTop, On, UPS WorldShip - Remote Workstation
	
	Clipboard :=
	
;	MsgBox, % CompanyName


	; 마우스를 화면 기준으로 세팅
	CoordMode, mouse, screen

	WinActivate, UPS WorldShip - Remote Workstation

;	Loop, 5
;	{
;		Send, {Esc}
;	}

	; X, Y점을 화면 기준으로 잡기.
	; A_ScreenHeight 값이 1080이라 화면 구석으로 옮기기 위해 1920 빼기
	; A_ScreenWidth 값이 1920이라 살짝 오른쪽으로 옮기기 위해 190 더하기
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
		
;	MsgBox, A_ScreenWidth`n`n%A_ScreenWidth%
;	MsgBox, A_ScreenHeight`n`n%A_ScreenHeight%
		
	PointOfX -= 650
	PointOfY -= 1080
		
		
	WinMove, UPS WorldShip - Remote Workstation, , %PointOfX%, %PointOfY%, 760, %A_ScreenHeight%


	; F10버튼 누르기 위한 세팅
	Send, {f10}
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 70
	PointOfY -= 450
	MouseMove, %PointOfX%, %PointOfY%
	Click




	; UPS WorldShip 에 경고창 나오면 없애기
	Loop, 5
	{
		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip

		Sleep 100
		ControlSend, , {F2}, UPS WorldShip - Remote Workstation
		Sleep 100
		ControlSend, , {F5}, UPS WorldShip - Remote Workstation
		Sleep 100

		IfWinExist, UPS WorldShip
			ControlSend, Button1, {Enter}, UPS WorldShip
	}



	Sleep 1000


	; History 버튼 누르기 위한 세팅
	loop, 6
	{		
		PointOfX = %A_ScreenWidth%
		PointOfY = %A_ScreenHeight%
		PointOfX -= 620
		PointOfY -= 980
		MouseMove, %PointOfX%, %PointOfY%
		Click
		
		Sleep 100		
	}




	; History 화면 나올때까지 기다리기
	; UPS History 창인지 찾아내기. SHIP 아이콘이 있으면 History 창이니까 그 이미지 찾기
	; pX, pY 값은 아무거나 입력해도 됨. 어차피 ByRef로 전역변수로 선언되어 값을 받기 위함이 아닌 이미지를 찾은 위치를 돌려주기 위한 변수
	; 버튼을 찾았으면 ErrorLevel 값은 0 못 찾았으면 1
	; 버튼을 찾을때까지 계속 루프 돌리기
	jpgLocation = %A_ScripDir%PICTURES\UPS_Shipping_Button.png
	
	Loop
		;PicSearch(jpgLocation, 1, 1, 1100, 20, 1600, 300)
		PicSearch(jpgLocation, 1, 1, 0, 0, 1000, 1000)
	until ErrorLevel == 0
	
	Sleep 200


	; History 화면 나올때까지 기다리기
;	Sleep 7000
	
	

	; 빈 공간 클릭하기
	loop, 4
	{
		PointOfX = %A_ScreenWidth%
		PointOfY = %A_ScreenHeight%
		PointOfX -= 220
		PointOfY -= 680
		MouseMove, %PointOfX%, %PointOfY%
		Click
	}



	; 첫번째 작은 십자가 아이콘 클릭
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 600
	PointOfY -= 832
	MouseMove, %PointOfX%, %PointOfY%
	Click
	
	; 두번째 작은 십자가 아이콘 클릭
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 581
	PointOfY -= 815
	MouseMove, %PointOfX%, %PointOfY%
	Click



	

/*
	; 상태바를 오른쪽으로 끌기
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 554
	PointOfY -= 515
	
	PointOftoX = %A_ScreenWidth%
	PointOftoY = %A_ScreenHeight%
	PointOftoX -= 450
	PointOftoY -= 515
	
	MouseClickDrag, l, %PointOfX%, %PointOfY%, %PointOftoX%, %PointOftoY%, 5
*/








	; CompanyName 값이 ComparingCompanyName 보다 길면 인식을 못하더라
	; 그래서 CompanyName 값을 잘라서 ModifiedCompanyName 에 넣었다
	ModifiedCompanyName := SubStr(CompanyName, 1, 16) ;, 19)


	; ModifiedCompanyName 길이는 위의 SubStr 함수로 인해 길이가 16이 되었는데
	; 만약 회사 이름이 16 글자보다 짧으면 그 안에 공란(Space)도 들어가게 되는 것 같음
	; 읽어온 회사명(ComparingCompanyName)과 ModifiedCompanyName를 비교할 때 공란 때문에 못 찾는 것 같아서
	; ModifiedCompanyName 문자열 앞 뒤의 공란과 특수문자 제거하기
	; 공란만 제거했으면 " " 가 되고 특수문자만 제거하려면 "`t"인데 두 개를 제거하려고 " `t" 사용
	ModifiedCompanyName := Trim(ModifiedCompanyName, " `t")
	
	
	; ModifiedCompanyName 에서 모든 Space(공란)를 제거하기. 이건 회사명 중간의 띄어쓰기도 모두 제거돼서 만들었다가 사용 안한다.
	;StringReplace, ModifiedCompanyName, ModifiedCompanyName, %A_Space%, , All







;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; TeamViewer연결했을때랑 실제랑 화면 포인트 값,위치,행동하는 게 다르다. 수정 시 참고;;;;;;;;;;;;;;;;;;;;;;;;


	Loop{

		; 현재 페이지에 있는 모든 값을 읽기위한 세팅
		; PointOftoY 값을 회사명 끝까지 내렸음
		PointOfX = %A_ScreenWidth%
		PointOfY = %A_ScreenHeight%
		;PointOfX -= 530
		;PointOfY -= 885
		PointOfX -= 530
		PointOfY -= 885		
	
		PointOftoX = %A_ScreenWidth%
		PointOftoY = %A_ScreenHeight%
		;PointOftoX -= 400
		;PointOftoY -= 530
		PointOftoX -= 400
		PointOftoY -= 530		
		
		; 지정된 위치(현재 페이지)에서 Capture2Text 이용해서 값 얻기
		MouseMove, %PointOfX%, %PointOfY%
		SendInput, #q
		Sleep 200
		MouseMove, %PointOftoX%, %PointOftoY%
		Sleep 200
		SendInput, #q
		Sleep 200
					

		; Capture2Text 창 닫기
		Sleep 100
		;IfWinExist, Capture2Text - OCR Text
		;	WinClose
		WinActivate, Capture2Text - OCR Text
		IfWinExist, Capture2Text - OCR Text
			WinClose
						

	;	IfWinExist, Capture2Text - OCR Text
	;		WinClose


		AllCompanyNamesOnCurrentPage := % Clipboard
		Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
		
					

	
		

;		MsgBox, % AllCompanyNamesOnCurrentPage
		
		
		; AllCompanyNamesOnCurrentPage 값에 개행문자(새 줄)이 나올때마다 나눠서 SortedCompanyName 에 저장
		StringSplit, SortedCompanyName, AllCompanyNamesOnCurrentPage, `n, %A_Space%
		;StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.
		
		; SortedCompanyName 배열에는 회사이름 포함 현재 페이지에 있는 모든 값이 들어있음.
		; SortedCompanyName 에 들어있는 값 갯수만큼만 루프 돌려서
		; SortedCompanyName 배열에 들어있는 값들을 루프가 한 번 돌면
		; ComparingCompanyName1 에 값을 넣고
		; 루프가 두 번 돌면
		; ComparingCompanyName1 에 값을 넣고 이렇게 계속 값을 넣는다
;		Loop, %SortedCompanyName0%
;		{
;			ComparingCompanyName%a_index% := SortedCompanyName%a_index%
			;MsgBox, Company name %a_index% is `n%ComparingCompanyName%
			;MsgBox, % ComparingCompanyName%a_index%
;		}
		
		; SortedCompanyName 에 들어있는 값 갯수만큼만 루프 돌려서
		; 루프가 돌 때마다 SortedCompanyName 배열에 들어있는 값을
		; ComparingCompanyName 에 넣는다
		Loop, %SortedCompanyName0%
		{

			ComparingCompanyName := SortedCompanyName%a_index%
			
;			MsgBox, ComparingCompanyName안에 있는 값은 : %ComparingCompanyName%
			
			; 읽어온 회사명(ComparingCompanyName)값 안에 PM값이 있으면
			; 목록에 시간이 나왔다는 뜻이고 회사 목록 끝이라는 뜻이니까
			; 스크롤 다운 할 필요 없고 찾는 회사명 찾으면 됨
			;if(RegExMatch(ComparingCompanyName, "imU)Pickup")){
			;if(RegExMatch(ComparingCompanyName, "imU)PM")){
			
/*;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 루프 돌때마다 ModifiedCompanyName 값을 변화시킬 필요 없어서 아예 처음 루프 시작 전 사용하려고 빼버렸다.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
			; CompanyName 값이 ComparingCompanyName 보다 길면 인식을 못하더라
			; 그래서 CompanyName 값을 잘라서 ModifiedCompanyName 에 넣었다
;			ModifiedCompanyName := SubStr(CompanyName, 1, 16) ;, 19)


			; ModifiedCompanyName 길이는 위의 SubStr 함수로 인해 길이가 16이 되었는데
			; 만약 회사 이름이 16 글자보다 짧으면 그 안에 공란(Space)도 들어가게 되는 것 같음
			; 읽어온 회사명(ComparingCompanyName)과 ModifiedCompanyName를 비교할 때 공란 때문에 못 찾는 것 같아서
			; ModifiedCompanyName 문자열 앞 뒤의 공란과 특수문자 제거하기
			; 공란만 제거했으면 " " 가 되고 특수문자만 제거하려면 "`t"인데 두 개를 제거하려고 " `t" 사용
;			ModifiedCompanyName := Trim(ModifiedCompanyName, " `t")
			
			; ModifiedCompanyName 에서 모든 Space(공란)를 제거하기. 이건 회사명 중간의 띄어쓰기도 모두 제거돼서 만들었다가 사용 안한다.
			;StringReplace, ModifiedCompanyName, ModifiedCompanyName, %A_Space%, , All
*/			
			
;			MsgBox, ComparingCompanyName is %ComparingCompanyName%`n`nModifiedCompanyName : %ModifiedCompanyName%
				
				; 읽어온 회사명(ComparingCompanyName)값 안에 ModifiedCompanyName(찾고 있는 회사이름)값이 있으면 
				; 현재 페이지의 처음 회사명부터 차례대로 하나씩 다시 찾아보기
;				if(RegExMatch(ComparingCompanyName, %ModifiedCompanyName%))
				IfInString, ComparingCompanyName, %ModifiedCompanyName%
				{
					
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; TeamViewer연결했을때랑 실제랑 화면 포인트 값,위치,행동하는 게 다르다. 수정 시 참고;;;;;;;;;;;;;;;;;;;;;;;;

;					MsgBox, It's similar

					PointOfX = %A_ScreenWidth%
					PointOfY = %A_ScreenHeight%
					;PointOfX -= 565
					;PointOfY -= 857
					PointOfX = 1388
					PointOfY = 193
					
					PointOftoX = %A_ScreenWidth%
					PointOftoY = %A_ScreenHeight%
					;PointOftoX -= 431
					;PointOftoY -= 839
					PointOftoX = 1516
					PointOftoY = 209
					
					; SortedCompanyName 에 들어있는 값 갯수만큼만 루프 돌려서
					Loop, %SortedCompanyName0%
					{

						; 지정된 위치에서 값 얻기
						MouseMove, %PointOfX%, %PointOfY%
						SendInput, #q
						Sleep 200
						MouseMove, %PointOftoX%, %PointOftoY%
						Sleep 200
						SendInput, #q
						Sleep 200
						
						
						;Sleep 1000
						;WinClose, Capture2Text - OCR Text
						
						; Capture2Text 창 닫기
						WinWaitActive, Capture2Text - OCR Text
						IfWinExist, Capture2Text - OCR Text
							WinClose
						
						
						
						HavingBeenReadIndividualCompanyName := % Clipboard
						Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
;						MsgBox, HavingBeenReadIndividualCompanyName : %HavingBeenReadIndividualCompanyName%`n`nModifiedCompanyName : %ModifiedCompanyName%
						;MsgBox, ModifiedCompanyName : %ModifiedCompanyName%




						; 읽어온 회사명(HavingBeenReadIndividualCompanyName)값 안에 ModifiedCompanyName(찾고 있는 회사이름)값이 있으면 
						; 마우스 오른쪽 버튼 눌러서 Tracking Number 얻기
;						if(RegExMatch(HavingBeenReadIndividualCompanyName, %ModifiedCompanyName%))
						IfInString, HavingBeenReadIndividualCompanyName, %ModifiedCompanyName%
						{
;							MsgBox, HavingBeenReadIndividualCompanyName`n`n%HavingBeenReadIndividualCompanyName%
							
							Clipboard :=
							
							PointOfX += 10
							PointOfY += 10
													
							MouseClick, r, %PointOfX%, %PointOfY%
							Sleep 200
							Loop, 33{
								SendInput, {Down}
								Sleep 50
							}
									
							SendInput, {Enter}
							Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
							

								
							TrackingNumber := % clipboard
							
							;MsgBox, % clipboard

;							MsgBox, Company Name(읽어온 회사명):%ComparingCompanyName%`n`nInvoice#(송장번호): %TrackingNumber%
							
							; 읽어온 회사명과 같은 것을 찾았으면 루프문을 아예 탈출해야 되는데 뭘로 할까 goto를 써야하나
							goto, InputTrackingNumber
							break					
						}
						
						; 그 다음 회사명을 찾기 위해 Y값들 재설정 하기
						PointOfY += 16
						PointOftoY := PointOfY + 15
						
					}
				}
			;}


		;MsgBox, ComparingCompanyName in loop : %ComparingCompanyName%
			;else{
				;찾는 회사명이 없으면 스크롤 다운 해서 다음 페이지에서 다시 찾아봐야 됨
			;}

		;MsgBox, loop out?
		
		
		
		; 읽어온 회사명(ComparingCompanyName)값 안에 PM값이 없으면
		; 목록에 시간이 없다는 뜻이고 회사 목록 끝이 아니라는 뜻이니까
		; 지금 loop를 빠져나간다.
		; 지금 loop를 빠져나가면 밖의 루프가 하는일 즉 현재 페이지의 값들을 전체 읽는 것을 수행하게 된다.
;		break
		}
		
		; 스크롤 다운 하기
		; 빈 공간 클릭 후 화면의 80%정도 스크롤 다운 하기
		PointOfX = %A_ScreenWidth%
		PointOfY = %A_ScreenHeight%
		PointOfX -= 627
		PointOfY -= 815
		MouseMove, %PointOfX%, %PointOfY%
		Click
		
		Send, {WheelDown 6}
		;Send, {WheelDown 7}

		; 전체 루프 닫기. 현재 페이지에 있는 값 한꺼번에 읽는 루프 닫기
		; ComparingCompanyName 안에 "imU)숫자 숫자 : 숫자 숫자 모든문자 AM 혹은 PM" 있으면, 
		; 즉 이름 다 찾아서 전날의 마감 시간이 나오는 목록표 끝이 나오면 루프 멈춤
		;if(RegExMatch(ComparingCompanyName, "imU)\d*:\d*.*AM|PM"))  ; 이것도 작동됨. "imU)숫자여러개 : 숫자여러개 모든문자 AM 혹은 PM"
		if(RegExMatch(ComparingCompanyName, "imU)\d\d:\d\d.*AM|PM"))
			break
	} ; 전체 루프 닫기. 현재 페이지에 있는 값 한꺼번에 읽는 루프
	;until ComparingCompanyName != PM ; PM이 들어있으면 시간값(업무가 오전에 끝나는 경우는 없으니까) 들어있다는 말이고 목록 끝까지 왔다는 뜻이니까 LOOP 멈추기
	;until ComparingCompanyName == PM ; PM이 들어있으면 시간값(업무가 오전에 끝나는 경우는 없으니까) 들어있다는 말이고 목록 끝까지 왔다는 뜻이니까 LOOP 멈추기



;MsgBox, % ComparingCompanyName














/*

	; 현재 페이지에 있는 모든 값을 읽기위한 세팅
	; PointOftoY 값을 회사명 끝까지 내렸음
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 605
	PointOfY -= 808
	
	PointOftoX = %A_ScreenWidth%
	PointOftoY = %A_ScreenHeight%
	PointOftoX -= 400
	PointOftoY -= 530
	
	; 지정된 위치(현재 페이지)에서 값 얻기
	MouseMove, %PointOfX%, %PointOfY%
	SendInput, #q
	Sleep 200
	MouseMove, %PointOftoX%, %PointOftoY%
	Sleep 200
	SendInput, #q
	Sleep 200
				
	; Capture2Text 창 닫기
	WinWaitActive, Capture2Text - OCR Text
	WinClose
	
;	IfWinExist, Capture2Text - OCR Text
;		WinClose


	AllCompanyNamesOnCurrentPage := % Clipboard
	Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨

	MsgBox, % AllCompanyNamesOnCurrentPage
	
	
	; AllCompanyNamesOnCurrentPage 값에 개행문자(새 줄)이 나올때마다 나눠서 SortedCompanyName 에 저장
	StringSplit, SortedCompanyName, AllCompanyNamesOnCurrentPage, `n, %A_Space%
	;StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.

	; SortedCompanyName 에 들어있는 값 갯수만큼만 루프 돌려서
	; 
	Loop, %SortedCompanyName0%
	{
		ComparingCompanyName := SortedCompanyName%a_index%
		MsgBox, Company name %a_index% is `n%ComparingCompanyName%


		; 읽어온 회사명(ComparingCompanyName)값 안에 PM값이 있으면
		; 목록에 시간이 나왔다는 뜻이고 회사 목록 끝이라는 뜻이니까
		; 스크롤 다운 할 필요 없고 찾는 회사명 찾으면 됨
		if(RegExMatch(ComparingCompanyName, "imU)Pickup")){
		;if(RegExMatch(ComparingCompanyName, "imU)PM"))
			
			; 읽어온 회사명(ComparingCompanyName)값 안에 CompanyName(찾고 있는 회사이름)값이 있으면 
			; 현재 페이지의 처음 회사명부터 차례대로 하나씩 다시 찾아보기
			IfInString, ComparingCompanyName, %CompanyName%
			{

	;			MsgBox, It's similar

				PointOfX = %A_ScreenWidth%
				PointOfY = %A_ScreenHeight%
				PointOfX -= 605
				PointOfY -= 808
				
				PointOftoX = %A_ScreenWidth%
				PointOftoY = %A_ScreenHeight%
				PointOftoX -= 400
				PointOftoY -= 792
				
				
				; SortedCompanyName 에 들어있는 값 갯수만큼만 루프 돌려서
				Loop, %SortedCompanyName0%
				{

					; 지정된 위치에서 값 얻기
					MouseMove, %PointOfX%, %PointOfY%
					SendInput, #q
					Sleep 200
					MouseMove, %PointOftoX%, %PointOftoY%
					Sleep 200
					SendInput, #q
					Sleep 200
					
					
					Sleep 1000
					WinClose, Capture2Text - OCR Text
					
					
					
					HavingBeenReadIndividualCompanyName := % Clipboard
					Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨



					; 읽어온 회사명(HavingBeenReadIndividualCompanyName)값 안에 CompanyName(찾고 있는 회사이름)값이 있으면 
					; 마우스 오른쪽 버튼 눌러서 Tracking Number 얻기
					IfInString, HavingBeenReadIndividualCompanyName, %CompanyName%
					{
						Clipboard :=
						PointOfX += 10
						PointOfY += 10
												
						MouseClick, r, %PointOfX%, %PointOfY%
						Sleep 200
						Loop, 33{
							SendInput, {Down}
							Sleep 100
						}
								
						SendInput, {Enter}
						Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
						

							
						TrackingNumber := % clipboard
						
						MsgBox, % clipboard

						MsgBox, 읽어온 회사명:%ComparingCompanyName%`n`n송장번호: %TrackingNumber%
						
						; 읽어온 회사명과 같은 것을 찾았으면 루프문을 아예 탈출해야 되는데 뭘로 할까 goto를 써야하나
						break					
					}
					
					; 그 다음 회사명을 찾기 위해 Y값들 재설정 하기
					PointOfY += 16
					PointOftoY := PointOfY + 15
					
				}
			}
		}










		else{
			;찾는 회사명이 없으면 스크롤 다운 해서 다음 페이지에서 다시 찾아봐야 됨
		}

	MsgBox, loop out?
	}		


*/































/*
	
	
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 605
	PointOfY -= 808
	
	PointOftoX = %A_ScreenWidth%
	PointOftoY = %A_ScreenHeight%
	PointOftoX -= 400
	PointOftoY -= 792

Loop, 8{

	; 지정된 위치에서 값 얻기
	MouseMove, %PointOfX%, %PointOfY%
	SendInput, #q
	Sleep 200
	MouseMove, %PointOftoX%, %PointOftoY%
	Sleep 200
	SendInput, #q
	Sleep 200
	
	; 그 다음 회사명을 찾기 위해 Y값들 재설정 하기
	PointOfY += 16
	PointOftoY := PointOfY + 15
	
	Sleep 1000
	WinClose, Capture2Text - OCR Text
	
}

*/




	
	
	

/*
	; 첫 번째 회사명 읽고 Company Name과 맞는지 비교하기
	; 첫 번째 회사명 읽기 위한 위치 지정
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 605
	PointOfY -= 808
	
	PointOftoX = %A_ScreenWidth%
	PointOftoY = %A_ScreenHeight%
	PointOftoX -= 400
	PointOftoY -= 792


	; 지정된 위치(첫 번째 회사명)에서 값 얻기
	MouseMove, %PointOfX%, %PointOfY%
	SendInput, #q
	Sleep 200
	MouseMove, %PointOftoX%, %PointOftoY%
	Sleep 200
	SendInput, #q
	Sleep 200
				
	; Capture2Text 창 닫기
	IfWinExist, Capture2Text - OCR Text
		WinClose


	ReadVal := % Clipboard
				
				
	; 읽어온 회사명(ReadVal)값이 CompanyName과 같으면 마우스 오른쪽 버튼 눌러서 Tracking Number 얻기
	if(ReadVal == CompanyName){
		Clipboard :=
		PointOfX += 10
		PointOfY += 10
								
		MouseClick, r, %PointOfX%, %PointOfY%
		Sleep 200
		Loop, 33{
			SendInput, {Down}
;			Sleep 100
		}
				
		SendInput, {Enter}
		Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
		

			
		TrackingNumber := % clipboard
		
		MsgBox, % clipboard

		MsgBox, 읽어온 회사명:%ReadVal%`n`n송장번호: %TrackingNumber%

	}
*/


/*;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	; Invoice_Wts값에 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 Wts_of_Boxes에 저장
	StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.

	; Wts_of_Boxes를 루프 돌려서 들어있는 값 개수만큼 No_of_Boxes에 저장
	Loop, %Wts_of_Boxes0%
	{
		
		this_Wts := Wts_of_Boxes%a_index%
	;	MsgBox, The Weight Box number %a_index% is %this_Wts%
		No_of_Boxes := % A_Index
	}	
*/	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;






	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;;;;회사 이름 찾았으면 이리로 goto 해서 옴;;;;
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	
	InputTrackingNumber:



	if(TrackingNumber){
		TrackingNumber := % Clipboard
		/*
		Send, % TrackingNumber
		Sleep 500
		;ControlSend, , {F3}, UPS WorldShip - Remote Workstation
		;ControlSend,
		;ControlSend, Button1, {Enter}, Tracking Number
		;ControlClick, Button1, Tracking Number
		ControlClick, Enter_, Tracking Number ;, , l, , NA
		SendInput, {Enter}
		*/
		
		;UPS 배송 창으로 넘어감으로 초기화 하면서 함수 종료
	;	Control_SnedButton("F2", windowtitle)
		Loop, 3
			Control_SnedButton(F2, windowtitle)
			
		; UPS 창 항상위에 설정 해제
		Winset, AlwaysOnTop, Off, UPS WorldShip - Remote Workstation
		return
	}
		





	; Tracking Number 넣기 위한 GUI
	Gui, Show, w330 h88, Tracking Number

	Gui, Add, Text, x12 y25 w60 h30 +Center, Tracking Number
	Gui, Add, Edit, x72 y28 w180 h20 vTrackingNumber
;	Gui, Add. Button, x92 y28 w30 h30, ASDF

	Gui, Add, Button, x260 y28 w60 h20 +default, Enter_


	;GUI시작 시 포커스를 Invoice_No 입력칸에 위치
	GuiControl, Focus, TrackingNumber
	
	; 창이 종료될때까지 대기
	WinWaitClose, Tracking Number
	
	;Enter 버튼 누르면 Tracking Number GUI창 종료
	ButtonEnter__:
	WinClose, Tracking Number
	
	; GUI에서 입력받은 값을 사용하기 위해
	Gui, Submit, nohide
;	GUI, TrackingNumber:Destroy
	GUI, Hide




;MsgBox, pause

	;UPS 배송 창으로 넘어감으로 초기화 하면서 함수 종료
;	Control_SnedButton("F2", windowtitle)
	Loop, 3
		Control_SnedButton(F2, windowtitle)

		
	; UPS 창 항상위에 설정 해제
	Winset, AlwaysOnTop, Off, UPS WorldShip - Remote Workstation
	
	
	
	
	return
}





