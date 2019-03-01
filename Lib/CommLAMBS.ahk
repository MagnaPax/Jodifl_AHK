#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;~ class CommonLAMBSProcessing{
class CommLAMBS{

	BasicLAMBSProcessing(){
		
		; 만약 LAMBS 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		loop{
			
			IfWinExist, LAMBS
				break			
			MsgBox, 262144, No LAMBS Warning, PLEASE RUN LAMBS
		}
		
		
		; 혹시 창을 열어놓은 상태로 중단했을 수도 있으니까 창 닫기
		IfWinExist, Transfer from Sales Order
			WinClose
		IfWinExist, Customer Order +Zoom In
			WinClose
		IfWinExist, Accounts Summary
			WinClose
		
		
		CommLAS.Start()


	}


	;LAMBS활성화 후 화면 초기화 하기
	Start(){		
		
		; 만약 LAMBS 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		loop{
			
			IfWinExist, LAMBS
				break			
			MsgBox, 262144, No LAMBS Warning, PLEASE RUN LAMBS
		}
		
		;LAMBS Window 활성화 하기
		WinActivate, LAMBS -  Garment Manufacturer & Wholesale Software
		windowtitle = LAMBS -  Garment Manufacturer & Wholesale Software
		
		Send, {WheelUp}
		
		CommLAS.CheckTheWindowPresentAndActiveIt(windowtitle)

		;Hide All 클릭해서 메뉴 바 없애기
		CommLAS.ClickAtThePoint(213, 65)
		
		return
	}
	

	;CC 버튼 찾고 클릭
	FindCCButtonAndClickIt(){

		Text:="|<CConLAMBS>*140$58.D002E0S005000882000M000UUE001UCQSLV0SQy0m+98409YM2DcYUE3YFU8UWG10GF50X2982194Hm7bYs7bYDU"

		if ok:=FindText(3016,149,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		}
	}
	
		
	;지정된 창이 존재하는지 무한정 확인 후 그 창 활성화
	CheckTheWindowPresentAndActiveIt(windowtitle){
		WinWait, % windowtitle
		WinActivate, % windowtitle
		return
	}
	
	
	;위치 받아서 클릭하기
	ClickAtThePoint(XPoint, YPoint){
		MouseClick, l, XPoint, YPoint, 1
		Sleep 1000
		return
	}
	

	;상태바에서 알트키, 방향키 등 눌러서 Create Sales Orders Small Tab열기
	ActiveCreateSalesOrdersSmallTab(){
		
		CommLAS.Start()
		
		WinActivate, LAMBS -  Garment Manufacturer & Wholesale Software
		Send, {WheelUp}
		

		; 혹시 위의 동작으로도 Create Sales Orders Small 탭으로 넘어가지 않았을 때 백업 용도로 비활성화된 Create Sales Orders Small 버튼 클릭하기
		; 이 코드는 메뉴를 클릭해서 Create Sales Orders Small 을 새로 여는 게 아니라 이미 열려있는데 활성화 되지 않았을 때만 유효함
		; 혹시 마우스 커서 등이 가려서 못 찾는 것을 방지하기 위해 일단 마우스를 옮기기
;		MouseMove, A_ScreenWidth/2, A_ScreenHeight/2


		Text:="|<>*176$26.Q002M000a0009UhnmKAY4YG8792W22EcUYbm8D9U"
		
		if ok:=FindText(296,90,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
				
			; 탭 버튼 클릭 후 화면의 정중앙으로 마우스 포인터 이동시키기
;			MouseMove, A_ScreenWidth/2, A_ScreenHeight/2
			;~ MsgBox
				
			; 버튼 찾아서 클릭했으면 그냥 함수를 빠져나가기
			;~ MsgBox, BEING FOUND BY FindText()
			return
		}
*/		

		Send, {Alt}
		Sleep 10
	
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10

		Send, {Down}
		Sleep 10
		
		Send, {Right}
		Sleep 10

		Send, {Down}
		Sleep 10
		
		Send, {Enter}
		Sleep 100		

		return
	}
	
	
	

}