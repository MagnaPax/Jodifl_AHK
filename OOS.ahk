/*
테스트용 SO 번호들
35451
33021 스타일번호 같은게 두 개 있음. 첫 번째 화면에 주문 갯수 8개
33048 마지막 화면에 8개의 Desc 있음
28600 클로즈 해도 되는 주문 클로즈 테스트에 이용하면 될듯함 (이미 클로즈 했음)
*/




#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ ; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; 에러 메세지 경고창 안 뜨게 하는 함수
;~ ComObjError(false)


#Include %A_ScriptDir%\lib\




#Include CommN41.ahk
#Include OOS_Lib.ahk
#Include Xl_Lib.ahk
;~ #Include N41_Lib.ahk
#Include <N41_Lib>

;~ #Include FindTextFunctionONLY.ahk
;~ #Include FG.ahk
;~ #Include N41.ahk
;~ #Include LAS.ahk
;~ #Include ChromeGet.ahk
;~ #Include MethodsForPT.ahk ; PT 프로세싱을 위한 모든 클래스, 메소드를 모아놓은 파일


;~ global preItems_Arr := object() ; 이전 아이템을 저장하기 위한 배열
;~ global currentItems_Arr := object() ; 현재 아이템을 저장하기 위한 배열


global allValuesOfItems ; 화면에 있는 모든 값을 저장하는 변수. 나중에 중복되는 값이 있나 비교하기 위함	
global HowMany8DescScrrenHasBeenAppeared	; Desc 갯수가 8개 나온 화면이 몇 개 인지 누적하는 변수


; 메소드에서 오는 여러 값들을 받기 위한 배열 선언
Array := object()
Info_Arr := object()
SO#_Arr := object()
CustCode_Arr := object()
CustPO#_Arr := object()
OrderQty_Arr := object()
itemsInfoThatJustReadFromScreen_Arr := object()

	
; 배열 초기화 해주기
Array := []
Info_Arr := []
SO#_Arr := []
CustCode_Arr := []
CustPO#_Arr := []
OrderQty_Arr := []
itemsInfoThatJustReadFromScreen_Arr := []
;~ preItems_Arr := []
;~ currentItems_Arr := []







;GUI Backgroud
Gui, Show, w245 h160, OUT OF STOCK, AlwaysOnTop Window
WinSet, AlwaysOnTop, On, OUT OF STOCK ; GUI창 항상 위에 위치시키기

;Input Start Order Id
Gui, Add, Text, x22 y21 Cred , OOS Style #
Gui, Add, Edit, x92 y20 w100 h20 vOOSStyle#, S1156 ;C1116-8 ;P2276-1 ;b1323 ;P2471 ;B3595-1(drop all) ;P4030 ;p2420-12(DROP ALL)             				            ;P4061 ;B3610 ;S1229 ;P2933

;Input End Order Id
Gui, Add, Text, x22 y51 CBlue , Color
Gui, Add, Edit, x92 y50 w100 h20 vOOSColor, MUSTARD ;navy mix ; olive mix

;Check Box For Drop All
Gui, Add, CheckBox, x22 y70 w80 h40 vDropAll, Drop All



;엔터 버튼
Gui, Add, Button, x22 y110 w200 h30 +default gClick_btn, Enter





;GUI시작 시 포커스를 OOS Style # 칸에 위치
GuiControl, Focus, OOSStyle#


return



Click_btn:


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; To use the values which input on GUI
	Gui Submit, nohide
	GUI, Destroy
	
	
	
	; 만약 값을 입력 안 했으면 프로그램 다시 시작하기
	if(OOSStyle# == ""){
		Reload		
	}
	else if(OOSColor == "")
		Reload
	

	; 혹시나 있을지도 모를 아이템 번호값과 색깔 값의 양쪽 공백 없애기
	OOSStyle# := Trim(OOSStyle#)
	OOSColor := Trim(OOSColor)
	StringUpper, OOSStyle#, OOSStyle# ; 대문자로 바꾸기
	StringUpper, OOSColor, OOSColor ; 대문자로 바꾸기	

	
	; 만약 N41 이 존재하지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
	IfWinNotExist, ahk_class FNWND3126
	{
		loop{
			MsgBox, 262144, No N41 Warning, PLEASE START N41
			IfWinExist, ahk_class FNWND3126
				break
		}
	}


	; N41 활성화하기
	WinActivate, ahk_class FNWND3126
	
	; 객체 선언
	CommN41_driver := new CommN41




/*
	Loop
	{
		; Style 클릭하기
		CommN41_driver.click_Style()
		
		; Style 탭 안에 있는 Style 클릭하기
		CommN41_driver.click_StyleButtonInTheStyleTab()
		
		; 스타일 입력칸 찾아서 클릭하기
		CommN41_driver.findTheSearchBoxOfStyle()
		
		; 입력칸에 이미 있던 값 지우기
		Loop, 30{
			Send, {Del}
			Sleep 50
		}
		; 스타일 번호 입력
		Send, %OOSStyle#%
		Send, {Enter}
	;	Sleep 1000
		
		
		; 색깔 입력칸 찾아서 클릭하기
		CommN41_driver.findTheSearchBoxOfColor()
		
		; 색깔 입력
		Send, %OOSColor%	
		Send, {Enter}
	;	Sleep 1000
		
		
		; 화면에 제대로 들어왔는 확인키 위해 화면에서 스타일 번호 읽기
		Style#_OnTheScreen := CommN41_driver.getStyleNumberOnStyleTab()
		
		; 화면에 제대로 들어왔는 확인키 위해 화면에서 색깔 읽기
		Color_OnTheScreen := CommN41_driver.getColorOnStyleTab()
		
		
;MsgBox, % Style#_OnTheScreen . "`n`n" . Color_OnTheScreen
		
		; 화면에서 읽은 값들과 입력한 값들이 같으면 루프 빠져나가기
		if(OOSStyle# == Style#_OnTheScreen && OOSColor == Color_OnTheScreen)
			break
	}
	
	

	; On Order 의 돋보기 버튼 눌러서 창 열기
	CommN41_driver.clickMagnifierNextTo_OnOrder()
	
	
	; 적당한 창 위치에서 (이 경우에는 Close 버튼을 찾은 뒤 그 위에 마우스를 위치했음) 마우스 오른쪽 버튼 클릭 후 Excel 파일 만들기 메뉴에서 엔터치기
	CommN41_driver.activeDownloadExcelFileFunction()
	
	
	; 엑셀파일을 바탕화면에 저장하기
	CommN41_driver.makeTheExcelFileOfTheStyleOnDesktopScreen(OOSStyle#, OOSColor)
	
	
	; 바탕화면에 저장된 엑셀 파일 옮기기
	moveTheFileOnDesktopFromItToTheExcelFileFolder(OOSStyle#, OOSColor)
	
*/	
	
	; Path 에 위치한 엑셀창 열기
	Path = %A_ScriptDir%\CreatedFiles\Excel Files\%OOSStyle#%-%OOSColor%.xlsx
	Xl := openTheExcelFile(Path, OOSStyle#, OOSColor)
	
	
/*
	; 필요 없는 열들 지우기 위해 변수값 지정
	; 맨 처음 유효값이 시작되기 1전 값과 맨 끝에 지울 줄 갯수 변수에 넣기
	1before_StartRow# = 4
	#oflasUselessRows = 2
	
	
	; 필요 없는 열들 지우기
	deleteUselessRows(Xl, 1before_StartRow#, #oflasUselessRows)
*/	
	
	
	SO#_rowLetter = B
	customerCode_rowLetter = F
	custPO#_rowLetter = U
	orderQTY_rowLetter = AC
	Info_Arr := getInfoFronExcel(Xl, SO#_rowLetter, customerCode_rowLetter, custPO#_rowLetter, orderQTY_rowLetter)
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	; Info_Arr[1][*] <- SO#				;
	; Info_Arr[2][*] <- Customer Code	;
	; Info_Arr[3][*] <- Custoer PO#		;
	; Info_Arr[4][*] <- Order QTY		;
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	


; 2차원 배열 Info_Arr 에서 값 읽기
For 1st_index, Row in Info_Arr{
	For 2nd_index, element in Row{
		
		; SO# 처리
		if(1st_index == 1){
;			MsgBox, 262144, Title, %2nd_index% 번째 SO 번호 : %element%
			SO#_Arr.Insert(element)
		}
		; Customer Code 처리
		else if(1st_index == 2){
;			MsgBox, 262144, Title, %2nd_index% 번째 Customer Code 번호 : %element%
			CustCode_Arr.Insert(element)
		}
		; Custoer PO# 처리
		else if(1st_index == 3){
;			MsgBox, 262144, Title, %2nd_index% 번째 Custoer PO 번호 : %element%
			CustPO#_Arr.Insert(element)
		}
		; Order QTY
		else if(1st_index == 4){
;			MsgBox, 262144, Title, %2nd_index% 번째  Order QTY : %element%
			OrderQty_Arr.Insert(element)
		}		
		
		;~ MsgBox, % "1st : " . 1st_index . "`n`n2nd : " . 2nd_index . "`n`n" . "Array" . "[" . 1st_index . "][" . 2nd_index . "]" . "`n`nelement : " . element
	}
}


; 배열에 들어있는 값의 갯수만큼 반복(엑셀 파일에 있는 주문 갯수만큼 반복)
; 어차피 어떤 배열이든 들어있는 값의 갯수는 같기 때문에 어떤 배열을 써도 무방함. 꼭 SO#_Arr 배열이 아니어도 됨
for index, element in SO#_Arr
{	
	; Desc 갯수가 8개 나온 화면이 몇 개 인지 누적하는 변수값 초기화
	HowMany8DescScrrenHasBeenAppeared = 0
	
	; Sales Order 창 열기
;	openSO(element)


	; 현재 화면에서 Desc 를 모두 찾아서 각각의 X,Y 위치값을 얻은 뒤 그 X,Y 값을 기준으로 스타일번호, 색깔, 수량값을 읽고 드롭할 값과 비교한 뒤 맞으면 드롭하기
	Loop{
		
		; Sales Order 화면에 있는 Desc. 전체의 위치 X와 Y값 가져오기
		XY_Arr := getXandYcoordinatesOf_Desc()
		;;;;;;;;;;;;;;;;;;;;;;;;;
		; XY_Arr[1][*] <- X 좌표	;
		; XY_Arr[2][*] <- Y 좌표	;
		;;;;;;;;;;;;;;;;;;;;;;;;;


		; XY_Arr 배열에 저장되어 있는 X 값과 Y 값을 얻기
		For 1st_index, Row in XY_Arr{
			For 2nd_index, element in Row{
				
				; 어차피 X 좌표는 다 똑같기 때문에 이렇게 해줘도 됨
				X := XY_Arr[1][1]
				
				; X 좌표는 똑같고 Y 좌표만 차이 나기때문에 X값이 들어있는 1st_index 가 1일때는 필요 없다 Y값이 들어있는 2일때만
				if(1st_index == 2){
					
; ##########################################################################################################################################################################################################################################
; 테스트를 위해 엑셀에서 읽은 값을 아래의 값으로 바꿨음
	OOSStyle# = B1772
	OOSStyle# = J2801
	OOSStyle# = P3015
	OOSStyle# = C1002
	OOSColor = GOLDEN YELLOW MIX
	OOSColor = BLACK
	OrderQty = 6
	DropAll = 0
; ##########################################################################################################################################################################################################################################
	
					; Desc 위치를 기준 삼아서 스타일번호, 색깔, 수량을 읽은 뒤 드롭해야 되는 스타일이면 드롭하기
					itemsInfoThatJustReadFromScreen_Arr := checkValuesAndDropItems(X, XY_Arr[2][2nd_index], OOSStyle#, OOSColor, OrderQty, DropAll)

					; 지금 읽은 아이템 정보가 allValuesOfItems 변수에 있는지 비교해서 두 번째 읽는 화면이면 아이템 읽는 동작 중단하고 나가기
					isThis2ndReadScreen(XY_Arr[2][2nd_index], itemsInfoThatJustReadFromScreen_Arr)
					
					; 현재 화면에서 읽어온 아이템 정보들 allValuesOfItems 변수에 저장. 나중에 이 변수를 이용해 마지막 화면인지 확인하게 됨
					saveInfoOfTheItemsJustReadToallValuesOfItemsVariable(XY_Arr[2][2nd_index], itemsInfoThatJustReadFromScreen_Arr)
					

					
					; Drop All 이 아니고 화면에서 드롭할 아이템을 찾아 드롭했다면 모든 Desc를 찾을 필요 없이 For 문 빠져나가기
					if(!DropAll && doesTheItemFound){
						break 3
					}
				} ; end끝 - if(1st_index == 2)				
				
				;~ MsgBox, % "1st : " . 1st_index . "`n`n2nd : " . 2nd_index . "`n`n" . "Array" . "[" . 1st_index . "][" . 2nd_index . "]" . "`n`nelement : " . element . "`n`nX : " . X . "`nY : " . XY_Arr[2][2nd_index]
				
			} ; end끝 - For 2nd_index, element in Row
		} ; end끝 - For 1st_index, Row in XY_Arr


		; DropAll 이거나(모든 화면에 있는 아이템을 찾아야 되기 때문) 위의 처리를 통해 현재 화면에서 드롭할 아이템을 찾지 못했다면 화살표를 찾아서 8번 클릭하여 다음 화면으로 넘어감
		if(DropAll || !doesTheItemFound){
			
			; 화면 끝까지 찾았으면 화살표 클릭하지 말고 나가기			
			; 화면에 Desc 갯수가 1~7개 밖에 없다면 첫째 화면이든 넷째 화면이든 어쨌거나 화면 끝까지 도달한 것이므로 화살표 클릭 않기
			if(XY_Arr.1.MaxIndex() <= 7){
				MsgBox, 262144, title, 끝에 도달했음. 루프 빠져나감
				break
			}
			
			; 화면에 Desc 갯수가 8개라면 HowMany8DescScrrenHasBeenAppeared 를 1 증가시켜서 지금까지 몇 번 나왔는지 확인
			;~ if(XY_Arr.1.MaxIndex() == 8){
				;~ HowMany8DescScrrenHasBeenAppeared++
				
				;~ if(HowMany8DescScrrenHasBeenAppeared == 1){
					;~ allValuesOfItems := ""	;이전 화면에서 읽어서 저장한 값들 초기화
				;~ }
			;~ }
			
			
			
			
; ##########################################################################################################################################################################################################################################
; ##########################################################################################################################################################################################################################################
; ##########################################################################################################################################################################################################################################
; 화면에 Desc 갯수가 8개 인데 마지막 화면일때 어떻게 처리할 지 고민하기
/*
global preItems_Arr := object() ; 이전 아이템을 저장하기 위한 배열
global currentItems_Arr := object() ; 현재 아이템을 저장하기 위한 배열

1 값을 읽어서 currentItems_Arr 배열에 저장
2 화살표 8번 클릭한 뒤 preItems_Arr := currentItems_Arr.Clone() 통해 복사함
3 화살표 8번 클릭 전 두 배열이 같은지 비교. 같다면 루프 빠져나오기
그런데 이러면 마지막 화면을 두 번 읽어야 되는데 만약 마지막 화면에 드롭할 아이템이 있으면 오류 나지 않으려나?
아니면 checkValuesAndDropItems 함수에서 OOSStyle#, OOSColor, OrderQty 세 개 값이 다 배열에 있으면 중지하면 될까?
*/
; ##########################################################################################################################################################################################################################################
; ##########################################################################################################################################################################################################################################
; ##########################################################################################################################################################################################################################################
			
			
			MsgBox, % "allValuesOfItems값은`n`n" . allValuesOfItems
			Sleep 500
			
			; 아래쪽 화살표 8번 클릭하기
			doesTheArrowButtonExist := clickDownArrow8Times()
			if(!doesTheArrowButtonExist){
				MsgBox, 262144, title, 화살표를 못 찾았음. 첫 화면에 Desc 갯수가 8개인 주문일 가능성 높음. 이 주문을 끝내고 다음 주문으로 넘어감
				break
			}
			
			
			; 화면을 바꾼 뒤에는 배열 초기화 해주기
			Duplicate_Arr := []
		}
		
		
		
	} ; end끝 - Loop{

	MsgBox, 262144, title, 저장하기. 저장할 때 나올 수 있는 경고창 어떻게 처리할 지 생각하기
	
	MsgBox, 262144, title, 화살표가 있으면 8번 클릭하는것 넣기
	
	MsgBox, % XY_Arr[1][1]
	MsgBox, % XY_Arr[1][1]
	MsgBox, % XY_Arr[1][1]
	MsgBox, % XY_Arr[1][1]
	
	;~ ; 스타일 번호, 색깔, 수량이 맞는 곳의 Desc. X와 Y 좌표 얻기
	;~ ; DropAll 체크됐으면 번호만 찾아서 모두 다 내리기
	;~ OrderQty := OrderQty_Arr[index]
	;~ findTheStyleOnSalesOrderTab(OOSStyle#, OOSColor, OrderQty, DropAll)
	
	
	MsgBox % "Element number " . index . " is " . element
	MsgBox, % index . " 번째 Customer Code 번호 : "  . CustCode_Arr[index]
}





for index, element in CustCode_Arr
{
	MsgBox % "Element number " . index . " is " . element
}

for index, element in CustPO#_Arr
{
	MsgBox % "Element number " . index . " is " . element
}

for index, element in OrderQty_Arr
{
	MsgBox % "Element number " . index . " is " . element
}





MsgBox, pause1	





Array:=[1,3,"ㅋㅋ"]
for index, element in Array
{
	MsgBox % "Element number " . index . " is " . element
}



	
	;~ Loop % Str_Total_ExcelInfo.Maxindex(){
	i = 1
	j = 1
	Loop % Info_Arr.Maxindex(){
		Loop % Info_Arr.Maxindex(){
			;~ MsgBox, % "Info_Arr[" . A_Index . "][" . A_Index . "]`n" . Info_Arr[A_Index][A_Index]
			MsgBox, % "Info_Arr[" . i . "][" . j . "]`n" . Info_Arr[i][j]
			j++
		}
		i++
	}
	
	
MsgBox, pause
	
	
	MsgBox, % Info_Arr[1][1]
	MsgBox, % Info_Arr[2][1]
	MsgBox, % Info_Arr[3][1]
	MsgBox, % Info_Arr[4][1]
	
MsgBox	
	
	
	MsgBox, % Info_Arr[1][1]
	MsgBox, % Info_Arr[1][2]
	MsgBox, % Info_Arr[1][3]
	MsgBox, % Info_Arr[1][4]
	MsgBox, % Info_Arr[1][5]
	
	
	
	
	;~ Xl.Range("A:A").Copy ;copy cell to clipboard	
	;~ MsgBox, % Clipboard
	
	
	;~ while(Xl.Range("A" . A_Index).value != "")
	;~ {
		;~ Xl.Range("A" . A_Index).Value := value
		;~ MsgBox, % Xl.Range("A" . A_Index).Value
	;~ }
	
	MsgBox, pause
	
	
	Array := Clipboard
	
	
	;~ Array := Xl.Range("A:A").Copy ;copy cell to clipboard
	

;~ /* 배열로부터 읽기 첫 번째 방법
Loop % Array.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
}
*/
	
	
	;~ return [so#, customerCode, custPO#, total#OfRows]


	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	Reload	
	

	

	Esc::
	Exitapp	