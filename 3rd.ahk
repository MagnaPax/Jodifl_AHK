/*
1. 1st.ahk 파일 실행해서 각 아이템파일 pending_styles_not_shipped 폴더에 저장
;~ 2. '저녁이있는엑셀 - 내용합치기 - 폴더선택 후 - 어느범위-시작셀 : A2' 이용해서 저장된 파일들을 통합
3. 통합된 파일을 연 뒤 이것 이용해서 중복된 고객 코드, 티파니 고객 코드 삭제하기
*/

/*
; 티파니 고객 명단

AHN AND AHN COLLECTI
ALLURE
AMBIANCE-PLACERVILLE
AUBURN UNIVERSITY
BETSEYS BOUTIQUE SHO
BETTY BE GOOD BOUTIQ
BOHO BLU
BOOP DE DOOP
CHARLI ROSE-CO
COBOS BOUTIQUE
COBOS-STARKVILLE
ENTOURAGE
FINNLEYS
FINNLEYS 2
FLAUNT BOUTIQUE-ROUN
FOCUSED HOLDINGS
JOSIES BOUTIQUE
JUST JEWELRY
K-S WHOLESALE
KIKILARUE
MAGNOLIA BTQ FRANKLI
ROOLEE
SHOP DRESS UP
STYLEFOX
THE COLLECTION DBA D
THE GIRLIE BOUTIQUE
THE PINK LILY BOUTIQ
THREADS BOUTIQUE
TY ALEXANDERS
UNIQUE BOUTIQUE-SUTT
UNIVERSITY BOOK STOR
WAKEFIELDS-INC
WREN AND IVORY LLC

*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\


#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk




	MsgBox, 262144, 시작, OK to START


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
	BlockInput, Mouse


	; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
	IfWinNotExist, ahk_class XLMAIN
	{
		loop{
			MsgBox, 262144, No Excel file Warning, PLEASE OPEN AN ORDER EXCEL FILE
			IfWinExist, ahk_class XLMAIN
				break
		}
	}	
	

	
	
	
	; 아이템 찾는 동안 보여줄 프로그래스 바
	TotalLoops = 10000
	Gui, -Caption +AlwaysOnTop +LastFound
	Gui, Add, Text, x12 y9 w170 h20 , P  R  O  C  E  S  S  I  N  G  .  .  .
	Gui, Add, Progress, w410 Range0-%TotalLoops% vProgress
	Gui, Show, w437 h84, REFINE THE EXCEL




Xl := ComObjActive("Excel.Application")	; 열려있는 엑셀 창 사용하기





; ################################ 중복된 고객명 삭제하기 ################################


	; 졍렬에 방해되는 열(Columns) 지우기
	XL_Col_Delete(XL,RG:="D|E|H:J|M|O|P|R|S|U|W|Z|AA") ;Delete columns
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	
	
MsgBox, 262144, Title, 정렬에 방해되는 통합 열 지웠음

	; 열의 넓이 설정하기
	XL_Col_Width_Set(XL,RG:="E=25")
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	


	; 중복되는 고객 코드를 지우기 위해 일단
	; 고객 코드가 있는 5 열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(5), 1)
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	


	;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
	XL_Handle(XL,1) ;get handle to Excel Application
	i := XL_Last_Row(XL)
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
	

		
	j = 1
	Loop{

		k := j + 1
		
		#ofRowToRead := "E" . j
		#ofRowToReadAddedOne :=  "E" . k
			
			
		var1 := Xl.Range(#ofRowToRead).Value
		var2 := Xl.Range(#ofRowToReadAddedOne).Value
			
;MsgBox, % var1 . "`n" . var2
			
			
		if(var1 == ""){
			break
		}
						

		; 만약 지금 얻은 고객코드 값이 이전 고객코드 값을 저장하고 있는 previousNumber 값과 같다면 
		; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
		IfEqual, var1, %var2%
		{
				
;MsgBox, % var1 . "`n" . var2 . "`n`n" . "delete" . "`n`n" . "i : " . i . "`nj : " . j
			Xl.Sheets(1).Range(#ofRowToReadAddedOne).EntireRow.Delete
				
			GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가				
			continue
		}
			
		IfNotEqual, var1, %var2%
		{
			j++
				
			GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
			continue
		}
			
			
		GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가

		}

;MsgBox 중복된 고객코드 지워졌나?








; ################################ 티파니 고객 찾아서 삭제하기 ################################

;~ Xl := ComObjActive("Excel.Application")	; 열려있는 엑셀 창 사용하기

; Constants
xlValues := -4163
xlWhole := 1


TiffanyOnlyCustomerList := ["AHN AND AHN COLLECTI", "ALLURE", "AMBIANCE-PLACERVILLE", "AUBURN UNIVERSITY", "BETSEYS BOUTIQUE SHO", "BETTY BE GOOD BOUTIQ", "BOHO BLU", "BOOP DE DOOP", "CHARLI ROSE-CO", "COBOS BOUTIQUE", "COBOS-STARKVILLE", "ENTOURAGE", "FINNLEYS", "FINNLEYS 2", "FLAUNT BOUTIQUE-ROUN", "FOCUSED HOLDINGS", "JOSIES BOUTIQUE", "JUST JEWELRY", "K-S WHOLESALE", "KIKILARUE", "MAGNOLIA BTQ FRANKLI", "ROOLEE", "SHOP DRESS UP", "STYLEFOX", "THE COLLECTION DBA D", "THE GIRLIE BOUTIQUE", "THE PINK LILY BOUTIQ", "THREADS BOUTIQUE", "TY ALEXANDERS", "UNIQUE BOUTIQUE-SUTT", "UNIVERSITY BOOK STOR", "WAKEFIELDS-INC", "WREN AND IVORY LLC"]


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


	; 아이템 이름순으로 정렬. 1열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(1), 1)
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	


	; 주문 날짜순으로 정렬. 8열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(8), 1)
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
	




	; 메뉴를 위한 첫째 행 삽입 후 각 열에 해당되는 메뉴 설명 넣기
	XL_Row_Insert(XL,RG:="1",HT:=13)  ;~ XL_Row_Insert(XL,RG:="1")
	Xl.Range("A1").Value := "Style Numbers"
	Xl.Range("B1").Value := "No"
	Xl.Range("C1").Value := "S.O #"
	Xl.Range("D1").Value := "Status"
	Xl.Range("E1").Value := "Customer"
	Xl.Range("F1").Value := "Cust. Priority"
	Xl.Range("G1").Value := "Order Type"
	Xl.Range("H1").Value := "Order Dt"
	
	
	
	; 프로그래스 바 닫기
	Gui Destroy	
	

	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 262144, Title, IT'S DONE`n`nA열 옮기기, 주문 날짜순으로 정렬 되어 있음.




ExitApp








; ### 티파니 고객 찾아서 삭제하기 ###
^#!t::

Xl := ComObjActive("Excel.Application")	; 열려있는 엑셀 창 사용하기

; Constants
xlValues := -4163
xlWhole := 1


TiffanyOnlyCustomerList := ["AHN AND AHN COLLECTI", "ALLURE", "AMBIANCE-PLACERVILLE", "AUBURN UNIVERSITY", "BETSEYS BOUTIQUE SHO", "BETTY BE GOOD BOUTIQ", "BOHO BLU", "BOOP DE DOOP", "CHARLI ROSE-CO", "COBOS BOUTIQUE", "COBOS-STARKVILLE", "ENTOURAGE", "FINNLEYS", "FINNLEYS 2", "FLAUNT BOUTIQUE-ROUN", "FOCUSED HOLDINGS", "JOSIES BOUTIQUE", "JUST JEWELRY", "K-S WHOLESALE", "KIKILARUE", "MAGNOLIA BTQ FRANKLI", "ROOLEE", "SHOP DRESS UP", "STYLEFOX", "THE COLLECTION DBA D", "THE GIRLIE BOUTIQUE", "THE PINK LILY BOUTIQ", "THREADS BOUTIQUE", "TY ALEXANDERS", "UNIQUE BOUTIQUE-SUTT", "UNIVERSITY BOOK STOR", "WAKEFIELDS-INC", "WREN AND IVORY LLC"]


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




Esc::
ExitApp
