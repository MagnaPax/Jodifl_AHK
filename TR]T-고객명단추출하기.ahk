; ######################################################################
; Allocation 파일 보기 편하게 쓸데 없는 값 지운 뒤 중복되는 SO 번호 지우기
; ######################################################################

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\


#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk




	MsgBox, 262144, REFINE ALLOCATION FILE, CLICK OK TO CONTINUE


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
	
	
;~ /*	
	; 아이템 찾는 동안 보여줄 프로그래스 바
	TotalLoops = 50000
	Gui, -Caption +AlwaysOnTop +LastFound
	Gui, Add, Text, x12 y9 w170 h20 , P  R  O  C  E  S  S  I  N  G  .  .  .
	Gui, Add, Progress, w410 Range0-%TotalLoops% vProgress
	Gui, Show, w437 h84, SEARCHING ITEMS
*/





/*


	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible		
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가


	;엑셀 값의 끝 row 번호 알아낸 후 end#ofTheRow 에 값 넣기
	XL_Handle(XL,1) ;get handle to Excel Application
	end#ofTheRow := XL_Last_Row(XL)
	lastLine#OfRow := XL_Last_Row(XL)
	GuiControl,,Progress, +1 ; 프로그래스 바 1씩 증가
	
;MsgBox, % "end#ofTheRow : " . end#ofTheRow




	oSheet := Xl.Worksheets.Add()		; add a new Sheet
	oSheet.Name := "Data1"	; name new sheet to "Data1"
	
	Xl.Worksheets("Data1").Range("A1").Value := "Boutique Name"
	Xl.Worksheets("Data1").Range("B1").Value := "Contact"
	Xl.Worksheets("Data1").Range("C1").Value := "Phone#"
	Xl.Worksheets("Data1").Range("D1").Value := "Email"
	Xl.Worksheets("Data1").Range("E1").Value := "State"
	
	
	; 열의 넓이 설정하기
	XL_Col_Width_Set(XL,RG:="A=25|B=25|C=25|D=45|E=15")
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	
	
	
	i = 7
	i_Data = 2
	Loop{
		
		;~ CellNumberCurrentlyBeingBeenHandled := "A" . i
		;~ CellNumberBeingHandledNow := "J" . i
		
		Cell#OfCustName := "J" . i		
		
		customerName := Xl.Range(Cell#OfCustName).Value		; Cell#OfCustName 셀에 있는 값 customerName 변수에 저장하기
		customerName := Xl.Worksheets("Sheet1").Range(Cell#OfCustName).Value		; Cell#OfCustName 셀에 있는 값 customerName 변수에 저장하기

		
		j := i+2
		
		
		Cell#OfPhone# := "L" . j		
		phone# := Xl.Worksheets("Sheet1").Range(Cell#OfPhone#).Value		; Cell#OfPhone# 셀에 있는 값 phone# 변수에 저장하기
		
		
		k := i+3
		
		Cell#OfContact := "N" . k		
		contact := Xl.Worksheets("Sheet1").Range(Cell#OfContact).Value	; Cell#OfContact 셀에 있는 값 customerName 변수에 저장하기
		
		
		Cell#OfEmail := "R" . k
		email := Xl.Worksheets("Sheet1").Range(Cell#OfEmail).Value	; Cell#OfEmail 셀에 있는 값 email 변수에 저장하기
		
		
		l := i+4
		
		Cell#OfState := "F" . l
		state := Xl.Worksheets("Sheet1").Range(Cell#OfState).Value	; Cell#OfState 셀에 있는 값 state 변수에 저장하기


		i := i+8
		
;MsgBox, % "i := " i . "`nend#ofTheRow : " . end#ofTheRow . "`ncustomerName : " . customerName . "`nphone# : " . phone# . "`nemail : " . email . "`nstate : " . state

		
		
		; Data1 탭에 원본 탭(Sheet1)에서 읽어온 값들 쓰기
		Loop{
			
			Data1_Cell#OfCustName := "A" . i_Data
			Data1_Cell#OfContact := "B" . i_Data
			Data1_Cell#OfPhone# := "C" . i_Data
			Data1_Cell#OfEmail := "D" . i_Data
			Data1_Cell#OfState := "E" . i_Data
			
			
			Xl.Worksheets("Data1").Range(Data1_Cell#OfCustName).Value := customerName
			Xl.Worksheets("Data1").Range(Data1_Cell#OfContact).Value := phone#
			Xl.Worksheets("Data1").Range(Data1_Cell#OfPhone#).Value := contact
			Xl.Worksheets("Data1").Range(Data1_Cell#OfEmail).Value := email
			Xl.Worksheets("Data1").Range(Data1_Cell#OfState).Value := state
			
			
			GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	
			
			i_Data++
			break
		}
		
		if(end#ofTheRow <= i){
			break
		}		
	}
	
	
	
	Gui Destroy
	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 262144, Title, IT'S DONE	
	
	
	MsgBox, 다 옮겼음
	
	
*/	
	

















; #############################################################
;	MsgBox, 262144, Title, 고객명(2열)로 정렬 후 중복값 지우기
; #############################################################	


	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible		
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
	
	; 중복값 지우기 위해 일단
	; 2 열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(4), 1)



	;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
	XL_Handle(XL,1) ;get handle to Excel Application
	i := XL_Last_Row(XL)


;~ MsgBox
		
		j = 2
		Loop{

			k := j + 1

			#ofRowToRead := "D" . j
			#ofRowToReadAddedOne :=  "D" . k
			
			
			var1 := Xl.Range(#ofRowToRead).Value
			var2 := Xl.Range(#ofRowToReadAddedOne).Value
			
;			MsgBox, % var1 . "`n" . var2
			
			
			if(var1 == ""){
				break
			}
						

			; 만약 지금 얻은 SO# 값이 이전 SO# 값을 저장하고 있는 previousNumber 값과 같다면 
			; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
			IfEqual, var1, %var2%
			{
				
;				MsgBox, % var1 . "`n" . var2 . "`n`n" . "delete" . "`n`n" . "i : " . i . "`nj : " . j
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
		
		
	Gui Destroy
	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 262144, Title, IT'S DONE	

	
	

	Exitapp

	Esc::
	Exitapp	
	
		
	
	