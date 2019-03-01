#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk
#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
	BlockInput, Mouse


returnMultipleArrays()
{
 Array1 := ["1", "2"]
 Array2 := ["3", "4"]
 Array3 := ["5", "6"]
 return [Array1, Array2, Array3]
}


GetInfoFromExcel(){
	
	Arr_Excel := object()
	
		
	; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
	IfWinNotExist, ahk_class XLMAIN
	{
		MsgBox, 262144, No Excel file Warning, Please Open an Excel File of BO list
		;~ continue
		GetInfoFromExcel()
	}

	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible
		

	loop{

		; Order ID 값은 D Column 에 있음
		; 앞에서 쓸 데 없는 값을 지워줬으니 D1 에 지금 사용 할 Order ID 값 있음
		RawOrderID := Xl.Range("D1").Value
		
		;~ MsgBox, % "STYLE NO `n" . Xl.Range("G1").Value . "`n`nColor`n" . Xl.Range("H1").Value . "`n`nQTY`n" . Xl.Range("O1").Value
			
		;소수점 뒷자리 정리
		RegExMatch(RawOrderID, "imU)(\d*)\.", SubPat)
			
		; 정리된 값 RefinedOrderID 에 넣기
		RefinedOrderID := SubPat1


		QTY := Xl.Range("O1").Value
		RegExMatch(QTY, "imU)(\d*)\.", SubPat)
		QTY := SubPat1


		; 지금 읽은 Order ID가 이전 값과 다르다면 같은 주문이 아니므로 루프 탈출
		IF(A_Index -1 != 0)
			IfNotEqual, RefinedOrderID, %previousNumber%
				break	


		Arr_Excel%A_Index% 									:= RefinedOrderID
		Arr_Excel%A_Index%_%A_Index% 						:=  Xl.Range("G1").Value		
		Arr_Excel%A_Index%_%A_Index%_%A_Index% 				:= Xl.Range("H1").Value
		Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index% 	:= QTY
		
		
		MsgBox, % Arr_Excel%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index%


		
		; 첫 번째 Row 값은 배열에 넣었으니 엑셀에서 지워주기
		Xl.Sheets(1).Range("A1").EntireRow.Delete
		
			
		; 중복되는 값의 비교를 위해 previousNumber 변수에 RefinedOrderID 값 넣기
		previousNumber := RefinedOrderID


	}
	
	MsgBox, loop out
	
	;~ Loop,% Arr_Excel.MaxIndex(){
		;~ Msgbox,% Arr_Excel[a_index]
		;~ MsgBox, % Arr_Excel%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index%
	;~ }
	
	Loop{
		
		if( Arr_Excel%A_Index% == "")
			break
		
		MsgBox, % Arr_Excel%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index%

	}
	
	MsgBox, loop out 2
	
	
	return [Arr_Excel]
	;~ return Arr_Excel
}
	




GetInfoFromExcel2(){
	
	Arr_Excel := object()
	
		
	; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
	IfWinNotExist, ahk_class XLMAIN
	{
		MsgBox, 262144, No Excel file Warning, Please Open an Excel File of BO list
		;~ continue
		GetInfoFromExcel()
	}

	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible
		

	loop{

		; Order ID 값은 D Column 에 있음
		; 앞에서 쓸 데 없는 값을 지워줬으니 D1 에 지금 사용 할 Order ID 값 있음
		RawOrderID := Xl.Range("D1").Value
		
		;~ MsgBox, % "STYLE NO `n" . Xl.Range("G1").Value . "`n`nColor`n" . Xl.Range("H1").Value . "`n`nQTY`n" . Xl.Range("O1").Value
			
		;소수점 뒷자리 정리
		RegExMatch(RawOrderID, "imU)(\d*)\.", SubPat)
			
		; 정리된 값 RefinedOrderID 에 넣기
		RefinedOrderID := SubPat1


		QTY := Xl.Range("O1").Value
		RegExMatch(QTY, "imU)(\d*)\.", SubPat)
		QTY := SubPat1


		; 지금 읽은 Order ID가 이전 값과 다르다면 같은 주문이 아니므로 루프 탈출
		IF(A_Index -1 != 0)
			IfNotEqual, RefinedOrderID, %previousNumber%
				break	


		; 작동되는 코드들
		;~ Arr_Excel%A_Index% 									:= RefinedOrderID
		;~ Arr_Excel%A_Index%_%A_Index% 						:=  Xl.Range("G1").Value		
		;~ Arr_Excel%A_Index%_%A_Index%_%A_Index% 				:= Xl.Range("H1").Value
		;~ Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index% 	:= QTY
		
		
		;~ MsgBox, % Arr_Excel%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index%

		Arr_OrderID%A_Index% := RefinedOrderID
		Arr_StyleNo%A_Index% :=  Xl.Range("G1").Value
		Arr_Color%A_Index% := Xl.Range("H1").Value
		Arr_Qty%A_Index% := QTY
		
;		MsgBox, % Arr_OrderID%A_Index% . "`n" . Arr_StyleNo%A_Index% . "`n" . Arr_Color%A_Index% . "`n" . Arr_Qty%A_Index%



		
		; 첫 번째 Row 값은 배열에 넣었으니 엑셀에서 지워주기
		Xl.Sheets(1).Range("A1").EntireRow.Delete
		
			
		; 중복되는 값의 비교를 위해 previousNumber 변수에 RefinedOrderID 값 넣기
		previousNumber := RefinedOrderID


	}
	
	MsgBox, loop out

	Loop{
		
		if( Arr_OrderID%A_Index% == "")
			break
		
		;~ MsgBox, % Arr_Excel%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index%
;		MsgBox, % Arr_OrderID%A_Index% . "`n" . Arr_StyleNo%A_Index% . "`n" . Arr_Color%A_Index% . "`n" . Arr_Qty%A_Index%

	}
	
	MsgBox, loop out 2
	
	
	return [Arr_OrderID, Arr_StyleNo, Arr_Color, Arr_Qty]
	
/*	
	Arr_OrderID%A_Index% := RefinedOrderID
	Arr_StyleNo%A_Index% :=  Xl.Range("G1").Value
	Arr_Color%A_Index% := Xl.Range("H1").Value
	Arr_Qty%A_Index% := QTY	
*/

}
	
GetInfoFromExcelAndPutThatInAArray(){
	return
}


































/*


	Arr_Excel := object()
	
	Loop{
		
		; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		IfWinNotExist, ahk_class XLMAIN
		{
			MsgBox, 262144, No Excel file Warning, Please Open an Excel File of BO list
			continue
		}

		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible
		
		
		loop{
			
			; Order ID 값은 D Column 에 있음
			; 앞에서 쓸 데 없는 값을 지워줬으니 D1 에 지금 사용 할 Order ID 값 있음
			RawOrderID := Xl.Range("D1").Value
			
			;소수점 뒷자리 정리
			RegExMatch(RawOrderID, "imU)(\d*)\.", SubPat)
			
			; 정리된 값 RefinedOrderID 에 넣기
			RefinedOrderID := SubPat1
			
			if(Arr_Excel[A_Index -1 ] != RefinedOrderID)
				break
			
			Arr_Excel[A_Index] := RefinedOrderID
			
			; 첫 번째 Row 값은 배열에 넣었으니 엑셀에서 지워주기
			Xl.Sheets(1).Range("A1").EntireRow.Delete
		}



		;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
		XL_Handle(XL,1) ;get handle to Excel Application
		i := XL_Last_Row(XL)
;		MsgBox % "last row: " XL_Last_Row(XL)  ;Last row
;		MsgBox, % i

		; 엑셀창 최소화하기
		;~ WinMinimize, ahk_class XLMAIN

		; 엑셀에 값이 들어간 만큼(i 값 만큼) 루프 돌면서 엑셀에서 값 읽기
		Loop, %i%{
		;Loop{
			
			
			
			; Order ID 값은 D Column 에 있음
			; 앞에서 쓸 데 없는 값을 지워줬으니 D1 에 지금 사용 할 Order ID 값 있음
			RawOrderID := Xl.Range("D1").Value
			
			;소수점 뒷자리 정리
			RegExMatch(RawOrderID, "imU)(\d*)\.", SubPat)
			
			; 정리된 값 RefinedOrderID 에 넣기
			RefinedOrderID := SubPat1
			
			Arr_Excel%i% := RefinedOrderID
			
			; 첫 번째 Row 값은 배열에 넣었으니 엑셀에서 지워주기
			Xl.Sheets(1).Range("A1").EntireRow.Delete
			
			
			
			
			RawOrderID2 := Xl.Range("D2").Value
			RegExMatch(RawOrderID2, "imU)(\d*)\.", SubPat)
			RefinedOrderID2 := SubPat1
			
			if(% Arr_Excel%i% == RefinedOrderID2)
				Arr_Excel[++i] := RefinedOrderID2
			
			i++
			
			; 에러 확인용
;			MsgBox, RefinedOrderID is : %RefinedOrderID%
			
			

			; 만약 RefinedOrderID 에 값이 없으면 파일이 끝났거나 이전에 사용했던 파일이 안 닫히고 열려있는 것이므로 파일 닫고 프로그램 다시 시작하기
			if(!RefinedOrderID)
			{				
				; 파일이 끝났으니 메세지 띄우고 프로그램 다시 시작하기
				MsgBox, 262144, Old File Notification, ALL ORDER ID NUMBERS HAVE BEEN PROCESSED`nPLEASE OPEN NEW BO LIST EXCEL FILE
				
				
				; 저장 않고 종료하는 법을 못 찾아서 그냥 일단 임시로 저장 후 바로 지우기		
				path = %A_ScriptDir%\CreatedFiles\temporary.xls
				XL.ActiveWorkbook.SaveAs(path) ;'path' is a variable with the path and name of the file you desire
				
				; 엑셀 종료하기
				;xL.ActiveWorkbook.SaveAs("testXLfile",56)               ;51 is an xlsx, 56 is an xls
				xl.WorkBooks.Close()                                    ;close file
				xl.quit
				
				; 방금 만든 파일 지우기
				FileDelete, %A_ScriptDir%\CreatedFiles\temporary.xls
				
				; 프로그램 재시작
				;~ Reload
			}
				
			
			
			
			
			; 만약 지금 얻은 RefinedOrderID 값이 이전 Order ID 값을 저장하고 있는 previousNumber 값과 같다면 
			; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
			IfEqual, RefinedOrderID, %previousNumber%		
			{
				;MsgBox, duplicated number
				Xl.Sheets(1).Range("A1").EntireRow.Delete
				continue
			}


	;		MsgBox, %RefinedOrderID%
			
			; 첫 번째 Row 값은 변수에 넣었으니 엑셀에서 지워주기
			Xl.Sheets(1).Range("A1").EntireRow.Delete
			
			
			; 중복되는 값의 비교를 위해 previousNumber 변수에 RefinedOrderID 값 넣기
			previousNumber := RefinedOrderID
			
			
			
			; RefinedOrderID 값을 넘겨주면서 BO_LAMBSProcessing 함수 호출하기
			BO_LAMBSProcessing(RefinedOrderID)




		}
	}
	
	
*/