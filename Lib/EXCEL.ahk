#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include function.ahk
#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk


#Include CommEXCEL.ahk



class EXCEL extends CommEXCEL{


	; 엑셀에서 값 읽고 배열에 저장
	GetInfoFromExcelThenPutThatInAArrayForTF(){
		
		Str_ExcelInfo := object()
		Str_Total := object()



		; 엑셀 창이 열려있지 않으면 열릴때까지 경고창 표시하기
		CommEXCEL.CheckingPresenceOfExcelFile()


		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible
		
		
		
		loop{
			
			RefinedOrderID = 
			
			; Order ID
			CustPO := Xl.Range("D1").Value		
			RegExMatch(CustPO, "imU)(\d*)\.", SubPat) ; 소수점 뒷자리 정리
			RefinedOrderID := SubPat1 ; 정리된 값 RefinedOrderID 에 넣기
			
			; 엑셀 파일 끝난 것 탐지하기
			if(RefinedOrderID == ""){
				MsgBox, 262144, No Date, ORDER FINISHED`n`nTHERE IS NO DATA IN THE EXCEL FILE`n`n%RefinedOrderID%
				ExitApp
			}
			
			

			; 지금 읽은 Order ID가 이전 값과 다르다면 같은 주문이 아니므로 루프 탈출
			IF(A_Index -1 != 0)
				IfNotEqual, RefinedOrderID, %previousNumber%
					break	
			

			
			; Customer PO Number
			PO# := Xl.Range("E1").Value
			PO# := RegExReplace(PO#, "\(ID:.*\)", "") ; 뒤에 붙은 (ID:) 제거
			PO# := Trim(PO#)
			StringUpper, PO#, PO# ; 대문자로 바꾸기
			
			; Customer Name
			Name := Xl.Range("F1").Value
			Name := RegExReplace(Name, "\(ID:.*\)", "")
			Name := Trim(Name)

			; Style
			Style := Xl.Range("H1").Value
			;~ Style := RegExReplace(Style, "[^a-zA-Z0-9]", "") ; 알파벳과 숫자만
			Style := Trim(Style)

			; Color
			Color := Xl.Range("I1").Value
			Color := Trim(Color)
			
			; Order Qty
			Qty := Xl.Range("O1").Value
			Qty := Qty/6 ; 낱장이 아닌 팩이니까 6으로 나누기
			RegExMatch(Qty, "imU)(\d*)\.", SubPat) ; 소수점 뒷자리 정리			
			Qty := SubPat1 ; 정리된 값 RefinedOrderID 에 넣기
			
			
			; EMAIL
			Email := Xl.Range("V1").Value
			Email := Trim(Email)
			
			if(Email == ""){
				MsgBox, 262144, No Date, ORDER FINISHED`n`nTHERE IS NO EMAIL ADDRESS`n`nIF YOU CLICK OK, THE APPLICATION WILL BE RELOADED
				Reload
			}
			
			
			Memo := Xl.Range("BK1").Value
			Memo := RegExReplace(Memo, "imU)Handling\sFee:\s.0.00", "")
			StringReplace, Memo, Memo, `r`n, , All ; 모든 CR+LF를 제거합니다:
			Memo := Trim(Memo, Memo = " `r")
			Memo := Trim(Memo, Memo = " `n")
			Memo := Trim(Memo)
			;~ MsgBox, % "||" . Memo . "||"
/*			
			if Memo contains VOID
			{
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Alert, CHECK THIS ORDER IF IT'S VOIDED ORDER OR NOT`n`nIF YOU'D LIKE TO CONTINUE THIS ORDER, CLICK YES.`n`nCLICKED NO WILL BE RELOAD THIS APPLICATION
				
				IfMsgBox, No
				{
					Reload
				}

			}
*/			

			; 주문 장소
			OrderLocation := Xl.Range("AT1").Value
			OrderLocation := Trim(OrderLocation)
			
			if OrderLocation contains FASHIONGO
				OrderLocation = F
			else if OrderLocation contains LASHOWROOM
				OrderLocation = L
			else if OrderLocation contains WEB
				OrderLocation = WEB			
			else if OrderLocation contains EMAIL
				OrderLocation = EMAIL
			else if OrderLocation contains PHONE
				OrderLocation = PHONE



			Str_ExcelInfo[1] := RefinedOrderID
			Str_ExcelInfo[2] := PO#
			Str_ExcelInfo[3] := Name
			Str_ExcelInfo[4] := Style
			Str_ExcelInfo[5] := Color
			Str_ExcelInfo[6] := Qty
			Str_ExcelInfo[7] := OrderLocation
			Str_ExcelInfo[8] := Email
			Str_ExcelInfo[9] := Memo

			


			; 첫 번째 Row 값은 배열에 넣었으니 엑셀에서 지워주기
			Xl.Sheets(1).Range("A1").EntireRow.Delete
			
			; 중복되는 값의 비교를 위해 previousNumber 변수에 RefinedOrderID 값 넣기
			previousNumber := RefinedOrderID
			
			

			
			
			Str_Total.Insert(Str_ExcelInfo) ; Str_Total안에 차곡차곡 넣기
			
			Str_ExcelInfo := [] ; Str_Total안에 넣을때마다 값이 중복되지 않게 Str_ExcelInfo 배열 초기화 해주기

		}


		return Str_Total
		
		


			MsgBox out
			i = 1
			j = 1
			;~ /* 배열로부터 읽기 첫 번째 방법
			Loop % Str_Total.Maxindex(){
				Loop, 6{
					MsgBox % "Element number " . A_Index . " is " . Str_Total[i][j] . "||"
					j++
				}
				
				i++
				j = 1
			}
			*/
			










		
		return Str_Total
	}
	
	
	
	
	
	
	
	

	; 아이템 업데이트 위해 배송정보 들어있는 Invoice && C.M Detail Register 파일에서 빈줄 삭제하고 쓸데없는 열(column) 지우기
	DeleteEmptyColumnToUpdateItems(){

		; 엑셀 창이 열려있지 않으면 열릴때까지 경고창 표시하기
		CommEXCEL.CheckingPresenceOfExcelFile()


		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible
		
		
		
		
		; 아이템 찾는 동안 보여줄 프로그래스 바 
		TotalLoops = 3000
		Gui, -Caption +AlwaysOnTop +LastFound
		Gui, Add, Text, x12 y9 w170 h20 , P  R  O  C  E  S  S  I  N  G  .  .  .
		Gui, Add, Progress, w410 Range0-%TotalLoops% vProgress
		Gui, Show, w437 h84, SEARCHING ITEMS

		
		
		
		
		
	
		i = 1
		loop{
			
			CellNumberCurrentlyBeingBeenHandled := "A" . i
			
			; A1 셀에 있는 값 IsItEmptyOrNot 변수에 저장하기
			;~ IsItEmptyOrNot := Xl.Range("A1").Value
			IsItEmptyOrNot := Xl.Range(CellNumberCurrentlyBeingBeenHandled).Value
			
			
			; 맨 마지막 A1 셀에는 'Grand Total' 값이 있기 때문에 그것이 나오면 여러가지 처리한 뒤 메소드 끝내기
			if IsItEmptyOrNot contains Grand
			{
				
				; 맨 마지막 줄 삭제하기
				Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
				
				; 맨 마지막 이전 줄 삭제하기
				CellNumberCurrentlyBeingBeenHandled := "A" . --i
				Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
				
				; 맨 처음에 남은 제목 줄 삭제하기
				Xl.Sheets(1).Range("A1").EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
				
				; 쓸 데 없는 행 삭제하기
				;~ XL_Col_Delete(XL,RG:="D:F|H|K|N:R") ;Delete columns
				XL_Col_Delete(XL,RG:="C|E:N|P:AI|AK:AX|AZ:BX|BZ:CC") ;Delete columns
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
				
				
				; Tracking # 열 넓이 설정하기
				XL_Col_Width_Set(XL,RG:="A=10|C=8|D=30|F=15|G=20")
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	



				; Tracking # 크기대로 정렬하기
				; 7 열(Columns)을 정렬하기
				xl.cells.sort(xl.columns(7), 1)
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
			
				; Tracking # 값이 들어있는 행 루프 돌아가며 모두 지우기
				Loop{
					
					; G 열의 첫 번째 행의 값, 즉 첫번째 Tracking # 값을 var1 변수에 넣기					
					var1 := Xl.Range("G1").Value
					
					; Tracking # 가 있다면 지우기
					if(var1){
						Xl.Sheets(1).Range("G1").EntireRow.Delete
						GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
					}
					; 다 지웠으면 루프 빠져나오기
					else
						break
				}
				
				
				; CUSTOER PO 로 정렬하기
				; 6 열(Columns)을 정렬하기
				xl.cells.sort(xl.columns(6), 1)
				
				

	/*			; 원하는 폴더에 원하는 이름으로 저장하기
				; https://autohotkey.com/board/topic/94976-save-as-in-excel/
				book_name := Xl.ActiveWorkbook.name
				SplitPath, book_name,,, ext
				;~ Xl.ActiveWorkbook.SaveAs(directory "\" new_file_name "." ext)
				Xl.ActiveWorkbook.SaveAs(%A_ScriptDir%\CreatedFiles "\" new_file_name "." xls)			
				
	*/

				Gui Destroy
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, IT'S DONE				
						
						
				return
			}
			
			

			; A1에 값이 없으면 그 줄 삭제하기
			if(IsItEmptyOrNot == ""){
				;~ Xl.Sheets(1).Range("A1").EntireRow.Delete
				Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
			}
			else
				i++



		}

		return
	}
	

	
	
	
	
	
	
	
	
	
		
	
	; 메모에 void 들어있는 것만 남기고 다 지우기
	DeleteColumnExceptOnesHavingVoidMemo(){
		

		; 엑셀 창이 열려있지 않으면 열릴때까지 경고창 표시하기
		CommEXCEL.CheckingPresenceOfExcelFile()


		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible

		i = 1
		loop{

		CellNumberCurrentlyBeingBeenHandled := "BK" . i

		CheckingEndOfExcel := "F" . i

		CustomerName := Xl.Range(CheckingEndOfExcel).Value


		; BK 셀에 있는 값 IsItEmptyOrNot 변수에 저장하기
		;~ IsItEmptyOrNot := Xl.Range("A1").Value
		IsItEmptyOrNot := Xl.Range(CellNumberCurrentlyBeingBeenHandled).Value

		;	MsgBox, % "IsItEmptyOrNot : " . IsItEmptyOrNot . "`n`n" . "BK" . i . "`n`n" . "CustomerName : " . CustomerName

		IF IsItEmptyOrNot not contains void
		{
		Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
		--i
		}

		if(CustomerName == "")
		{
		MsgBox, End of excel
		return
		}


		i++


		}

		return
	}

		
	












	; 엑셀에서 정보 읽기. 같은 고객의 모든 정보가 다중배열로 반환됨
	GetInfoFromExcelThenPutThatInAArrayToIU(){
		
		Str_ExcelInfo := object()
		Str_Total := object()



		; 엑셀 창이 열려있지 않으면 열릴때까지 경고창 표시하기
		CommEXCEL.CheckingPresenceOfExcelFile()


		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible
		
		
		
		loop{			

			; Cust PO
			CustPO := Xl.Range("F1").Value


;			MsgBox, % "previousNumber  " . previousNumber . "`n`n" . "CustPO  " . CustPO . "`n`n" . "A_Index -1`n`n" . A_Index -1


			; 지금 읽은 Cust PO가 이전 값과 다르다면 같은 주문이 아니므로 배열 반환하면서 메소드 끝내기
			IF(A_Index -1 != 0){
				;~ IfNotEqual, CustPO, %previousNumber%
				;~ IfNotEqual, CustPO, previousNumber
				if(CustPO != previousNumber)
				{
;					MsgBox, FINISH READING INFO FROM SAME CUST PO
					return Str_Total
				}
			}



			
			; 엑셀 파일 끝난 것 탐지하기
			if(CustPO == ""){
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 262144, No Date, ORDER FINISHED`n`nTHERE IS NO DATA IN THE EXCEL FILE`n`n%CustPO%
				ExitApp
			}
			
			



			; Customer Name
			Name := Xl.Range("D1").Value
			;~ Name := RegExReplace(Name, "\(ID:.*\)", "")
			Name := Trim(Name)

			; Style
			Style := Xl.Range("A1").Value
			;~ Style := RegExReplace(Style, "[^a-zA-Z0-9]", "") ; 알파벳과 숫자만
			Style := Trim(Style)

			; Color
			Color := Xl.Range("B1").Value
			Color := Trim(Color)
			
			; Invoice Number
			Invoice_# := Xl.Range("C1").Value
			Invoice_# := Trim(Invoice_#)
			RegExMatch(Invoice_#, "imU)(\d*)\.", SubPat) ; 소수점 뒷자리 정리
			Invoice_# := SubPat1 ; 정리된 값 다시 Invoice_# 에 넣기			
			
			
			; Tracking Number
			Tracking_# := Xl.Range("G1").Value
			
			
			; Total Qty
			Total_Qty := Xl.Range("E1").Value
			RegExMatch(Total_Qty, "imU)(\d*)\.", SubPat) ; 소수점 뒷자리 정리
			Total_Qty := SubPat1 ; 정리된 값 다시 Total_Qty 에 넣기
			
;			MsgBox, % "Total_Qty : " . Total_Qty


			Str_ExcelInfo[1] := CustPO
			Str_ExcelInfo[2] := Invoice_#
			Str_ExcelInfo[3] := Tracking_#
			Str_ExcelInfo[4] := Style
			Str_ExcelInfo[5] := Color
			Str_ExcelInfo[6] := Total_Qty
			;~ Str_ExcelInfo[7] := ShippingFee ; 이 값은 엑셀에 안 들어있음

			


			; 첫 번째 Row 값은 배열에 넣었으니 엑셀에서 지워주기
			Xl.Sheets(1).Range("A1").EntireRow.Delete

			
			; 만약 FG 나 LAS 오더가 아니면 메인으로 값을 넘기지 말고 처음으로 가서 다른 오더값 읽기
			if CustPO not contains MTR
			{
				if CustPO not contains OP
				{
;					MsgBox, % "This Cust PO is neither FG nor LAS`n`n" . CustPO
					previousNumber := Xl.Range("E1").Value ; 이렇게 안 해주면 continue로 루프를 다시 올라가서 다음 PO를 읽을 때 previousNumber 값과 CustPO랑 달라서 다른 주문이라고 판단하여 이 메소드를 탈출한다. 그것을 방지하기 위해.
					continue
				}
			}		
			
			
			; 중복되는 값의 비교를 위해 previousNumber 변수에 CustPO 값 넣기
			previousNumber := CustPO
			
			
			
			Str_Total.Insert(Str_ExcelInfo) ; Str_Total안에 차곡차곡 넣기
			
			Str_ExcelInfo := [] ; Str_Total안에 넣을때마다 값이 중복되지 않게 Str_ExcelInfo 배열 초기화 해주기

		}


		

/*
			; 다중 배열 Str_Total에 값이 제대로 잘 들어갔는지 확인하는 동작
			MsgBox out
			i = 1
			j = 1
			
			Loop % Str_Total.Maxindex(){
				Loop, 6{
					MsgBox % "Element number " . A_Index . " is " . Str_Total[i][j] . "||"
					j++
				}
				
				i++
				j = 1
			}
*/
			

		
		return Str_Total
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}