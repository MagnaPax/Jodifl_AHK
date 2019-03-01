
#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk




; Path 에 위치한 엑셀창 열기
openTheExcelFile(Path, OOSStyle#, OOSColor){
	
	Xl := ComObjCreate("Excel.Application")
	Xl.Workbooks.Open(Path) ;open an existing file
	Xl.Visible := True
	
	return Xl	
}









; 필요 없는 열들 지우기
; 1before_StartRow# 유효값 보다 1 작은 값
; #oflasUselessRows 마지막에 몇 줄 지울지
deleteUselessRows(Xl, 1before_StartRow#, #oflasUselessRows){
	
	; 유효값이 나오기 전의 열까지 지우기
	Loop, %1before_StartRow#%{		
		Xl.Sheets(1).Range("A1").EntireRow.Delete
	}
;MsgBox, 처음의 필요 없는 줄 지웠음

	; 총 몇 줄(열)이 있는지 확인
	total#ofRows := XL_Last_Row(XL)
	

	delRow# := total#ofRows
	Loop{
		
		if(A_Index == #oflasUselessRows){
;MsgBox, 다 지웠음 루프 빠져나감
			break
		}
		
;MsgBox, % "A" . delRow# . " 지울 예정"
		Xl.Sheets(1).Range("A" . delRow#).EntireRow.Delete
		delRow#--
	}
	
	; 현재 활성화된 엑셀 저장
	Xl.ActiveWorkbook.save()
	
	
	return Xl
}









getInfoFronExcel(Xl, SO#_L, customerCode_L, custPO#_L, orderQTY_L){
	
	; 배열 선언
	SO#_Arr := object()
	customerCode_Arr := object()
	custPO#_Arr := object()
	orderQTY_Arr := object()
		
	; 배열 초기화 해주기
	SO#_Arr := []
	customerCode_Arr := []
	custPO#_Arr := []
	orderQTY_Arr := []	
	
	
	;~ Xl.Range("AC1").NumberFormat := "00"
	;~ Xl.range("E1").NumberFormat := "000000000"
	
	while(Xl.Range("A" . A_Index).value != "")
	{
		SO#_Arr[A_Index] := Xl.Range(SO#_L . A_Index).Value
		customerCode_Arr[A_Index] := Xl.Range(customerCode_L . A_Index).Value
		custPO#_Arr[A_Index] := Xl.Range(custPO#_L . A_Index).Value
		orderQTY_Arr[A_Index] := Xl.Range(orderQTY_L . A_Index).Value
	}
	
	; 값 다 얻었으니 엑셀창 닫기
	Xl.ActiveWorkbook.Close()
	
	return [SO#_Arr, customerCode_Arr, custPO#_Arr, orderQTY_Arr]
}



; 임시 메소드
tempMethod_ForTheWholeList(){
}













