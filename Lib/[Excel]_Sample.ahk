#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



/*
; ####################################### 참조 #######################################
https://autohotkey.com/board/topic/56987-com-object-reference-autohotkey-v11/page-4#entry381256

https://autohotkey.com/board/topic/69033-basic-ahk-l-com-tutorial-for-excel/

https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/range-offset-property-excel
*/


	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible		
	
	;~ Xl.Save
	Xl.ActiveWorkbook.save()




	;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
	;~ XL_Handle(XL,1) ;get handle to Excel Application
	;~ i := XL_Last_Row(XL)


/*
FullPath := "C:\Book1.xlsx"	; please adjust full path to your Workbook...
oWorkBook := ComObjGet(FullPath)		; get reference to WorkBook
oWorkbook.Application.Windows(oWorkbook.Name).Visible := 1	; just do it - too long to explain why...

oWorkbook.Worksheets("Sheet1").Range("A2").Value := "Test"	; set "Test" in "Sheet1" sheet, "A2" cell
oSheet := oWorkbook.Worksheets.Add()		; add a new Sheet
oSheet.Name := "Data1"	; name new sheet to "Data1"
oWorkbook.Worksheets("Data1").Range("A2").Value := "Test 2"	; set "Test 2" in "Data1" sheet, "A2" cell
oWorkbook.Close(1)	; save changes and close Workbook
oWorkBook := "", oSheet := ""	; release references
ExitApp
*/

/*
; 영역 복사1
!h::
XL := ComObjCreate("Excel.Application")
XL.Workbooks.Open("C:\Book2.xlsx")
XL.Workbooks.Open("C:\Book1.xlsx")
XL.Visible := True

XL.Range("A1:B5").Copy
XL.Windows("Book2.xlsx").Activate
XL.Range("A10").Select
XL.ActiveSheet.Paste
return
*/


/*
; 영역 복사2
sheetOne.range("a1").copy
sheetTwo.range("a2").select
xlTwo.ActiveSheet.Paste
xlOne.Application.CutCopyMode := False
*/












	Exitapp

	Esc::
	Exitapp	
	