/*
===========================================

  YouTube Demonstration(이 함수의 유튜브 시연 영상 주소)
  https://youtu.be/zyMPxe2vISE

===========================================

*/


#Include [Excel]_Joe Glines'sExcelFunctions.ahk



/*
; 실행 부분

XL_Handle(XL,1) ;get handle to Excel Application
first_row_of_data:=XL_First_Row(XL)+1 ;First row
MsgBox % "first row of data: " first_row_of_data
MsgBox % "last row: " XL_Last_Row(XL)  ;Last row
MsgBox % "used rows: " XL_Used_Rows(XL) ;Used Rws
return
*/



XL_First_Row(PXL){
 Return, PXL.Application.ActiveSheet.UsedRange.Rows(1).Row ;first used row in Excel
}
 
XL_Last_Row(PXL){
 Return, PXL.Application.ActiveSheet.UsedRange.Rows(PXL.Application.ActiveSheet.UsedRange.Rows.Count).Row
}
 
XL_Used_rows(PXL){
 Return,PXL.Application.ActiveSheet.UsedRange.Rows.Count
}



; 공통되는 부분이라 [Excel]_Joe Glines'sExcelFunctions.ahk 파일로 옮겨놨음
/*
;  Function for connecting to Excel
XL_Handle(ByRef PXL,Sel){
ControlGet, hwnd, hwnd, , Excel71, ahk_class XLMAIN ;identify the hwnd for Excel
IfEqual,Sel,1, Return, PXL:= ObjectFromWindow(hwnd,-16).application ;Handle to Excel Application
IfEqual,Sel,2, Return, PXL:= ObjectFromWindow(hwnd,-16).parent ;Handlle to active Workbook
IfEqual,Sel,3, Return, PXL:= ObjectFromWindow(hwnd,-16).activesheet ;Handle to Active Worksheet
}
;***********adapted from ACC.ahk*******************
ObjectFromWindow(hWnd, idObject = -4){
(if Not h)?h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
 If DllCall("oleacc\AccessibleObjectFromWindow","Ptr",hWnd,"UInt",idObject&=0xFFFFFFFF,"Ptr",-VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
 Return ComObjEnwrap(9,pacc,1)
}
*/