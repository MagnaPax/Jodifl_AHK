/*
===========================================

  YouTube Demonstration(이 함수의 유튜브 시연 영상 주소)
  https://www.youtube.com/watch?v=Q7LGKiWM9z4
  
  http://the-automator.com/insert-delete-columns-in-excel-via-autohotkey/

===========================================

*/


#Include [Excel]_Joe Glines'sExcelFunctions.ahk



;아래 함수들 실행부분
/*
#SingleInstance,Force
RAlt::Reload
RControl::
XL_Handle(XL,1) ;get handle to Application
 
XL_Col_Width_Set(XL,RG:="A:B=-1|D:F=-1|H=15|K=3") ;-1 is auto
XL_Col_Insert(XL,RG:="B:C",WD:=8)   ;insert column(s) and set width
XL_Col_Delete(XL,RG:="A|D:F|Z|BB:BC") ;Delete columns
return
*/



;***********column width*******************
XL_Col_Width_Set(PXL,RG=""){
 for k, v in StrSplit(rg,"|") ;Iterate over array
    (StrSplit(v,"=").2="-1")?(PXL.Application.ActiveSheet.Columns(StrSplit(v,"=").1).AutoFit):(PXL.Application.ActiveSheet.Columns(StrSplit(v,"=").1).ColumnWidth:=StrSplit(v,"=").2)
}
 
;***********************Column Insert********************************.
XL_Col_Insert(PXL,RG="",WD:="5"){ ;Default width is 5
PXL.Application.ActiveSheet.Columns(RG).Insert(-4161).Select
PXL.Application.ActiveSheet.Columns(RG).ColumnWidth:=WD
}
 
;***********************Column Delete********************************.
XL_Col_Delete(PXL,RG=""){
    for j,k in StrSplit(rg,"|")
 (instr(k,":")=1)?list.=k ",":(list.=k ":" k ",") ;need to make for two if only 1 col
PXL.Application.ActiveSheet.Range(SubStr(list,1,(StrLen(list)-1))).Delete ;use list but remove final comma
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