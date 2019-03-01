/*
===========================================

  YouTube Demonstration(이 함수의 유튜브 시연 영상 주소)
  https://youtu.be/T_VlNym1CN4

===========================================

*/


#Include [Excel]_Joe Glines'sExcelFunctions.ahk



; 실행 부분
/*
#SingleInstance,Force

Clipboard:="1`r 2`r3`r4`r5`r6`r7`r8`r9`r10`r11`r12`r13`r14`r15`r16`r17`r18`r19"
^t::
SendEvent, ^v
return

RAlt::
Browser_Forward::Reload
RControl::
Browser_Back::

XL_Handle(XL,1) ;get handle to Application
XL_Row_Height(XL,RG:="1:4=-1|10:13=50|21=15") ;rows first then height -1 is auto
XL_Row_Insert(XL,RG:="2:4",HT:=26)  ;~ XL_Row_Insert(XL,RG:="1")
XL_Row_Delete(XL,RG:="4:5|9|67|10") ;can be out of order

;i = 1
;XL_Row_Delete(XL,RG:=i) ;can be out of order

;XL_Row_Delete(XL,RG:="1:8") ;can be out of order


return
*/



;***********************Row Height********************************.
XL_Row_Height(PXL,RG=""){
for k, v in StrSplit(rg,"|") ;Iterate over array
    (StrSplit(v,"=").2="-1")?(PXL.Application.ActiveSheet.rows(StrSplit(v,"=").1).AutoFit):(PXL.Application.ActiveSheet.rows(StrSplit(v,"=").1).RowHeight:=StrSplit(v,"=").2)
}
 
;***********************Row Insert********************************.
;~ XL_Row_Insert(XL,RG:="1:5",HT:=16)  ;~ XL_Row_Insert(XL,RG:="1")
XL_Row_Insert(PXL,RG="",HT:="15"){ ;default height is 15
PXL.Application.ActiveSheet.Rows(RG).Insert(-4161).Select
PXL.Application.ActiveSheet.Rows(RG).RowHeight:=HT
}
 
;***********************Row Delete*rows in Excel via AutoHotkey*******************************.
;~ XL_Row_Delete(XL,RG:="4:5|9|67|9|10") ;range or single but cannot overlap
XL_Row_Delete(PXL,RG=""){
for j,k in StrSplit(rg,"|")
    (instr(k,":")=1)?list.=k ",":(list.=k ":" k ",") ;need to make for two if only 1 Row
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