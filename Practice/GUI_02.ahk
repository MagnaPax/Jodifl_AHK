

MsgBox, 260, Add Items to, Would you like to transfer Items from Sales Order?
IfMsgBox, Yes
{
	run, notepad
	return
}

MsgBox, 292, AA, BB

/*
Gui, Add, CheckBox, x62 y279 w100 h30 vApplyCredit +Center, CheckBox
; Generated using SmartGUI Creator for SciTE
Gui, Show, w479 h379, Untitled GUI

Gui, Add, Button, x50 y19 w100 h80 gClick_btn, Enter
return


Click_btn:
		Gui Submit, nohide
if(ApplyCredit = 1)
	MsgBox, Apply Credit 체크 되어있음
else if(ApplyCredit = 0)
	MsgBox, Apply Credit 체크 되어있지 않음
	
*/
GuiClose:
ExitApp