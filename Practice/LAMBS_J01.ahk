
i := 1

; 인보이스에 뭔가 쓰여져 있으면 체크할 수 있는 체크박스 만들까
; 이게 체크되면 무조건 진행되는게 아니라 중간에 멈춰서 노트에 적힌것 확인하게끔

;GUI Backgroud
Gui, Show, w327 h137, JODIFL
Gui, Add, Button, x212 y19 w100 h80 gClick_btn, Enter

;Input invoice Number
;Gui, Add, Text, x22 y19 w70 h20 , Invoice No.
Gui, Add, Text, x22 y19 Cred , Invoice No.
Gui, Add, Edit, x92 y19 w100 h20 vInvoice_No, 

;Input Wts of boxes
;Gui, Add, Text, x22 y49 w70 h20 , Wts
Gui, Add, Text, x22 y49 CBlue , Wts
Gui, Add, Edit, x92 y49 w100 h20 vInvoice_Wts, 

;Input # of boxes
Gui, Add, Text, x22 y79 w70 h20 , No. of Box
Gui, Add, Edit, x92 y79 w100 h20 vH1,
Gui, Add, UpDown, x172 y79 w20 vNo_of_Boxes, 1

return


Click_btn:

	; To use the values which input on GUI
	Gui Submit, nohide


	;Activate LAMBS Window
	
	Loop,
	IfWinExist, LAMBS -  Garment Manufacturer & Wholesale Software
	break

	WinActivate, LAMBS -  Garment Manufacturer & Wholesale Software

	sleep 1000
	;return
	
	
	/*
	Gui Submit, nohide
	  
	msgbox, %N1%
	msgbox, %N2%
	msgbox, %H1%
	;msgbox, %No_of_Boxes%
	  
	return
	*/	  

	
	; Find New & Clear Button then double click it
	while(i = 1)
	{
		
		ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\00New_Clear.png
		if(errorlevel = 0){
			MsgBox, good
			
			MouseGetPos, %pX%, %pY%
			pX += 20
			pY += 10
			MouseClick, l, %pX%, %pY%, , d
			
			;SendInput, %Click_btn%
			
			/*
			MouseMove, %pX%, %pY%
			MsgBox, %pX%, %pY%
			
			MouseMove, %pX%, %pY%
			MsgBox, %pX%, %pY%
			
			;MouseClick, l, %pX%, %pY%, 11, d
			;MouseClick,
			*/
			i := 0
			;return
		} else if(errorlevel = 1){
			MsgBox, bad
		} else if(errorlevel = 2){
			MsgBox, an error
		}
		sleep 1000
		;return
	}
	
	
	
	; Find Invoice No Button then Put the invoice number in it
	
	ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\01Invoice_No.png
	if(errorlevel = 0){
		MsgBox, good2
		
		MouseGetPos, %pX%, %pY%
		
		pX += 90
		pY += 10
		
		MouseClick, l, %pX%, %pY%, , d
		sleep 500
		
		;SendInput, 67629
		SendInput, %Invoice_No%
		SendInput, {Enter}
		
		return
	} else if(errorlevel = 1){
		MsgBox, bad
	} else if(errorlevel = 2){
		MsgBox, an error
	}
	sleep 5000
	;return
	
	
		
	; Find Security Button and Click it to release LAMBS
	ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\02Security_Button.png
	if(errorlevel = 0){
		MsgBox, good3
		
		MouseGetPos, %pX%, %pY%
		
		pX += 20
		pY += 20
		
		MouseClick, l, %pX%, %pY%, , d
		sleep 500
		
		SendInput, 1
		SendInput, {Enter}

		return
	} else if(errorlevel = 1){
		MsgBox, bad
	} else if(errorlevel = 2){
		MsgBox, an error
	}
	sleep 1000
	
	
	; Find The black of Invoice Wts to put in Wts of boxes
	ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\03Invoice_Wts.png
	if(errorlevel = 0){
		MsgBox, good4
		
		MouseGetPos, %pX%, %pY%
		
		pX += 120
		pY += 5
		
		MouseMove, %pX%, %pY%
;		sleep 500
		
		MouseClick, l, %pX%, %pY%, 2, d

;		sleep 500
		
		SendInput, %Invoice_Wts%
		SendInput, {Enter}

		return
	} else if(errorlevel = 1){
		MsgBox, bad
	} else if(errorlevel = 2){
		MsgBox, an error
	}
	sleep 1000


	; Find The black of No of Box to put in Number of boxes
	ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\04No_Box.png
	if(errorlevel = 0){
		MsgBox, good5
		
		MouseGetPos, %pX%, %pY%
		
		pX += 120
		pY += 5
		
		MouseMove, %pX%, %pY%
;		sleep 500
		
		MouseClick, l, %pX%, %pY%, 2, d

;		sleep 500
		
		SendInput, %No_of_Boxes%
		SendInput, {Enter}
		

		return
	} else if(errorlevel = 1){
		MsgBox, bad
	} else if(errorlevel = 2){
		MsgBox, an error
	}
	sleep 100
	








GuiClose:
ExitApp
  
  
  
F4::
 Exitapp
 return

