F1::

i = 1

/*
	; Find New & Clear Button then double click it
	while(i = 1)
	{
		
		ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, 00New_Clear.png
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
			*//*
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
*/	

SetWorkingDir %A_ScripDir%

	global pX, pY, jpgLocation
	jpgLocation = %A_ScripDir%PICTURES\00New_Clear.png
	
	
	MsgBox, %jpgLocation%
	PicSearch()
	
	MouseMove, %pX%, %pY%
	pX += 10
	pY += 10
	
	sleep 1000
	MouseMove, %pX%, %pY%
	MsgBox, Return1

sleep 1000


	jpgLocation = %A_ScripDir%PICTURES\01Invoice_No.png
	MsgBox, %jpgLocation%
	
	PicSearch()
	
	MouseMove, %pX%, %pY%
	pX += 10
	pY += 10
	
	sleep 1000
	MouseMove, %pX%, %pY%
	MsgBox, Return2
	
	
	
	
	


	PicSearch(){
		Loop, 5{
			ImageSearch, pX, pY, 0, 0, A_ScreenWidth, A_ScreenHeight, %jpgLocation%  ;%A_ScripDir%PICTURES\00New_Clear.png
			
			if(ErrorLevel = 0){
				MsgBox, Found it
				return ;%pX%, %pY%
			}
			else if(ErrorLevel = 1){
				MsgBox, Couldn't find it!
				;break
			}
			else if(ErrorLevel = 2){
				MsgBox, Unexpected Error Occur!
				;break
			}
		
		}
		MsgBox, Out!
	}
	
	








GuiClose:
ExitApp
  
  
  
F4::
 Exitapp
 return

