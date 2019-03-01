#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


/*
	#Persistent
	#SingleInstance, Force
	W:=A_ScreenWidth-200, H:=A_ScreenHeight-200
	Gui, Add, ActiveX, w%W% h%H% vNav, Shell.Explorer
	Gui, Show, , "urin79.com" Ahk_L Test
	;Gui, hide 
	Nav.Navigate("http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR1BE058626B")
	Nav.Silent := True

	Sleep, 2000 
	;WinMaximize, A
	test := Nav.document.documentElement.InnerHTML
	Sleep, 2000

	msgbox, % test
	Sleep, 2000 

	Return
*/	
	
	



/*
UserName = customer3
Password = Jo123456789
URL = http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR1BE058626B

WB := ComObjCreate("InternetExplorer.Application")
WB.Visible := True
WB.Navigate(URL)
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10

wb.document.getElementById("username").value := UserName
wb.document.getElementById("password").value := Password
msgbox, lets log in after you click OK to this message box
wb.document.all.Login.click()
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10
SplashTextOn,200,200, text, Now logged in and loaded!
sleep 2000
SplashTextOff
return
*/



F10::
	run, http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR2B4FC1B2C

	Urldownloadtofile,http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR2B4FC1B2C, %A_scriptDir%\1.html
	FileRead, CurrentWebPageSource, 1.html
	MsgBox, CurrentWebPageSource


	
	
	
	

GuiClose:
ExitApp


Esc::
ExitApp