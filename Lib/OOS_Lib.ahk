#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.




; 바탕화면에 저장된 엑셀 파일 옮기기
moveTheFileOnDesktopFromItToTheExcelFileFolder(OOSStyle#, OOSColor){

	FileMove, C:\Users\JODIFL4\Desktop\%OOSStyle#%-%OOSColor%.xls*, %A_ScriptDir%\CreatedFiles\Excel Files
	
	Sleep 1000
	
	IfNotExist, %A_ScriptDir%\CreatedFiles\Excel Files\%OOSStyle#%-%OOSColor%.xls*
		MsgBox, 262144, NO FILE ERROR, NO FILE EXIST. THE APPLICATION IS RESTARTING.	
	
	return
}


