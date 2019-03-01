#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



class CommEXCEL{

	; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기 
	CheckingPresenceOfExcelFile(){
		
		IfWinNotExist, ahk_class XLMAIN
		{
			loop{
				MsgBox, 262144, No Excel file Warning, Please Open an Excel File to get info from
				IfWinExist, ahk_class XLMAIN
					break
			}
		}		
		return
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
}