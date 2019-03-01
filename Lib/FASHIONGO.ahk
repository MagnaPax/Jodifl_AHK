#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include function.ahk


/*
////////////////////////////////////////////////////////////////////////
;전체실행 시 에러가 나서 일단 주석 처리해놓았음

;F1::


F_PO = MTR1B08BD00EA



;FASHIONGO(F_PO){

	;FASHIONGO 창이 존재하는지 무한정 확인 후 FASHIONGO 창 활성화
;	windowtitle = All Orders - Google Chrome
;	CheckTheWindowPresentAndActiveIt(windowtitle)
	
	;New & Clear 버튼 클릭
	MouseClick, l, 694, 329, 2,
	SendInput, %F_PO%
	SendInput, {Enter}
	sleep 500

;}
*/


/*
ExitApp
  
  
  
Esc::
 Exitapp

