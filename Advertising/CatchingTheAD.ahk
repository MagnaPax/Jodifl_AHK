﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


F1::

WinActivate, UPS WorldShip - Remote Workstation

;To get Activated window title, width, height, and x and y point of the window
WinGetActiveStats, Title, Width, Height, X, Y
;MsgBox, "%Title%" 활성 창은 너비가 %Width%이고`, 높이가 %Height%`, 그리고 %X%`,%Y%에 위치합니다.
;MouseClick, l, 300, 800, 3
Send, {F2}

X_1 := X + 100

return





; 이 핫키로 모든 열린 브라우저 창을 예전부터 새로운 순서로 방문할 수 있습니다:
F2::

SetTitleMatchMode, 2
WinActivateBottom, - Google Chrome
return


^!z::  ; Control+Alt+Z 핫키.
MouseGetPos, MouseX, MouseY
PixelGetColor, color, %MouseX%, %MouseY%
MsgBox 현재 커서 위치의 컬러는 %color%입니다.
return


F3::

/*
	while(PixelSearch, , , 655, 650, 660, 655, 0x228B22, , Fast){
		MouseClick, l, 1258, 331
		MsgBox, Foun it!
	}



*/
/*
	while(PixelSearch, , , 50, 60, 200, 300, 0x228B22, , Fast){
		MouseClickDrag, l, 660, 350, 770, 460
		;MouseClick, l, 1258, 331
		MsgBox, Foun it!
	}
	0xBBBBBB
*/	
	loop{
		;MouseClick, l, 1254, 319, 1 ;광고 잡아야 되는 달의 하이퍼 링크를 계속 클릭
		MouseClick, l, 1335, 389, 1 ;광고 잡아야 되는 달의 하이퍼 링크를 계속 클릭
		;sleep 1800
		;MsgBox, 1
		PixelSearch, , , 1078, 480, 1100, 590, 0x228B22, , fast
			if ErrorLevel{	;픽셀 서치에서 색깔을 발견 못하면 이곳을 계속 반복
			;MsgBox, 해당 컬러를 지정된 구역에서 발견하지 못했습니다.`n%ErrorLevel%
			;MouseClickDrag, l, 660, 350, 770, 460
			MouseClick, l, 1254, 319, 1 ;광고 잡아야 되는 달의 하이퍼 링크를 계속 클릭
			sleep 1830
			}
			else{	;색깔 발견 했으면 이곳을 실행
			;MsgBox, 3 쉐이드의 변형 안에서 컬러가 X%Px% Y%Py%에서 발견되었습니다.
			MouseClickDrag, l, 650, 331, 1330, 958 ;달력의 왼쪽위와 오른쪽 아래를 드래그 하고
			MouseClick, l, 1185, 260, 2 ; Purchase Ad 버튼을 클릭 한 뒤
			break ; 루프를 빠져나옴
			}
	}
	
F4::
	
MouseClick, l, 1185, 260, 2 ; Purchase Ad 버튼을 클릭 한 뒤

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;마우스 드래그를 왼쪽 위부터 아래 오른쪽까지 하지 말고 월화수쪽만 밑에서부터(경쟁율이 낮게) 위로 드래그 해보자
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;







ExitApp
  
Esc::
 Exitapp