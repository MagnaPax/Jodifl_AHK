/*
1. Style Master-List 에서 아래의 필터 조건으로 만든 엑셀 파일을 읽어서(이 프로그램 실행 전 수동으로 만들어야 됨)
( orders >= 6 ) And ( picks >= 6 ) And ( orders > picks ) And ( ohs > picks )

2. Style Master-Style 연 뒤 스타일 번호와 색깔이 맞는 것을 찾음

3. 찾은 스타일의 On Order 열어서 엑셀 파일로 저장한 뒤 바탕화면의 pending_styles_not_shipped 폴더로 옮김
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include %A_ScriptDir%\lib\


#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk


#Include CommN41.ahk
#Include OOS_Lib.ahk
#Include Xl_Lib.ahk
;~ #Include N41_Lib.ahk
#Include <N41_Lib>



MsgBox, 262144, Title, Click Ok to continue


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
	BlockInput, Mouse


	; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
	IfWinNotExist, ahk_class XLMAIN
	{
		loop{
			MsgBox, 262144, No Excel file Warning, PLEASE OPEN AN ORDER EXCEL FILE
			IfWinExist, ahk_class XLMAIN
				break
		}
	}	



	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible	


	;엑셀 값의 끝 row 번호 알아낸 후 lastLine#OfRow 에 값 넣기
	XL_Handle(XL,1) ;get handle to Excel Application
	lastLine#OfRow := XL_Last_Row(XL)
	lastLine#OfRow--	; 맨 마지막에 토탈값(쓰레기값)이 들어있기 때문에 1 줄여주기


;MsgBox, 262144, Title, 끝번호 : %lastLine#OfRow%


	; 사용자의 마우스 이동 막음
	BlockInput, MouseMove


	; 아이템 찾는 동안 보여줄 프로그래스 바
	TotalLoops := % lastLine#OfRow
	Gui, -Caption +AlwaysOnTop +LastFound
	Gui, Add, Text, x12 y9 w170 h20 , P  R  O  C  E  S  S  I  N  G  .  .  .
	Gui, Add, Progress, w410 Range0-%TotalLoops% vProgress
	Gui, Show, w437 h84, SEARCHING ITEMS
	





	; 스타일 번호와 색깔이 들어있는 Row 번호
	#ofLineOfStyle#AndColor = 3
	Loop{
		
		; 스타일 번호 들어있는 셀 위치
		cell#OfStyle#ToBeCopied := "A" . #ofLineOfStyle#AndColor
		
		; 그 스타일 색깔이 들어있는 셀 위치
		cell#OfColorOfTheStyle#ToBeCopied := "B" . #ofLineOfStyle#AndColor
		
		
		; Style 번호와 색깔 읽어와서 각각 변수에 저장
		Style# := Xl.Range(cell#OfStyle#ToBeCopied).Value
		Color := Xl.Range(cell#OfColorOfTheStyle#ToBeCopied).Value		
		
;MsgBox, 262144, Title, %Style#%`n%Color%
		
		
		
		; N41 활성화하기
		WinActivate, ahk_class FNWND3126
	
		; 객체 선언
		CommN41_driver := new CommN41
		
		Loop
		{
			; Style 클릭하기
			CommN41_driver.click_Style()
			
			; Style 탭 안에 있는 Style 클릭하기
			CommN41_driver.click_StyleButtonInTheStyleTab()
			
			; 스타일 입력칸 찾아서 클릭하기
			CommN41_driver.findTheSearchBoxOfStyle()			
			
	
			; 입력칸에 이미 있던 값 지우기
			Loop, 10{
				Send, {Del}
				Sleep 50
			}
		
			; 스타일 번호 입력
			Send, %Style#%
			Send, {Enter}
		;	Sleep 1000
			
			
			; 색깔 입력칸 찾아서 클릭하기
			CommN41_driver.findTheSearchBoxOfColor()
			
			; 색깔 입력
			;~ SetKeyDelay 50,200
			SetKeyDelay 10,50
			SendEvent %Color%
			Sleep 3000			
			;~ Send, %Color%	
			;~ Send, {Enter}
			
			
			; 화면에 제대로 들어왔는 확인키 위해 화면에서 스타일 번호 읽기
			Style#_OnTheScreen := CommN41_driver.getStyleNumberOnStyleTab()
			
			; 화면에 제대로 들어왔는 확인키 위해 화면에서 색깔 읽기
			Color_OnTheScreen := CommN41_driver.getColorOnStyleTab()
			
			
;	MsgBox, % Style#_OnTheScreen . "`n`n" . Color_OnTheScreen
			
			; 화면에서 읽은 값들과 입력한 값들이 같으면 루프 빠져나가기
			if(Style# == Style#_OnTheScreen && Color == Color_OnTheScreen)
				break
		}
		
		
		; On Order 의 돋보기 버튼 눌러서 창 열기
		CommN41_driver.clickMagnifierNextTo_OnOrder()
		
		
		; 적당한 창 위치에서 (이 경우에는 Close 버튼을 찾은 뒤 그 위에 마우스를 위치했음) 마우스 오른쪽 버튼 클릭 후 Excel 파일 만들기 메뉴에서 엔터치기
		CommN41_driver.activeDownloadExcelFileFunction()
		
		
		; 엑셀파일을 바탕화면에 저장하기
		CommN41_driver.makeTheExcelFileOfTheStyleOnDesktopScreen(Style#, Color)
		
		
		
		
		;;; 바탕화면에 저장된 파일 특정 폴더로 옮기기
		
		; 바탕화면에 폴더 없으면 만든 뒤
		IfNotExist, %A_Desktop%\pending_styles_not_shipped
			FileCreateDir, %A_Desktop%\pending_styles_not_shipped
		
		; 방금 전 만든 파일 그 폴더로 옮기기		
		StringReplace, Color, Color, /, , All	; 색깔 중간에 '/' 있으면 지우기
		
		; 폴더 안에 방금 만든 파일명과 같은 파일이 있으면 지우기
		IfExist, %A_Desktop%\pending_styles_not_shipped\%Style#%-%Color%.xls*
			FileDelete, %A_Desktop%\pending_styles_not_shipped\%Style#%-%Color%.xls*
		
		; 방금 전 만들어서 바탕화면에 저장 한 파일 플더로 옮기기
		FileMove, %A_Desktop%\%Style#%-%Color%.xls*, %A_Desktop%\pending_styles_not_shipped
		
		
		
		
		; Path 에 위치한 엑셀창 열기. 바탕화면에 있는 파일 열기
		Path = %A_Desktop%\pending_styles_not_shipped\%Style#%-%Color%.xlsx
		Xl_theStyle := openTheExcelFile(Path)



		; 방금 연 파일의 마지막 줄 번호 얻기
		#ofLastCell := XL_Last_Row(Xl_theStyle)
;MsgBox, % #ofLastCell

		; 쓸 데 없는 마지막 2줄 삭제하기
		Loop, 1{
			Xl_theStyle.Sheets(1).Range("A" . #ofLastCell).EntireRow.Delete
			#ofLastCell--	
		}
		

		; 쓸 데 없는 첫 3줄 지우기
		Loop, 3
			Xl_theStyle.Sheets(1).Range("A1").EntireRow.Delete
		
		; 저장한 뒤 파일 닫기
		Xl_theStyle.ActiveWorkbook.save()
		Xl_theStyle.Quit


		; 프로그래스 바 1씩 증가
		GuiControl,,Progress, +1


		; 다음 줄 읽기 위해 1 증가시키기
		#ofLineOfStyle#AndColor++
		
		; 마지막에 도착하면 루프 끝내고 나가기
		if(#ofLineOfStyle#AndColor > lastLine#OfRow)
			break				
	}
	
	

	; 프로그래스 바 닫기
	Gui Destroy
	
	; 사용자의 마우스 이동 허용
	BlockInput, MouseMoveOff	
	
	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 262144, REACH THE END, IT'S END
	
	ExitApp


	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
Esc::
ExitApp
