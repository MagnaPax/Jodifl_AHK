#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



#Include %A_ScriptDir%\lib\




#Include FindTextFunctionONLY.ahk
#Include FG.ahk
#Include N41.ahk
#Include CommN41.ahk
#Include LAS.ahk
#Include ChromeGet.ahk
#Include MethodsForPT.ahk ; PT 프로세싱을 위한 모든 클래스, 메소드를 모아놓은 파일


#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk



L_driver := new LAMBS
N_driver := new N41
F_driver := New FG
LAS_driver := New LA
CommN41_driver := New CommN41
JO_driver := New JODIFL



		; 사용자의 마우스 이동 막음
		BlockInput, MouseMove

			
		; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		IfWinNotExist, ahk_class XLMAIN
		{
			Loop{
				MsgBox, 262144, No Excel file Warning, PLEASE OPEN BO LIST EXCEL FILE
					
				IfWinExist, ahk_class XLMAIN
					break
			}
		}

			
			
		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible


		; 엑셀에 값이 들어간 만큼(i 값 만큼) 루프 돌면서 엑셀에서 값 읽기
		Loop{

			; SO# 값은 B Column 두 번째에 에 있음
			CustCode := Xl.Range("A2").Value
			
			
			
			; N41 활성화 시키기
			WinActivate, ahk_class FNWND3126

			Sleep 150
			
			; 혹시 이전에 열린 창이 있으면 닫기
			IfWinExist, Customer P.O or S.O #
				WinClose

			Sleep 150


			; 왼쪽 메뉴바에 있는 Customer 클릭하기
			N_driver.ClickCustomerMarkOnTheLeftBar()			
			
			
			; 왼쪽 메뉴바에 있는 SO Manager 클릭하기
			N_driver.ClickSOManagerOnTheLeftBar()
			
			
			; SO Manager 에 있는 Customer 표시 찾기. 고객 코드 입력하기 위함
			N_driver.FindCustomerMarkToFillInTheBlank()
			
			
			; 검색창에 고객 코드 입력 후 엔터쳐서 검색하기		
			Sleep 300
			Send, ^a
			Sleep 150
			Send, % CustCode
			Sleep 150
			SendInput, {Enter}
			Sleep 150
			
			
			
			Sleep 3000
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}			
			while (A_cursor = "Wait")
				Sleep 1000



						; Open SO 에 아이템이 있는지 확인하기
						isThereItemsOnOpenSo := N_driver.checkOpenSoIfThereAreItemsShipOut()
						
						; 아이템이 있다면 정지
						if(!isThereItemsOnOpenSo){
;							MsgBox, 262144, title, 처리하기. 아이템 있음

					; 사용자의 마우스 이동 허용
					BlockInput, MouseMoveOff

							
							; Yes 는 pre-authorized 받고 인쇄
							; No 는 CBS등 돈 받지 않고 인쇄
							; Cancel 는 인쇄하지 말고 다음으로 넘어가기
							SoundPlay, %A_WinDir%\Media\Ring06.wav
							MsgBox, 266243, OPTIONS, Yes : PRINT P.T. WITH pre-authorize BY N41`n`nNo : PRINT P.T. WITH no pre-authorize BY N41`n`nCancel : NO PRINTING`n`n%CustMemoFromJODIFLcom%				
							
							N41_ProcessingForPT_driver := New N41_ProcessingForPT
							
							; pre-authorized 받고 프린트하기
							IfMsgBox, Yes
							{
								
								N41_ProcessingForPT_driver.OpenSO_Print_WITH_PreAuthorized(CustomerPO, IsItFromNewOrder)
								
								;~ PrintByPayment(IsItFromAllocation)
				;				MsgBox, 262144, Title, WAIT TO CONTINUE NEXT ORDER
							}
							; CBS 등 돈 받지 말고 프린트 하기
							IfMsgBox, No
							{
								N41_ProcessingForPT_driver.OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder)
								
								;~ CBS_Print(IsItFromAllocation)
				;				MsgBox, 262144, Title, WAIT TO CONTINUE NEXT ORDER
							}		

		; 사용자의 마우스 이동 막음
		BlockInput, MouseMove
							
						}
						
						
						; 아이템 없으면 A2 줄 지우고 다음 값 읽기
						Xl.Sheets(1).Range("A2").EntireRow.Delete
						continue
			
		}
		
		
		BasicProcessing(SO#, EndSO#, CustomerPO, IsItFromNewOrder, IsItFromExcelFile, SO#FromJODIFLcom, CustMemoFromJODIFLcom, IsItFromAllocation){
			return
		}
		
		
		
		
		
		
		
		
		
		
		
			;~ ; Allocation 파일을 읽는 호출일 때(IsItFromAllocation 변수에 값이 있을 때) 실행
			;~ ; 엑셀에서 SO# 읽어서 변수에 저장
			;~ if(IsItFromAllocation){
				;~ ; SO# 값은 A Column 첫 번째에 에 있음
				;~ SO# := Xl.Range("A1").Value				
			;~ }
			
			;~ ;소수점 뒷자리 정리
			;~ RegExMatch(SO#, "imU)(\d*)\.", SubPat)
				
			;~ ; 정리된 값 SO# 에 넣기
			;~ SO# := SubPat1			
			
	

			;~ ; 만약 지금 얻은 SO# 값이 이전 SO# 값을 저장하고 있는 previousNumber 값과 같다면 
			;~ ; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
			;~ ; 그 전에 SO# 변수와 previousNumber 변수에 값이 있을때만 비교하기
			;~ if(SO# || previousNumber){
				
				;~ IfEqual, SO#, %previousNumber%
				;~ {
					;~ Xl.Sheets(1).Range("A2").EntireRow.Delete
					;~ continue
				;~ }				
			;~ }

			

;~ ;			MsgBox, % SO#
			
			;~ ; 해당 SO# 검색하고 처리하는 메소드 실행
			;~ ; StartSO# 값의 자리에 SO#값을 넣어서 호출하기
			;~ ; 여기서 호출했다는 것을 표시하기 위해 IsItFromExcelFile 변수 1로 만들어서 호출
			;~ ; 뉴오더에서 호출된 것이 아니니까 IsItFromNewOrder 값은 0으로 만들기
			;~ IsItFromExcelFile = 1
			;~ IsItFromNewOrder = 0
			;~ SO#FromJODIFLcom = 0
			;~ BasicProcessing(SO#, EndSO#, CustomerPO, IsItFromNewOrder, IsItFromExcelFile, SO#FromJODIFLcom, CustMemoFromJODIFLcom, IsItFromAllocation)
			
;~ ;MsgBox, 262144, Title, BasicProcessing 함수에서 돌아왔음 2398

			
			;~ ; 값이 없으면 루프 나오면서 끝내기
			;~ if(!SO#){
				;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
				;~ MsgBox, 262144, END OF THE LIST, END OF THE LIST		
					
				;~ ; 엑셀 종료하기
;~ ;				xl.WorkBooks.Close()                                    ;close file
;~ ;				xl.quit
				
				;~ Reload
	;~ ;			break
			;~ }


			;~ ; 백오더 처리일때 실행
			;~ ; Allocation 파일을 읽는 호출이 아닐때 때(IsItFromAllocation 변수에 값이 없을 때) 실행
			;~ if(!IsItFromAllocation){

				;~ ; 두 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
				;~ Xl.Sheets(1).Range("A2").EntireRow.Delete


				;~ ; 중복되는 값의 비교를 위해 previousNumber 변수에 SO# 값 넣기
				;~ previousNumber := SO#
			;~ }
			

			
			;~ ; Allocation 파일을 읽는 호출일 때(IsItFromAllocation 변수에 값이 있을 때) 실행			
			;~ if(IsItFromAllocation){

				;~ ; 첫 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
				;~ Xl.Sheets(1).Range("A1").EntireRow.Delete
				
				;~ ; previousNumber 변수에 항상 같은 값을 집어넣어서 위의 IfEqual, SO#, %previousNumber% 코드가 실행되지 않게 하기
				;~ previousNumber = 1
			;~ }
			
			;~ ; 크롬 창 닫기
			;~ WinClose, ahk_class Chrome_WidgetWin_1

		;~ }
		
		
		
		Esc::
			ExitApp