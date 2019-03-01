#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

/*

<WEB, TIFFANY가 주문받은 전화주문 >
, Sales #43320/PO #TIFFANY, Sales #43321/PO #MTR171EC8, Sales #44310/PO #MTR171F67, Sales #45015/PO #MTR171FC7 TRACKING#

<패션고만 여러개>
, Sales #37692/PO #MTR1A6E6C51B0, Sales #41447/PO #MTR2B2EAAB11, Sales #42582/PO #MTR1B2D9747A0

<WEB, 맨 마지막에 메모가 있는 경우>
, Sales #44421/PO #MTR171DD8     ps number 0FF511 TRACKING#

<LASHOWROOM 여러개>
, Sales #37836/PO #OP032015034, Sales #39354/PO #OP040516341, Sales #42376/PO #OP051519990, Sales #43341/PO #OP052420816, Sales #43798/PO #OP053121354, Sales #44816/PO #OP061122269 TRACKING#

<패션고, LASHOROOM>
, Sales #42761/PO #MTR1B370DDB96 TRACKING# 1ZW947530366833053, Sales #37836/PO #OP032015034 TRACKING#

*/

	F_arr := []
	L_arr := []
	W_arr := []
	i = 1
	j = 1
	
	PO_Numbers = , Sales #42761/PO #MTR1B370DDB96, Sales #37836/PO #OP032015034, Sales #44421/PO #MTR171DD8     ps number 0FF511 TRACKING#
	MsgBox, % strLen(PO_Numbers)

/*
	arr := [] 
	i = 1
	j = 10

	loop, 3{
		arr[i] := j
	
		MsgBox, % i
		MsgBox, % j
		MsgBox, % arr[i]
	
		i += 1
		j += j
	}
*/



	;MTR로 시작하는 FASHIONGO, 웹 주문 PO찾기
	;인보이스 맨 끝에 '공란TRACKING#' 먼저 입력하고 PO번호 찾아야 됨. 문장 맨 뒤까지 검색하는 방법을 모르겠음
	;array에 SubPat2값 입력하고 A_index로 마지막 PO번호 가져올까
	FoundPos = 1
	while(FoundPos := RegExMatch(PO_Numbers, "U)(Sales #\d.*/PO\s#(MTR.*)[\,|\s])", SubPat, FoundPos + strLen(SubPat)))
	{
		;패션고 여는 함수 호출(SubPat2값 넘겨줘야 됨)
		if(strLen(SubPat2) >= 10){
			MsgBox, It's a FASHIONGO
			MsgBox, % strLen(SubPat2)
			F_arr[i] := SubPat2
		}
		else{
				MsgBox, It's a WEB
				W_arr[i] := SubPat2   ;MsgBox, It's Web ;웹 처리하는 함수 호출(SubPat2값 넘겨줘야 됨)
			}
		
			
	MsgBox, FoundPos = %FoundPos%`nSubPat = %SubPat%`nSubPat1 = %SubPat1%`nSubPat2 = %SubPat2%
	MsgBox, It's MTR
	
	i += 1
	j += j
	
	}
*/	MsgBox, % F_arr[1]
	MsgBox, % W_arr[1]


	;OP로 시작하는 LASHOROOM 주문 PO찾기
	;인보이스 맨 끝에 '공란TRACKING#' 먼저 입력하고 PO번호 찾아야 됨. 문장 맨 뒤까지 검색하는 방법을 모르겠음
	FoundPos = 1
	while(FoundPos := RegExMatch(PO_Numbers, "U)(Sales #\d.*/PO\s#(OP.*)[\,|\s])", SubPat, FoundPos + strLen(SubPat))) 
	{ 
	MsgBox, FoundPos = %FoundPos%`nSubPat = %SubPat%`nSubPat1 = %SubPat1%`nSubPat2 = %SubPat2%
	MsgBox, It's OP
	;LASHOROOM 처리하는 함수 호출(SubPat2값 넘겨줘야 됨)
	}
*/	


	ExitApp
	
Esc::
ExitApp