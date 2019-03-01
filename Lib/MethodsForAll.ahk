#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


class MethodsForAll{


	; 만료일 4자리 숫자로 만들기
	Refine_ExpDate(ExpDate){
		
		;~ ExpDate = 8/2022
		;~ ExpDate = 10/2022
		;~ ExpDate = 04 / 2020
		;~ ExpDate = 12 / 2020


		StringReplace, ExpDate, ExpDate, %A_SPACE%, , All ; 모든 스페이스 제거
		StringReplace, ExpDate, ExpDate, /, , All ; 모든 / 제거
		ExpDate := Trim(ExpDate)
		if(StrLen(ExpDate) == 5)
			ExpDate := "0"ExpDate
		
		StringLeft, leftOf, ExpDate, 2
		StringRight, RightOf, ExpDate, 2
		
		ExpDate := % leftOf . RightOf
		
		
;		MsgBox, % ExpDate
		
		
		return ExpDate
	}
	
	
	

	; Arr_Original 배열을 받아서 UnquotedOutputVar조건에 맞는 문자열 돌려주기
	FindAdd2_In_Add1(Arr_Original, UnquotedOutputVar){
;		MsgBox, % "It's in FindAdd2_In_Add1 method`nArr_Original is : " . Arr_Original
		;~ while(RegExMatch(Arr_Original, "im)\s((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)", FoundAdd2)){
		while(RegExMatch(Arr_Original, UnquotedOutputVar, FoundAdd2)){
			if(ErrorLevel = 0){
				Temp := FoundAdd2
				Temp := Trim(Temp)
;				MsgBox, % "Add2 : " . Temp . "|"
			}
		break
		}
		return Temp
	}




	; 배열 받아서 UnquotedOutputVar 조건에 맞는 문자열 삭제하기
	DeleteAdd2_In_Add1(Arr_Original, UnquotedOutputVar){		
		;~ while(FoundPos := RegExMatch(Arr_Original, "im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)", FoundAdd2)){
		while(FoundPos := RegExMatch(Arr_Original, UnquotedOutputVar, FoundAdd2)){
			if(ErrorLevel = 0){
				Temp := SubStr(Arr_Original, 1, FoundPos-1)
				Temp := Trim(Temp)
;				MsgBox, % Temp . "|"

				return Temp
			}
		break
		}
		return Arr_Original ; UnquotedOutputVar 조건 맞는 문자열 못 찾으면 그냥 Arr_Original 리턴하기
	}


	
	
	
	
}