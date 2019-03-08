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







; 지금 읽은 아이템 정보가 allValuesOfItems 변수에 있는지 비교해서 두 번째 읽는 화면이면 아이템 읽는 동작 중단하고 나가기
isThis2ndReadScreen(Y, itemsInfoThatJustReadFromScreen_Arr){
	
	Style#_onTheScreen := itemsInfoThatJustReadFromScreen_Arr[1]
	Color_onTheScreen := itemsInfoThatJustReadFromScreen_Arr[2]
	OrderQTY_onTheScreen := itemsInfoThatJustReadFromScreen_Arr[3]	

	infoOfTheItemBeingJustRead = <%Y%>\|\s%Style#_onTheScreen%\s\|\s%Color_onTheScreen%\s\|\s%OrderQTY_onTheScreen%
	FoundPos := RegExMatch(allValuesOfItems, infoOfTheItemBeingJustRead)

	if(FoundPos){
;		MsgBox, 262144, 두 번째 읽음, 방금 읽은 아이템 정보들과 정확히 일치하는 값들이 allValuesOfItems변수에 있습니다. 즉 현재 페이지를 두 번째 읽은 것입니다.
		return 0
	}


;	MsgBox, % "Y : " . Y . "`nStyle#_onTheScreen : " . Style#_onTheScreen . "`nColor_onTheScreen : " . Color_onTheScreen . "`nOrderQTY_onTheScreen : " . OrderQTY_onTheScreen . "`n`ninfoOfTheItemBeingJustRead`n" . infoOfTheItemBeingJustRead . "`n`n`nallValuesOfItems`n`n" . allValuesOfItems


	return 1
}






; 현재 화면에서 읽어온 아이템 정보들 allValuesOfItems 변수에 저장. 나중에 이 변수를 이용해 마지막 화면인지 확인하게 됨
saveInfoOfTheItemsJustReadToallValuesOfItemsVariable(Y, itemsInfoThatJustReadFromScreen_Arr){
	
	Style#_onTheScreen := itemsInfoThatJustReadFromScreen_Arr[1]
	Color_onTheScreen := itemsInfoThatJustReadFromScreen_Arr[2]
	OrderQTY_onTheScreen := itemsInfoThatJustReadFromScreen_Arr[3]

	
	; 현재 화면에서 읽어온 아이템 정보들 전역변수 allValuesOfItems 에 저장. 나중에 이 변수를 이용해 마지막 화면인지 확인하게 됨
	allValuesOfItems := allValuesOfItems . "<" . Y . ">| " . Style#_onTheScreen . " | " . Color_onTheScreen . " | " . OrderQTY_onTheScreen . "`n"	
	
	return
}





