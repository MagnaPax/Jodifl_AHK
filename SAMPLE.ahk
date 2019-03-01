#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include %A_ScriptDir%\lib\

#Include ChromeGet.ahk
#Include CommN41.ahk
#Include FindTextFunctionONLY.ahk
#Include N41.ahk





; 예를 들어, 다음은 왼쪽 10번째 문자부터 검색을 시작합니다
;~ StringGetPos, OutputVar, InputVar, abc, , 9

;~ locOf_Style# = 3
;~ Color_onTheScreen = 14

allValuesOfItems = | a1234 | Gold Yellow Mix | 6
Style#_onTheScreen = a1234
Color_onTheScreen = Gold Yellow Mix
OrderQTY_onTheScreen = 6


	; allValuesOfItems 값 중 지금 읽은 스타일 번호 값이 들어있다면
	IfInString, allValuesOfItems, %Style#_onTheScreen%
	{
		StringGetPos, locOf_Style#, allValuesOfItems, %Style#_onTheScreen%
		StringGetPos, locOf_Color, allValuesOfItems, %Color_onTheScreen%, , locOf_Style#	; 먼저 찾은 locOf_Style# 위치 다음부터 검색
		StringGetPos, locOf_Qty, allValuesOfItems, %OrderQTY_onTheScreen%, , locOf_Color	; locOf_Qty 위치 다음부터 검색
		
		MsgBox, % "locOf_Style# : " . locOf_Style# . "`nlocOf_Color : " . locOf_Color . "`nlocOf_Qty : " . locOf_Qty
		
		; 일단 StringGetPos 명령을 실행한 결과중 하나라도 거짓이면 자세히 살펴볼 것도 없이 무조건 똑같은 아이템 정보가 이 화면에는 없다는 뜻
		if ErrorLevel   ; 즉, 비어 있지 않거나 0이면.
			MsgBox, 중복값이 없습니다.
		
		; 아이템에 관련된 세 가지 값을 allValuesOfItems 안에서 모두 찾았을 때
		; 하지만 이런 경우에도 2번째 아이템과 아이템 번호가 같고 4번째 아이템과 색깔이 같고 4번째 아이템과 주문량이 같을 수 있다. StringGetPos 함수는 어떤 변수 안에서 찾으려는 값이 있기만 하면 위치를 반환하기 때문
		; 정확히 중복된 아이템이라면 처음 찾은 아이템 번호의 위치를 기준으로 아이템 색깔과 아이템 주문량을 찾은 위치가 아이템 번호를 찾은 위치 바로 옆에 있어야 된다
		else{
			if(locOf_Style#+15 >= locOf_Color && locOf_Color+20 >= locOf_Qty){
				MsgBox, 262144, title, 지금 읽은 아이템 정보와 일치하는 아이템 정보값을 allValuesOfItems 변수에서 찾았습니다. 즉, 이 화면은 두 번째 들어온 것입니다.
			} ; 끝 - if(locOf_Style#+15 >= locOf_Color && locOf_Color+20 >= locOf_Qty){
		} ; 끝 - else		
	} ; 끝 - IfInString, allValuesOfItems, %Style#_onTheScreen%

if(locOf_Style#+15 >= Color_onTheScreen){
	MsgBox, it's in the scape
}

MsgBox





Style#_onTheScreen = A1234
Color_onTheScreen = MAUVE
OrderQTY_onTheScreen = 6
allValuesOfItems := allValuesOfItems . "|" . Style#_onTheScreen . "|" . Color_onTheScreen . "|" . OrderQTY_onTheScreen . "  "

MsgBox, % allValuesOfItems

allValuesOfItems := allValuesOfItems . "|" . Style#_onTheScreen . "|" . Color_onTheScreen . "|" . OrderQTY_onTheScreen . "  "

MsgBox, % allValuesOfItems


ToAdd = 1111
Haystack = abcdefghijklmnopqrs
Haystack := Haystack . ToAdd

MsgBox, % Haystack
Needle = def

IfInString, Haystack, %Needle%
	MsgBox, % INSTR(Haystack, Needle)
else
	MsgBox, 발견 안됨




nameArray := ["Chris","Joe","Marcy","Chris","Elina","Timothy","Joe"]

;~ trimmedArray := trimArray(nameArray)

for i, nameArray in names
    for j, inner_namearray in names
        if (A_Index > i && nameArray = inner_namearray)
            names.Remove(A_Index)

Loop % nameArray.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . nameArray[A_Index]
}

trimArray(arr) { ; Hash O(n) 

    hash := {}, newArr := []

    for e, v in arr
        if (!hash[v])
            hash[(v)] := 1, newArr.push(v)

    return newArr
}

MsgBox





;~ /* 배열로부터 읽기 첫 번째 방법
Array:=[1,3,"ㅋㅋ"]
Array.push("aa")

;~ IfInString, Haystack, %Needle%

Haystack = abcdefghijklmnopqrs
Needle := "ㅋㅋ"
IfInString, Array, %Needle%
{
    MsgBox, 문자열이 발견되었습니다.
}

MsgBox, % InStr(Array, "aa")

Loop % Array.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
}
*/

MsgBox




data1=
(
1572 6/17
2697 6/17
2728 6/17
2521 6/17
4545 6/17
)

data2=
(
1572 6/10
2521 6/18
2697 6/17
9922 6/10
2728 6/17
)

for key, var in ["data1", "data2"] {
	o%var% := {}
	Loop, Parse, %var%, `n
		SpacePos := InStr(A_LoopField, A_Space)
		, o%var%["_" SubStr(A_LoopField, 1, SpacePos - 1)] := SubStr(A_LoopField, SpacePos + 1)
}

for key, val in odata1
	if (odata2[key] = val)
		No_Change .= SubStr(key, 2) " " odata2.Remove(key) "`n"
	else if (odata2[key] != "")
		Date_changed .= SubStr(key, 2) " " odata2.Remove(key) "`n"
	else
		Removed .= SubStr(key, 2) " " val "`n"

for key, val in odata2
	New_Campaigns .= SubStr(key, 2) " " val "`n"

MsgBox, % "No Change:`n" No_Change "`n`nNew Campaigns:`n" New_Campaigns "`n`nDate changed:`n" Date_changed "`n`nRemoved:`n" Removed








q:: ;array - remove duplicates (case insensitive)
oArray := ["a","B","c","A","B","C",1,1.0,"1","1.0"]
oArray2 := [], oTemp := {}
for vKey, vValue in oArray
{
	if (ObjGetCapacity([vValue], 1) = "") ;is numeric
	{
		if !ObjHasKey(oTemp, vValue+0)
			oArray2.Push(vValue+0), oTemp[vValue+0] := ""
	}
	else
	{
		if !ObjHasKey(oTemp, "" vValue)
			oArray2.Push("" vValue), oTemp["" vValue] := ""
	}
}
vOutput := ""
for vKey, vValue in oArray2
	vOutput .= vKey " " vValue "`r`n"
MsgBox, % vOutput
return

w:: ;array - remove duplicates (case sensitive)
oArray := ["a","B","c","A","B","C",1,1.0,"1","1.0"]
oArray2 := [], oTemp := ComObjCreate("Scripting.Dictionary")
for vKey, vValue in oArray
	if !oTemp.Exists(vValue)
		oArray2.Push(vValue), oTemp.Item(vValue) := ""
vOutput := ""
for vKey, vValue in oArray2
	vOutput .= vKey " " vValue "`r`n"
MsgBox, % vOutput
return




List := ["orange", "apple", "orange", "banana", "apple"]

Msgbox % "Before:`n" . ToString(List)

for i, fruit in List
	for j, inner_fruit in list
		if (A_Index > i && fruit = inner_fruit)
			List.Remove(A_Index)

Msgbox % "After:`n" . ToString(List)
		
ToString(array) {
	for i, key in array
		outstr .= key . "`n"
	return outstr
}


;~ Data := "Hi Everyone"
;~ MyArray := StrSplit(Data)
;~ msgbox, % MyArray.1
;~ OtherArray := MyArray.Clone()
;~ msgbox, % OtherArray.1



;~ MyTable := []

;~ MyTable[1] := "d"
;~ MyTable[2] := "o"

;~ Array:=[1,3,"ㅋㅋ"]
;~ OtherArray := Array.Clone()
;~ Loop % Array.Maxindex(){
	;~ MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
;~ }

;~ Loop % OtherArray.Maxindex(){
	;~ MsgBox % "Element number " . A_Index . " is " . OtherArray[A_Index]
;~ }

;~ MsgBox


; 다중배열에 있는 값 읽기
MyTable := []

MyTable[1, 1] := "d"
MyTable[1, 2] := "o"
MyTable[1, 3] := "g"
MyTable[2, 1] := "e"

OtherArray := MyTable.Clone()

if(MyTable == OtherArray){
	MsgBox, same
}

MsgBox % MyTable.1.MaxIndex() "`n" MyTable.2.MaxIndex() ; 배열에 총 몇개의 값이 있는지 확인

For 1st, Row in MyTable
	For 2nd, element in Row
		MsgBox, % "1st : " . 1st . "`n`n2nd : " . 2nd . "`n`n" . "Array" . "[" . 1st . "][" . 2nd . "]" . "`n`nelement : " . element



For 1st, Row in OtherArray
	For 2nd, element in Row
		MsgBox, % "1st : " . 1st . "`n`n2nd : " . 2nd . "`n`n" . "Array" . "[" . 1st . "][" . 2nd . "]" . "`n`nelement : " . element

MsgBox, % MyTable.1.HasKey("d") ;1
MsgBox, % MyTable.HasKey("o") ;1
MsgBox, % MyTable.HasKey("a") ;0


;~ q:: ;arrays and HasKey
;~ ;if an array has a key called HasKey,
;~ ;HasKey() will fail, so use ObjHasKey instead
;~ oArray := {hello:0}
;~ MsgBox, % oArray.HasKey("hello") ;1
;~ oArray := {Haskey:0}
;~ MsgBox, % oArray.HasKey("HasKey") ;(blank)
;~ MsgBox, % oArray.HasKey("abc") ;(blank)
;~ MsgBox, % ObjHasKey(oArray, "HasKey") ;1
;~ MsgBox, % ObjHasKey(oArray, "abc") ;0
;~ return





oArray := {}
oArray[1] := "a"
oArray[1.0] := "b"
oArray["1"] := "c"
oArray["1.0"] := "d"
vOutput := ""
for vKey, vValue in oArray
	vOutput .= vKey " " vValue "`r`n"
MsgBox, % vOutput
;MsgBox(vOutput)
return

;AHK v1
;1 a
;1 c
;1.0 d

;AHK v2
;1 c
;1.0 d












HouseMemoOnSOTab = UPDATED PO#

if(HouseMemoOnSOTab == "UPDATED PO#"){
	MsgBox, it is
}

MsgBox

; 다중배열에 있는 값 읽기
MyTable := []

MyTable[1, 1] := "d"
MyTable[1, 2] := "o"
MyTable[2, 1] := "e"

;~ MsgBox % ArrA.1.MaxIndex() "`n" ArrA.2.MaxIndex()
MsgBox % MyTable.1.MaxIndex() "`n" MyTable.2.MaxIndex() ; 배열에 총 몇개의 값이 있는지 확인
if(XY_Arr.1.MaxIndex() <= 7){
	MsgBox, 262144, title, 끝에 도달했음. 루프 빠져나감
}

For 1st, Row in MyTable
	For 2nd, element in Row
		MsgBox, % "1st : " . 1st . "`n`n2nd : " . 2nd . "`n`n" . "Array" . "[" . 1st . "][" . 2nd . "]" . "`n`nelement : " . element		
	
	
	

MsgBox, % Clipboard
Clipboard := ""
MsgBox, % Clipboard


X = 1
Y = 2

		preX := X
		preY := Y
		
		MsgBox, % preX . preY


aa = 11
bb := "22"
cc = 33

a = 11
b = 22
c = 33

;~ if(!aa||!bb){
if(aa == a && bb == b && cc == c){
	
	MsgBox, no
}

MsgBox







MsgBox


#include <Vis2>  ; Equivalent to #include .\lib\Vis2.ahk

Vis2.Graphics.Subtitle.Render("Running Test Code... Please wait", "t7000 xCenter y67% p1.35% c88EAB6 r8", "s2.23% cBlack")
Vis2.Graphics.Subtitle.Render("Press [Win] + [c] to highlight and copy anything on-screen.", "time: 30000 xCenter y92% p1.35% cFFB1AC r8", "c000000 s2.23%")
MsgBox % text := OCR("test.jpg")

#c:: OCR()              ; OCR to clipboard
#i:: ImageIdentify()    ; Label images
;~ Esc:: ExitApp





MsgBox











MyTable := []



MyTable[1, 1] := "d"

MyTable[1, 2] := "o"

MyTable[1, 3] := "g"

MyTable[2, 1] := "e"



For Each, Row in MyTable

	For, Each, item in Row

		MsgBox %item%
MsgBox, pause







SO#_rowLetter = B

MsgBox % SO#_rowLetter


DropStyle# = P2392
StyleColor = BURGUNDY

innerTextOfCheckedOne = P2392-EGGPLANT
innerTextOfCheckedOne = P2392-BURGUNDY

dropStyle_Color = %DropStyle#%-%StyleColor%

MsgBox, % dropStyle_Color

if dropStyle_Color in %innerTextOfCheckedOne%
{
	MsgBox, it's mached
}

MsgBox

aaa = 11
bbb = 22

ccc = 33
ccc = 11
ddd = 44
ddd = 22
Loop
{
	MsgBox, not matched
	if(aaa == ccc && bbb == ddd){
		break
	}
}
	
	MsgBox






; 배열 선언
Arr_ShippingAdd := object()
Arr_CC := object()
	
; 배열 초기화
Arr_ShippingAdd := []
Arr_CC := []


Arr_CC.Insert("han")


MsgBox



N41_Driver := New N41

N41_Driver.00000000000000NEW__UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)



MsgBox







	
	Arr_OrdinalNumbers := abc(Arr_OrdinalNumbers)
	
	
;~ /* 배열로부터 읽기 첫 번째 방법
Loop % Array.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
}
*/	

driver := Arr_OrdinalNumbers[1]

MsgBox, % "driver : " . driver
	
	abc(Arr_OrdinalNumbers){
		
		Arr_OrdinalNumbers.Insert("29")
		
		aaa = 1
		bbb = 2
		
		return [aaa, bbb]
	}
	
	
MsgBox
	
	Arr_OrdinalNumbers.Insert("29")
	Arr_OrdinalNumbers.Insert("30")
	Arr_OrdinalNumbers.Insert("31")
	Arr_OrdinalNumbers.Insert("32")


find# := Arr_OrdinalNumbers[1] - 2

MsgBox, % find#

MsgBox, % Arr_OrdinalNumbers.Maxindex() / 2

; DropStyle# 에 해당되는 Xpath 갯수가 2개씩 찾아지기 때문에 나누기 2 해준 횟수만큼만 루프 돌면서 해당되는 곳의 Status 값이 Disabled 인지 Enabled 인지 얻기
ordinal#ofTheArray = 1  ; 몇 번째 배열 값을 볼 지 정하기 위해
Loop % Arr_OrdinalNumbers.Maxindex()/2+2{
;	find# := Arr_OrdinalNumbers[ordinal#ofTheArray]  ; ordinal#ofTheArray 에 있는 숫자의 위치에 있는 Arr_OrdinalNumbers 배열의 값을 find# 에 넣기 예를 들어 ordinal#ofTheArray 값이 1이면 Arr_OrdinalNumbers[1] 에 들어있는 값을 find# 에 넣기
	;~ MsgBox % "Element number " . A_Index . " is " . Arr_OrdinalNumbers[ordinal#ofTheArray]
	;~ MsgBox % "Element number " . A_Index . " is " . find#
;	ordinal#ofTheArray := ordinal#ofTheArray + 2 ; 1,3,5,7 이런식으로 보기위해. 왜냐면 홀수번째와 짝수번째의 값이 똑같기 때문에 두 번 확인할 필요 없으니까

	MsgBox, % a_index
}

MsgBox

Loop % Arr_OrdinalNumbers.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . Arr_OrdinalNumbers[A_Index]
	howManyDropStyleXpathsOnThisPage++
}

MsgBox
	
	
	

	Loop{
		
		loop# := A_Index
		
		MsgBox, % loop#
	}

DropStyle# = P2471-1
StyleColor = MOCHA

; 현재 Xpath 에서 읽어온 값
valueOnThePage = P2471-5-SNAKE

DSvalue = %DropStyle#%-%StyleColor%

MsgBox, % "DSvalue : " . DSvalue . "`n`nvalueOnThePage : " . valueOnThePage


; 글자 갯수만큼 #ofLetter 에 넣기. 이 숫자만큼만 루프 돌기 위해
#ofLetter := StrLen(valueOnThePage)

; #ofLetter 에 담긴 valueOnThePage 의 글자 갯수만큼만 루프 돌기
Loop, % #ofLetter{
	MsgBox, % A_Index

	; valueOnThePage 변수에 담긴 글자를 오른쪽부터 루프 회차만큼 읽어서 OutputVar 변수에 넣기
	StringRight, OutputVar, valueOnThePage, % A_Index
	MsgBox, % OutputVar
	
	; valueOnThePage 변수를 오른쪽부터 읽은 뒤 저장한 OutputVar 변수에 -가 발견되면 그것은 색깔을 구분하기 위한 - 가 발견됐다는 뜻
	; - 이후의 값 즉 색깔 값을 지운뒤 앞의 스타일 번호 값만 sTyle#WithOutColor_OnThisPage 변수에 넣기
	if OutputVar contains -
	{
		MsgBox, % "- 발견됨."
		StringTrimRight, sTyle#WithOutColor_OnThisPage, valueOnThePage, % A_Index
		break
	}
	#ofLetter--
}

MsgBox, % "sTyle#WithOutColor_OnThisPage : " . sTyle#WithOutColor_OnThisPage

; DropStyle# 와 sTyle#WithOutColor_OnThisPage 가 같지 않으면 이것은 화면에 검색되어 나타나기는 했지만 필요 없는 스타일번호 값이라는 뜻
if(DropStyle# != sTyle#WithOutColor_OnThisPage){
	MsgBox, 같지 않음
}

; 이렇게 같을때만 동작시키기
else if(DropStyle# == sTyle#WithOutColor_OnThisPage){
}

MsgBox, pause







MsgBox, 262144, Title, JOB FINISHED. THE STYLE '%DropStyle#% %StyleColor%' HAS BEEN DROPPED.

	itHasToBeActivated = 1
	DropAll = 1
	
	
	MsgBox, % "DropAll : " . DropAll
	
	
	if(itHasToBeActivated == 0 || DropAll){
		MsgBox, 모든 Status 값이 Disable 이므로 이 아이템 전체가 Disabled 되어야 됨.
	}




	
	
	MsgBox
		



















/*
; 중첩루프 빠져나오기
; 내포 회돌이 안으로부터 바깥 회돌이로 빠져 나옵니다.
outer:
Loop 3
{
    x := A_Index
    Loop 3
    {
        if (x*A_Index = 6)
            break outer  ; break 2 또는 goto break_outer와 동등합니다.
        MsgBox %x%,%A_Index%
    }
}
break_outer: ; goto를 위한 라벨.
*/





/*
; 메소드에서 리턴 값 배열로 받기

; 배열 선언
Array := object()
	
; 배열 초기화
Array := []


driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver

Array := aabbcc(driver)

MsgBox, % "메소드에서 넘어온 value 값은 : " . Array[2]

Array[1].Get("http://duckduckgo.com/")

MsgBox, pause




aabbcc(driver){	
	driver.Get("https://google.com")	
	value = 111
	
	return [driver, value]
}
*/



/*
; 홀수 짝수 알아내기
Number:=5
If (Number & 1)  ; if last bit is 1
   MsgBox Odd    ; then it's odd
Else MsgBox even ; else it's even
*/	





/*
	; 마우스 커서의 현재 위치를 실시간으로 열람
;	MouseGetPos, xpos, ypos
;	Msgbox, The cursor is at X%xpos% Y%ypos%. 

	; 이 예제에서 마우스를 이동시켜서 현재 마우스 아래에 있는
	; 창의 제목을 볼 수 있습니다:
	#Persistent
	SetTimer, WatchCursor, 100
	return

	WatchCursor:
	MouseGetPos, , , id, control
	MouseGetPos, , , , hWnd, 2
	WinGetTitle, title, ahk_id %id%
	WinGetClass, class, ahk_id %id%
	ToolTip, ahk_id(WinID):  %id%`nahk_class:  %class%`nWindow_Title:  %title%`nControl(ClassNN):  %control%`nhWnd:  %hWnd%
	return
*/





/*
; 사용자의 마우스 이동 막음
BlockInput, MouseMove

; 사용자의 마우스 이동 허용
BlockInput, MouseMoveOff
*/



/*
; 커서 상태가 작업처리중이면 로딩 중이면 끝날때까지 기다리기
while (A_cursor = "Wait")
	Sleep 3000
*/




/*
;  원하는 값을 못 찾았으면 ErrorLevel 통해서 if문 실행시키기
HTML_Source = 1234
HTML_Source = asdf

MsgBox, % HTML_Source

FoundPos := RegExMatch(HTML_Source, "1234")

MsgBox, % FoundPos

if ErrorLevel
{
	MsgBox, Not Found
}
*/



/*
StringUpper, Name, Name ; 대문자로 바꾸기
*/


/*
; https://autohotkey.com/boards/viewtopic.php?t=23286
; 배열에서 값을 찾아서 그 위치를 반환

HasVal(haystack, needle) {
	if !(IsObject(haystack)) || (haystack.Length() = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}
*/


/*
; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)
*/


/*
; 루프 언제 탈출하는지 확인할 수 있는 예제
Loop{ ; 1

	MsgBox, loop 1 out
	break

	Loop{ ; 2
		
		MsgBox, loop 2 out
		break
		
		Loop{ ; 3
			
			MsgBox, Loop 3 out
			break			
		}

	}

}

MsgBox, all loop out


*/


/*
; gui로 만든 프로그래스 바
TotalLoops = 57
Gui, -Caption +AlwaysOnTop +LastFound
Gui, Add, Text, x12 y9 w100 h20 , S E A R C H I N G . . .
Gui, Add, Progress, w410 Range0-%TotalLoops% cRed vProgress

;~ Gui, Show
Gui, Show, w437 h84, SEARCHING ITEMS


Loop, %TotalLoops%{
	
	GuiControl,,Progress, +1
	Sleep, 100
}

Gui Destroy
;~ GuiClose:
;~ Gui Destroy
;~ ExitApp

MsgBox pause

*/


/*
Array := object() ; 배열 선언
Array.Insert("B1129") ; 배열에 값 넣기
Array := [] ; 배열 초기화
*/


/*
; 화면에서 찾는 값이 여러개일 때 첫번째, 두번째 찾기
driver := ChromeGet()

if(driver.FindElementByXPath("//*[contains(text(), '" keyword "')]"))
{
	MsgBox, % driver.FindElementByXPath("(//*[contains(text(), '" keyword "')])").Attribute("innerText")
	MsgBox, % driver.FindElementByXPath("(//*[contains(text(), '" keyword "')])[2]").Attribute("innerText")
}

; 이건 배열에서 해당하는 것
Loop, % Array_AvailableDate_Sorted.MaxIndex(){
	MsgBox % "Element number " . A_Index . " is " . Array_AvailableDate_Sorted[A_Index]
	MsgBox, % driver.FindElementByXPath("//*[text() = '" Array_AvailableDate_Sorted[A_Index] "']//parent::td//child::a").Attribute("innerText") ; 해당하는 첫 번째 값
	MsgBox, % driver.FindElementByXPath("(//*[text() = '" Array_AvailableDate_Sorted[A_Index] "'])[2]//parent::td//child::a").Attribute("innerText") ; 해당하는 두 번째 값
}


; 아래처럼 해도 된다
Array_AvailableDate_Sorted := object()
Array_AvailableDate_Sorted.Insert("05/10/2018")
driver := ChromeGet()
i = 2
Value := Array_AvailableDate_Sorted[A_Index]
Xpath = (//*[text() = '%Value%'])[%i%]//parent::td//child::a

Loop, % Array_AvailableDate_Sorted.MaxIndex(){
	MsgBox, % driver.FindElementByXPath(Xpath).Attribute("innerText")
}



*/



/*
; 배열에서 값을 찾아서 그 위치를 반환
arr := ["a", "b", "", "d"]

MsgBox % HasVal(arr, "a") "`n"    ; return 1
       . HasVal(arr, "e") "`n"    ; return 0
       . HasVal(arr, "d")         ; return 4

HasVal(haystack, needle) {
	if !(IsObject(haystack)) || (haystack.Length() = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}
*/

/*
https://www.google.com/search?newwindow=1&source=hp&ei=kIGDWpXYLdLajwOtnLT4DA&q=autohotkey+array+sort&oq=autohotkey+array+sort&gs_l=psy-ab.3..0i22i30k1.3181.7947.0.8239.32.20.0.0.0.0.148.1671.8j8.17.0....0...1.1.64.psy-ab..19.13.1372.6..0j35i39k1j0i20i264k1j0i203k1.104.qKKh-koFig0	
; 배열 작은 값부터 차례대로 정렬
QS := new Quickselect
Loop, % Arr.MaxIndex()
	MsgBox, % QS.Select(Arr, A_Index)
*/


/*
; 화면에서 PONumber 찾기
if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
	driver.FindElementByXPath("//*[text() = '" PONumber "']").click()

Xpath = //*[text() = '%PONumber%']
if(driver.FindElementByXPath(Xpath))
	driver.FindElementByXPath(Xpath).click()
*/

/*
; REFRESH THIS PAGE
driver.refresh()
*/

/*
; 브라우저 닫기
;~ driver.quit() ; 이거 작동하지 않음 closing all the browsers
driver.close() ; closing just one tab of the browser
*/


/*
; PO Number 링크 나타날때까지 기다림
Xpath = //*[text() = '%PONumber%']
Sleep 500
while(!driver.FindElementByXPath(Xpath).Attribute("innerText"))
	Sleep 100

Sleep 1000
; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾아서 클릭하기
; 예를들어 MTR1D39747D26 로 검색했으면 검색된 화면에는 MTR1D39747D26 뿐만 아니라 MTR1D39747D26-BO1 등도 같이 표시되기 때문에 딱 원하는 PO Number만 클릭하기 위해
if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
	driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
*/

/*
체크박스 체크되었는지 확인하기
Xpath 를 type="checkbox" 가 있는 곳으로 해야된다. 안 그러면 작동 안 한다
체크가 되어있으면 -1을 리턴하고 
체크가 안 되어있으면 0을 리턴한다
MsgBox, % driver.FindElementByXPath(Xpath).isSelected()
*/

/*
; Element 가 있으면 if문 실행
if(driver.FindElementByXPath(Xpath).isDisplayed()){
	MsgBox, DISPLAYED
}
*/

/*
; Element 가 없으면 if문 실행
if(!driver.FindElementByXPath(Xpath).isDisplayed()){
	MsgBox, NOT DISPLAYED
}
*/

/*
; Element 가 화면에 표시될때까지 기다린 후 클릭하기
loop{
	if(driver.FindElementByXPath(Xpath).isDisplayed()){
		driver.FindElementByXPath(Xpath).click()
		break
	}
	Sleep 100
}
*/

/*
; Element 가 화면에 표시됐는지 알아보는 코드
driver.FindElementByXPath(Xpath).isDisplayed()
*/

/*
; element status 어떤지 알아보는 함수. 0을 반환하면 문제 있는 것
MsgBox, % driver.FindElementByXPath(Xpath).isEnabled()
*/


/*
; 라디오 버튼 중 선택된 버튼이 어떤 것인지 나타내기(if 쓰면 그에 맞는 동작도 시킬 수 있겠지)
action:=["Nothing", "Something", "Everything possible", "In Jesus"]
    
Gui, New
Gui, Add, Text,, What should I do?
Gui, Add, Radio, vRadioGroup, % action[1]
Gui, Add, Radio,, % action[2]
Gui, Add, Radio,, % action[3]
Gui, Add, Radio,, % action[4]
Gui, Show
Return
    
Guiclose:
Gui, Submit
MsgBox % "I'll do " action[RadioGroup]
exitApp
return
*/

/*
; 파일을 삭제
FileDelete, C:\temp files\*.tmp
*/

/*
; 파일이나 폴더 있는지 확인
IfExist, D:\
    MsgBox, 드라이브가 존재합니다.
IfExist, D:\Docs\*.txt
    MsgBox, 적어도 하나의 .txt 파일이 존재합니다.
IfNotExist, C:\Temp\FlagFile.txt
    MsgBox, 목표 파일이 존재하지 않습니다.
*/


/*
; CustomerNoteOnWebVal 변수 안에 있는 메모 내용 CustomerNoteOnWebVal.txt 파일에 저장하기
FileAppend, %CustomerNoteOnWebVal%, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt
*/	

/*
; CustomerNoteOnWebVal.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1
*/

/*
; CustomerNoteOnWebVal.txt 내용을 CustomerNoteOnWebVal 변수에 저장하기
FileRead, CustomerMemoOnLAMBS, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt
*/

/*
OutArray := StrSplit(Input, ",")  ; 콤마 나올때마다 문자열 나누기
Loop % OutArray.Maxindex(){
	OutArray[A_Index] := Trim(OutArray[A_Index])
	MsgBox % "Element number " . A_Index . " is |||" . OutArray[A_Index] . "|||"
}
*/

/* 알파벳과 숫자만 저장 (알페벳과 숫자 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
HumanVerificationCode := RegExReplace(HumanVerificationCode, "[^a-zA-Z0-9]", "")
*/

/*
Data = abc123123
NewStr := RegExReplace(Data, "123$", "xyz")  ; 맨 마지막 123만 xyz로 바꿔서 변수에 "abc123xyz" 저장. abcxyzxyz가 안 되고 맨 끝에만 바뀌는 이유는 $는 끝에만 부합을 허용하기 때문에
*/

/*
NewStr := RegExReplace("abcXYZ123", "abc(.*)123", "aaa$1zzz")  ; $1 역참조를 사용하여 "aaaXYZzzz"을 돌려 줍니다.
NewStr := RegExReplace("abcXYZ123", "abc(.*)123", "$1")  ; $1 역참조를 사용하여 "XYZ"을 돌려 줍니다.
NewStr := RegExReplace("abcXYZ123", "bc(.*)123", "$1")  ; $1 역참조를 사용하여 "aXYZ"을 돌려 줍니다.
MsgBox, % NewStr
*/

/*
; IfEqual 결과 어떻게 나오는지
URLofCustPO = 1234
CurrentURL = 1234

IfNotEqual, URLofCustPO, %CurrentURL%
{
	MsgBox, IT'S NOT EQUAL
}
IfEqual, URLofCustPO, %CurrentURL%
{
	MsgBox, IT'S EQUAL
}

;~ if CurrentURL contains URLofCustPO
if(CurrentURL != URLofCustPO)
{
	MsgBox, % "not matched"
}

MsgBox
*/


/*
; 소스에서 숫자만 추출해서 특정 스트링 뒤에 붙여서 url 만들기
SourceStr = <a _ngcontent-c10="" href="#/order/13371385">MTR1F21F7899C</a>
UnquotedOutputVar = .*/order/(\d*).*>.*

RefinedStr := RegExReplace(SourceStr, UnquotedOutputVar, "$1")  ; UnquotedOutputVar 조건을 보고 $1 역참조를 사용하여 숫자만 RefinedStr 변수에 저장
AddedFinalURL = https://vendoradmin.fashiongo.net/#/order/

URLofCustPO := AddedFinalURL . RefinedStr

MsgBox, % "[SourceStr]`n" . SourceStr . "`n`n`n[URLofCustPO]`n" . URLofCustPO
*/


/*
; RegExMatch 사용해서 문자열 추출하기
; RegExMatch 명령어 사용 시 첫 번째와 두 번째의 괄호 차이에 따른 SubPat2 결과값의 변화, 세 번째와 네 번째의 imU) 유무에 따른 결과값의 변화를 보라
FileRead, Source, %A_ScriptDir%\RegExMatchSample[Do Not Delete].txt ; 공백과 여러 쓸데없는 문자열이 포함된 소스를 읽어서 Source 변수에 저장
MsgBox, % Source


; 이 동작에서는 조건식을 전체 괄호 한 뒤 필요한 부분에 또 괄호를 했을 때 원하는 결과값이 SubPat2 변수에 저장됐다
RegExMatch(Source, "imU)Via\](.*)\[Ship", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2

RegExMatch(Source, "imU)(Via\](.*)\[Ship)", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2


; 이 동작에서는 앞에 imU)를 붙였을 때는 조건식 다음부터 문자열 끝까지의 결과가 추출이 안 되다가 없애니 추출됐다
RegExMatch(Source, "imU)(Via\].*Memo](.*))", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2

RegExMatch(Source, "(Via\].*Memo](.*))", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2
*/


/* 항상 위 메세지 Ok
MsgBox, 262144, Title, Message Placed Here
*/

/* 항상 위 메세지 Yes or No
MsgBox, 4100, Wintitle, Click Ok to continue
*/

/*
	;StringSplit 예제. word_array의 쓰임
	TestString = This is a test.
	; 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 word_array에 저장
	StringSplit, word_array, TestString, `,|%A_Space%, .  ; 점은 제외합니다.

	MsgBox, The 4th word is %word_array4%.

	Colors = red,green,blue
	StringSplit, ColorArray, Colors, `,
	Loop, %ColorArray0%
	{
		this_color := ColorArray%a_index%
		MsgBox, Color number %a_index% is %this_color%.
	}
*/	



/*	
	; 메소드 예제
	; 배열에 들어있는 값 갯수 구하는 메소드
	; arr[0]에 111이 들어가고 arr[10]에 222가 들어가서 배열에 들어간 값 갯수는 2개
	array := Object() 
	array.length := "array_length" 
	array_length(object) 
	{ 
		  current_length := "0" 

		  loop_count := object.maxIndex() + 1 
		  loop % loop_count 
		  { 
				  current_index := a_index - 1 
				  if(object[current_index] != "") 
				  { 
						current_length++ 
				  } 
		  } 

		  return current_length 
	} 


	arr := Object("base", array)              ;==========> 객체 생성시에 다음과 같이 하면 상부에 정의된 length 를 사용할 수 있습니다. 
	arr[0] := 111 
	arr[10] := 222 
	Msgbox % arr.length()
*/


/*
	; 마우스 커서의 현재 위치를 실시간으로 열람
;	MouseGetPos, xpos, ypos
;	Msgbox, The cursor is at X%xpos% Y%ypos%. 

	; 이 예제에서 마우스를 이동시켜서 현재 마우스 아래에 있는
	; 창의 제목을 볼 수 있습니다:
	#Persistent
	SetTimer, WatchCursor, 100
	return

	WatchCursor:
	MouseGetPos, , , id, control
	MouseGetPos, , , , hWnd, 2
	WinGetTitle, title, ahk_id %id%
	WinGetClass, class, ahk_id %id%
	ToolTip, ahk_id(WinID):  %id%`nahk_class:  %class%`nWindow_Title:  %title%`nControl(ClassNN):  %control%`nhWnd:  %hWnd%
	return
*/

/*
q:: ;get control information for the active window (ClassNN and text)
WinGet, vCtlList, ControlList, A
vOutput := ""
Loop, Parse, vCtlList, `n
{
	vCtlClassNN := A_LoopField
	ControlGetText, vText, % vCtlClassNN, A
	vOutput .= vCtlClassNN "`t" vText "`r`n"
}

Clipboard := vOutput
MsgBox, % "done"
return
*/


/*
q::
title = ahk_class Notepad
Loop{
ControlGetText, OutputVar, Edit1, ahk_class FNWND3126
ControlSend, Edit1, %A_Index%, %title%
;~ ControlGetText, A, , Edit66, ahk_class FNWND3126
;~ Send, %a_index%
Sleep 1000
}
*/


/*
; 예제 #4: 실시간으로 활성 창의 콘트롤 리스트를 보여줍니다:
#Persistent
SetTimer, WatchActiveWindow, 200
return
WatchActiveWindow:
WinGet, ControlList, ControlList, A
ToolTip, %ControlList%
return
*/

	
/*
	; 다음의 작동하는 예제는 계속 갱신하면서
	; 현재 마우스 아래의 콘트롤의 위치와 이름을 보여줍니다:
	Loop
	{
		Sleep, 100
		MouseGetPos, , , WhichWindow, WhichControl
		ControlGetPos, x, y, w, h, %WhichControl%, ahk_id %WhichWindow%
		ToolTip, %WhichControl%`nX%X%`tY%Y%`nW%W%`t%H%
	}
*/			
		

/*
	;  크롬 새창에서 열기

	url = http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR1BE9C4B535

	run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) url

	return
*/



/*
웹페이지 자동 로그인 
Loginname = user name
Password = pass word
URL = www.google.com

WB := ComObjCreate("InternetExplorer.Application")
WB.Visible := True
WB.Navigate(URL)
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10

wb.document.getElementById("login").value := Loginname
wb.document.getElementById("password").value := Password
wb.document.getElementsByTagName("Button")[1].Click()

While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10
Msgbox, Something like that i hope!
return
*/


/*
; 크롬 현재 창 url 얻기
hwndChrome := WinExist("ahk_class Chrome_WidgetWin_1")
AccChrome := Acc_ObjectFromWindow(hwndChrome)
AccAddressBar := GetElementByName(AccChrome, "Address and search bar")
MsgBox % AccAddressBar.accValue(0)

GetElementByName(AccObj, name) {
   if (AccObj.accName(0) = name)
      return AccObj
   
   for k, v in Acc_Children(AccObj)
      if IsObject(obj := GetElementByName(v, name))
         return obj
}
*/


/*
; Trim 함수 적용 예
; 문자열 양쪽이나 왼쪽 오른쪽의 문자를 없앤다
text := "  <-Here is Space / Here is Tap ->	"
MsgBox % "No trim:`t '" text "'"
    . "`nTrim:`t '" Trim(text) "'"
    . "`nLTrim:`t '" LTrim(text) "'"
    . "`nRTrim:`t '" RTrim(text) "'"
MsgBox % LTrim("00000123","0")
*/

/*
; continue 예제
; 이 예제는 5개의 MsgBox를 보여줍니다. 각각 6부터 10까지 담고 있습니다.
; 회돌이의 앞쪽 5회에, "continue" 명령어 때문에
; 회돌이가 MsgBox 줄에 도착하기 전에 다시 시작하는 것을 주목하십시오.
Loop, 10
{
    if A_Index <= 5
        continue
    MsgBox %A_Index%
}
*/


/*
; continue 예제
; 이 예제는 5개의 MsgBox를 보여줍니다. 각각 6부터 10까지 담고 있습니다.
; 회돌이의 앞쪽 5회에, "continue" 명령어 때문에
; 회돌이가 MsgBox 줄에 도착하기 전에 다시 시작하는 것을 주목하십시오.
Loop, 10
{
	if A_Index >= 5
	{
		MsgBox, %A_Index% oo
        continue
	}
    MsgBox %A_Index%
}
*/


/*  커서 상태가 작업처리중이면 끝날때까지 기다리기
Loop{
	if(A_cursor = "Wait"){
		Sleep 1000
	}
	else
		break
}

while (A_cursor = "Wait")
	Sleep 2000	
*/


/* 배열 여러개 리턴하는 예제
;~ https://stackoverflow.com/questions/5760058/how-to-return-multiple-arrays-from-a-function-in-javascript

Test := returnMultipleArrays()
MsgBox, % Test[1][1] "`, " Test[1][2] . "`n" . Test[2][1] "`, " Test[2][2] . "`n" . Test[3][1] "`, " Test[3][2]


returnMultipleArrays()
{
 Array1 := ["1", "2"]
 Array2 := ["3", "4"]
 Array3 := ["5", "6"]
 return [Array1, Array2, Array3]
}

*/

/* 다차원 배열 리턴
;~ https://autohotkey.com/board/topic/127830-a-newbs-request-for-help-2d-array/

MyArray := Basic2D()
MsgBox, % MyArray[1, 1]
return

Basic2D() {
    Arr := [] 
    Arr[0, 0] := "0,0"
    Arr[0, 1] := "0,1"
    Arr[1, 0] := "1,0"
    Arr[1, 1] := "1,1"
    return Arr
}
*/


/* 배열로부터 읽기 첫 번째 방법
Array:=[1,3,"ㅋㅋ"]
Loop % Array.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
}
*/


/* 배열로부터 읽기 두 번째 방법
Array:=[1,3,"ㅋㅋ"]
for index, element in Array
{
	MsgBox % "Element number " . index . " is " . element
}
*/


/*
; 다중배열에 있는 값 읽기
; Traversing a multidimensional array
;	https://autohotkey.com/board/topic/81734-traversing-a-multidimensional-array/
MyTable:={}
MyTable[1,1,1]:="Cell: A1-A1-A1"
MyTable[1,1,2]:="Cell: A1-A1-A2"
MyTable[1,2,1]:="Cell: B1-B2-B1"
MyTable[1,2,2]:="Cell: B1-B2-B2"

Loop, % MyTable.MaxIndex()	
For Each, Row in MyTable[A_Index]
	For Each, item in Row
		MsgBox %item%
	
*/


/*
; 다중배열에 있는 값 읽기
MyTable := []

MyTable[1, 1] := "d"
MyTable[1, 2] := "o"
MyTable[1, 3] := "g"
MyTable[2, 1] := "e"

MsgBox % MyTable.1.MaxIndex() "`n" MyTable.2.MaxIndex() ; 배열에 총 몇개의 값이 있는지 확인

For 1st, Row in MyTable
	For 2nd, element in Row
		MsgBox, % "1st : " . 1st . "`n`n2nd : " . 2nd . "`n`n" . "Array" . "[" . 1st . "][" . 2nd . "]" . "`n`nelement : " . element		
*/



/*
	; 배열에 들어있는 값 갯수 구하는 함수
	; arr[0]에 111이 들어가고 arr[10]에 222가 들어가서 배열에 들어간 값 갯수는 2개
	arr := Object() 
	arr[0] := "a" 
	arr[12] := "b" 
	MsgBox, % Obj_Length(arr) ; 2출력 
	return 
	
	Obj_Length(obj) { 
		length := 0 
		for idx in obj 
			length++ 
		return length 
	}
*/




/*
Test := returnMultipleArrays()
MsgBox, % Test[1][1] "`, " Test[1][2] . "`n" . Test[2][1] "`, " Test[2][2] . "`n" . Test[3][1] "`, " Test[3][2]


returnMultipleArrays()
{
	
 Array1%1% := ["1"]
 Array2%1% := ["3"]
 Array3%1% := ["5"]
 	
 Array1%2% := ["2"]
 Array2%2% := ["4"]
 Array3%2% := ["6"]
 
 return [Array1, Array2, Array3]
}
*/





		
		
		
Esc::
 Exitapp
 Reload		