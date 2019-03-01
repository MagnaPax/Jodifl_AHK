#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



;~ #Include function.ahk


global InvoiceBalance



		
	; 마우스 커서는 사용자의 물리적 마우스 이동에 반응하지 않습니다
	BlockInput, MouseMove
			
	
	
	CoordMode, mouse, relative
	
	;LAMBS활성화 후 화면 초기화
	Start()
	

	; LAMBS에 있는 인보이스 밸런스 값 읽어오기
	GetInvoiceBalanceOnLAMBS()


	
	;CC 버튼 찾고 클릭
	FindCCButtonAndClickIt()
	Sleep 2000
		
	
	windowtitle := "Credit Card ( P999131 )"
	; CC번호 저장
	ControlGetText, CCNumbers, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	; 만료일 저장
	ControlGetText, ExpDate, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	; CVS 저장
	ControlGetText, CVV, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate 를 넘겨서 Month 와 Year 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate, "")
	
	; CCNumbers 받아서 4개씩 읽어서 PartiallyCCNum1,2,3,4,5에 넣기
	StartingPos = 1
	
	loop, 5{
		PartiallyCCNum%A_Index% := SubStr(CCNumbers, StartingPos, 4)
		StartingPos := StartingPos + 4
	}
	
	
	
	
	; 두 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	;MouseClick, l, 240, 390
	MouseClick, l, 200, 465
	
	Sleep 100
	
	ControlGetText, CCNumbers2, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate2, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV2, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate2 를 넘겨서 Month2 와 Year2 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate2, 2)

	
	StartingPos = 1
	
	loop, 5{
		2PartiallyCCNum%A_Index% := SubStr(CCNumbers2, StartingPos, 4)
		StartingPos := StartingPos + 4
	}
	
	
	
	
	
	; 세 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	;MouseClick, l, 240, 410
	MouseClick, l, 200, 480
	
	Sleep 100
	
	ControlGetText, CCNumbers3, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate3, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV3, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate3 를 넘겨서 Month3 와 Year3 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate3, 3)

	
	StartingPos = 1
	
	loop, 5{
		3PartiallyCCNum%A_Index% := SubStr(CCNumbers3, StartingPos, 4)
		StartingPos := StartingPos + 4
	}	
	
	
	
	
	; 네 번째 CC 읽고 처리하기
	WinActivate, %windowtitle%
	;MouseClick, l, 240, 430
	MouseClick, l, 200, 495
	
	Sleep 100
	
	ControlGetText, CCNumbers4, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%
	ControlGetText, ExpDate4, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
	ControlGetText, CVV4, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
	
	; ExpDate4 를 넘겨서 Month4 와 Year4 에 값 넣기
	OrganizingLAMBSCCinfo(ExpDate4, 4)

	
	StartingPos = 1
	
	loop, 5{
		4PartiallyCCNum%A_Index% := SubStr(CCNumbers4, StartingPos, 4)
		StartingPos := StartingPos + 4
	}		
	
	



;	MsgBox, 같은 값 있나 비교`n`n%CCNumbers%   %CCNumbers2%   %CCNumbers3%   %CCNumbers4%`n`n`n`n%ExpDate%   %ExpDate2%   %ExpDate3%   %ExpDate4%`n`n`n`n%CVV%   %CVV2%   %CVV3%   %CVV4%
		
	
	
	
	; 이것이 중요한 처리.
	; 이렇게 값을 초기화 해 줌으로서 다른 함수에서 램스의 cc 정보를 얻어갈 때 각 CCNumbers2,3,4 변수에 값이 있는지 없는지 여부를 갖고 행동을 결정하는 데 사용된다
	; 두 번째 카드 번호가 첫 번째와 같으면 두 번째 카드정보 포함 나머지 cc 정보 값들을 초기화하기
	
	; 하지만 만약 두 번째 cc 정보는 있는데 세 번째 cc가 없으면 어떡하냐?
	; 할 수 없음. 그냥 3,4번째 cc 정보는 두 번째 cc 정보가 그대로 카피됨.
	; 그것까지 구현하기 귀찮음
	IfEqual, CCNumbers, %CCNumbers2%
	{
		Loop, 4{
			CCNumbers%A_Index% :=
			ExpDate%A_Index% :=
			CVV%A_Index% :=
		}
	}
	

	
	
	
	
	
	
	
	
	
	;MsgBox, %PartiallyCCNum1%    %PartiallyCCNum2%    %PartiallyCCNum3%    %PartiallyCCNum4%    %PartiallyCCNum5%
	
	
	;cc 창을 닫기
;	WinClose, %windowtitle%
	
	

	; 사용자에게 마우스 커서를 이동하도록 허용합니다.
	BlockInput, MouseMoveOff
	
	
	
	

	;cc정보 띄워주기
/*	
		MsgBox, %InvoiceBalance%`n`n%PartiallyCCNum1%   %PartiallyCCNum2%   %PartiallyCCNum3%   %PartiallyCCNum4%   %PartiallyCCNum5%`n`n%ExpDate%`n`n%CVV%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%2PartiallyCCNum1%   %2PartiallyCCNum2%   %2PartiallyCCNum3%   %2PartiallyCCNum4%   %2PartiallyCCNum5%`n`n%ExpDate2%`n`n%CVV2%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%3PartiallyCCNum1%   %3PartiallyCCNum2%   %3PartiallyCCNum3%   %3PartiallyCCNum4%   %3PartiallyCCNum5%`n`n%ExpDate3%`n`n%CVV3%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%4PartiallyCCNum1%   %4PartiallyCCNum2%   %4PartiallyCCNum3%   %4PartiallyCCNum4%   %4PartiallyCCNum5%`n`n%ExpDate4%`n`n%CVV4%`n`n%BillingAdd1%`n`n%BillingZip%
			
	
	
	;cc 창을 닫기
	WinClose, %windowtitle%
*/		
		
	
	
	
	
	
;~ /*	
;	MsgBox, Invoice Balance : %InvoiceBalance%`n`nCC number : %PartiallyCCNum1%    %PartiallyCCNum2%    %PartiallyCCNum3%    %PartiallyCCNum4%    %PartiallyCCNum5%`n`nExpiration Date : %ExpDate%`n`nSecurity Code : %CVV%`n`nAddress1 : %Address1%`n`nZip Code : %ZipCode%


	; CCNumbers 에 값이 있으면, 즉 첫 번째 CC 번호가 있으면 cc정보 띄워주기
	if(CCNumbers)
	{
		MsgBox, %InvoiceBalance%`n`n%PartiallyCCNum1%   %PartiallyCCNum2%   %PartiallyCCNum3%   %PartiallyCCNum4%   %PartiallyCCNum5%`n`n%ExpDate%`n`n%CVV%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%2PartiallyCCNum1%   %2PartiallyCCNum2%   %2PartiallyCCNum3%   %2PartiallyCCNum4%   %2PartiallyCCNum5%`n`n%ExpDate2%`n`n%CVV2%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%3PartiallyCCNum1%   %3PartiallyCCNum2%   %3PartiallyCCNum3%   %3PartiallyCCNum4%   %3PartiallyCCNum5%`n`n%ExpDate3%`n`n%CVV3%`n`n%BillingAdd1%`n`n%BillingZip%`n`n`n`n%InvoiceBalance%`n`n%4PartiallyCCNum1%   %4PartiallyCCNum2%   %4PartiallyCCNum3%   %4PartiallyCCNum4%   %4PartiallyCCNum5%`n`n%ExpDate4%`n`n%CVV4%`n`n%BillingAdd1%`n`n%BillingZip%
		
		;~ return
	}
	else{ ;CC창 안에 정보 없으면 안내문 띄운 뒤 5초 뒤 자동으로 나가기
		MsgBox, , , No Credit Card Information at LAMBS`n`n`nTHIS WINDOW WILL BE CLOSED IN 5 SECONDS, 5

		WinClose, %windowtitle%
	}
	
	
	;cc 창을 닫기
	WinClose, %windowtitle%
*/		
	
Exitapp
	
	
	
; LAMBS에 있는 인보이스 밸런스 값 읽어오기
GetInvoiceBalanceOnLAMBS()
{

	;LAMBS활성화 후 화면 초기화
	Start()
	
	;Invoice Balance 값 얻기
	MouseClick, l, 870, 594
	Send, ^a^c
	Sleep 700
	InvoiceBalance := Clipboard		

/*
	;Invoice Balance 값 얻기
	MouseMove, 870, 594
	Sleep 100		
	MouseGetPos, , , , control
	ControlGetText, InvoiceBalance, %control%, LAMBS
*/

;	MsgBox, % InvoiceBalance
	
	return
}

	
	

	;LAMBS활성화 후 화면 초기화 하기
	Start(){
		
		;LAMBS Window 활성화 하기
		WinActivate, LAMBS -  Garment Manufacturer & Wholesale Software
		windowtitle = LAMBS -  Garment Manufacturer & Wholesale Software
		CheckTheWindowPresentAndActiveIt(windowtitle)

		;Hide All 클릭해서 메뉴 바 없애기
		ClickAtThePoint(213, 65)
		
		return
	}
	
	
	;CC 버튼 찾고 클릭
	FindCCButtonAndClickIt(){					
		Text:="|<>*140$58.D002E0S005000882000M000UUE001UCQSLV0SQy0m+98409YM2DcYUE3YFU8UWG10GF50X2982194Hm7bYs7bYDU"

		if ok:=FindText(1373,149,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		}
				
	}
	
	
	
	
	
	FindCCButtonAndClickIt2(){
		

		
		
		jpgLocation = %A_ScripDir%PICTURES\CCButton.png
		jpgLocation2 = %A_ScripDir%PICTURES\CCButton2.png
		
		Loop, 10{
			ImageSearch, pX, pY, 0, 0, 2000, 300, %jpgLocation%
		
			;그림이 찾아졌으면 클릭
			if(ErrorLevel = 0){
				
				pX += 20
				pY += 10
		
				MouseClick, l, %pX%, %pY%, 2

				sleep 500
				
				return 0
			}
			;못 찾았으면 CCButton2 그림으로 다시 찾고 클릭
			else if(ErrorLevel = 1){
				;MsgBox, Couldn't find it!
				ImageSearch, pX, pY, 0, 0, 2000, 300, %jpgLocation2%
				pX += 20
				pY += 10
				MouseClick, l, %pX%, %pY%, 2
				
				return 0
				
				;break
			}
			else if(ErrorLevel = 2){
				MsgBox, Unexpected Error Occur!
				;break
			}
		
		}
		MsgBox, Out!
	}	
		
	
	
OrganizingLAMBSCCinfo(ObtainedExpDate, __Num__){
	
	
	; LAMBS 에 있는 첫 번째 cc 정보로 카드 승인 취소 나면 CCNumbers2, CCNumbers3 등 변수를 넘겨주면서 함수를 호출할테니
	; 받은 신용카드번호, 보안번호를 전역변수인 CCNumbers 와 CVV에 
;	CCNumbers := ObtainedCCNumbers
;	CVV := ObtainedCVV
	

	
	; Month 찾기.
	FoundPos := RegExMatch(ObtainedExpDate, "mU)((\d\d)\s/)", SubPat)

	; Month 찾았으면 전역변수인 Month 에 값 넣기
	if(FoundPos){
		Month%__Num__% := SubPat2
		
		;MsgBox, % StrLen(Month)
		
		; Month 문자 길이를 StringLength 에 넣기
		;StringLength := StrLen(Month)
		
		;MsgBox, % StringLength
		
		; StringLength 길이가 1이면 숫자가 하나만 있다는 뜻이므로 왼쪽에 0 붙여주기 
		;IfEqual, StringLength, 1
			;Month = 0%Month%
		
		;MsgBox, % Month
	}
		
	
	; Year 찾기. 
	FoundPos := RegExMatch(ObtainedExpDate, "mU)(.*/\s(\d\d\d\d))", SubPat)

	; Year 찾았으면 전역변수인 Year 에 값 넣기
	if(FoundPos){
		Year%__Num__% := SubPat2
	}
	
			
	
	
	return
}
	
	

	
	;지정된 창이 존재하는지 무한정 확인 후 그 창 활성화
	CheckTheWindowPresentAndActiveIt(windowtitle){
		WinWait, % windowtitle
		WinActivate, % windowtitle
		return
	}
		
	
	
	
	;위치 받아서 클릭하기
	ClickAtThePoint(XPoint, YPoint){
		MouseClick, l, XPoint, YPoint, 1
		Sleep 1000
		return
	}
		
	
	
	


;===== Copy The Following Functions To Your Own Code Just once =====


; Note: parameters of the X,Y is the center of the coordinates,
; and the W,H is the offset distance to the center,
; So the search range is (X-W, Y-H)-->(X+W, Y+H).
; err1 is the character "0" fault-tolerant in percentage.
; err0 is the character "_" fault-tolerant in percentage.
; Text can be a lot of text to find, separated by "|".
; ruturn is a array, contains the X,Y,W,H,Comment results of Each Find.

FindText(x,y,w,h,err1,err0,text)
{
  xywh2xywh(x-w,y-h,2*w+1,2*h+1,x,y,w,h)
  if (w<1 or h<1)
    return, 0
  bch:=A_BatchLines
  SetBatchLines, -1
  ;--------------------------------------
  GetBitsFromScreen(x,y,w,h,Scan0,Stride,bits)
  ;--------------------------------------
  sx:=0, sy:=0, sw:=w, sh:=h, arr:=[]
  Loop, 2 {
  Loop, Parse, text, |
  {
    v:=A_LoopField
    IfNotInString, v, $, Continue
    Comment:="", e1:=err1, e0:=err0
    ; You Can Add Comment Text within The <>
    if RegExMatch(v,"<([^>]*)>",r)
      v:=StrReplace(v,r), Comment:=Trim(r1)
    ; You can Add two fault-tolerant in the [], separated by commas
    if RegExMatch(v,"\[([^\]]*)]",r)
    {
      v:=StrReplace(v,r), r1.=","
      StringSplit, r, r1, `,
      e1:=r1, e0:=r2
    }
    StringSplit, r, v, $
    color:=r1, v:=r2
    StringSplit, r, v, .
    w1:=r1, v:=base64tobit(r2), h1:=StrLen(v)//w1
    if (r0<2 or h1<1 or w1>sw or h1>sh or StrLen(v)!=w1*h1)
      Continue
    ;--------------------------------------------
    if InStr(color,"-")
    {
      r:=e1, e1:=e0, e0:=r, v:=StrReplace(v,"1","_")
      v:=StrReplace(StrReplace(v,"0","1"),"_","0")
    }
    mode:=InStr(color,"*") ? 1:0
    color:=RegExReplace(color,"[*\-]") . "@"
    StringSplit, r, color, @
    color:=Round(r1), n:=Round(r2,2)+(!r2)
    n:=Floor(255*3*(1-n)), k:=StrLen(v)*4
    VarSetCapacity(ss, sw*sh, Asc("0"))
    VarSetCapacity(s1, k, 0), VarSetCapacity(s0, k, 0)
    VarSetCapacity(rx, 8, 0), VarSetCapacity(ry, 8, 0)
    len1:=len0:=0, j:=sw-w1+1, i:=-j
    ListLines, Off
    Loop, Parse, v
    {
      i:=Mod(A_Index,w1)=1 ? i+j : i+1
      if A_LoopField
        NumPut(i, s1, 4*len1++, "int")
      else
        NumPut(i, s0, 4*len0++, "int")
    }
    ListLines, On
    e1:=Round(len1*e1), e0:=Round(len0*e0)
    ;--------------------------------------------
    if PicFind(mode,color,n,Scan0,Stride,sx,sy,sw,sh
      ,ss,s1,s0,len1,len0,e1,e0,w1,h1,rx,ry)
    {
      rx+=x, ry+=y
      arr.Push(rx,ry,w1,h1,Comment)
    }
  }
  if (arr.MaxIndex())
    Break
  if (A_Index=1 and err1=0 and err0=0)
    err1:=0.05, err0:=0.05
  else Break
  }
  SetBatchLines, %bch%
  return, arr.MaxIndex() ? arr:0
}

PicFind(mode, color, n, Scan0, Stride
  , sx, sy, sw, sh, ByRef ss, ByRef s1, ByRef s0
  , len1, len0, err1, err0, w, h, ByRef rx, ByRef ry)
{
  static MyFunc
  if !MyFunc
  {
    x32:="5589E583EC408B45200FAF45188B551CC1E20201D08945F"
    . "48B5524B80000000029D0C1E00289C28B451801D08945D8C74"
    . "5F000000000837D08000F85F00000008B450CC1E81025FF000"
    . "0008945D48B450CC1E80825FF0000008945D08B450C25FF000"
    . "0008945CCC745F800000000E9AC000000C745FC00000000E98"
    . "A0000008B45F483C00289C28B451401D00FB6000FB6C02B45D"
    . "48945EC8B45F483C00189C28B451401D00FB6000FB6C02B45D"
    . "08945E88B55F48B451401D00FB6000FB6C02B45CC8945E4837"
    . "DEC007903F75DEC837DE8007903F75DE8837DE4007903F75DE"
    . "48B55EC8B45E801C28B45E401D03B45107F0B8B55F08B452C0"
    . "1D0C600318345FC018345F4048345F0018B45FC3B45240F8C6"
    . "AFFFFFF8345F8018B45D80145F48B45F83B45280F8C48FFFFF"
    . "FE9A30000008B450C83C00169C0E803000089450CC745F8000"
    . "00000EB7FC745FC00000000EB648B45F483C00289C28B45140"
    . "1D00FB6000FB6C069D02B0100008B45F483C00189C18B45140"
    . "1C80FB6000FB6C069C04B0200008D0C028B55F48B451401D00"
    . "FB6000FB6C06BC07201C83B450C730B8B55F08B452C01D0C60"
    . "0318345FC018345F4048345F0018B45FC3B45247C948345F80"
    . "18B45D80145F48B45F83B45280F8C75FFFFFF8B45242B45488"
    . "3C0018945488B45282B454C83C00189454C8B453839453C0F4"
    . "D453C8945D8C745F800000000E9E3000000C745FC00000000E"
    . "9C70000008B45F80FAF452489C28B45FC01D08945F48B45408"
    . "945E08B45448945DCC745F000000000EB708B45F03B45387D2"
    . "E8B45F08D1485000000008B453001D08B108B45F401D089C28"
    . "B452C01D00FB6003C31740A836DE001837DE00078638B45F03"
    . "B453C7D2E8B45F08D1485000000008B453401D08B108B45F40"
    . "1D089C28B452C01D00FB6003C30740A836DDC01837DDC00783"
    . "08345F0018B45F03B45D87C888B551C8B45FC01C28B4550891"
    . "08B55208B45F801C28B45548910B801000000EB2990EB01908"
    . "345FC018B45FC3B45480F8C2DFFFFFF8345F8018B45F83B454"
    . "C0F8C11FFFFFFB800000000C9C25000"
    x64:="554889E54883EC40894D10895518448945204C894D288B4"
    . "5400FAF45308B5538C1E20201D08945F48B5548B8000000002"
    . "9D0C1E00289C28B453001D08945D8C745F000000000837D100"
    . "00F85000100008B4518C1E81025FF0000008945D48B4518C1E"
    . "80825FF0000008945D08B451825FF0000008945CCC745F8000"
    . "00000E9BC000000C745FC00000000E99A0000008B45F483C00"
    . "24863D0488B45284801D00FB6000FB6C02B45D48945EC8B45F"
    . "483C0014863D0488B45284801D00FB6000FB6C02B45D08945E"
    . "88B45F44863D0488B45284801D00FB6000FB6C02B45CC8945E"
    . "4837DEC007903F75DEC837DE8007903F75DE8837DE4007903F"
    . "75DE48B55EC8B45E801C28B45E401D03B45207F108B45F0486"
    . "3D0488B45584801D0C600318345FC018345F4048345F0018B4"
    . "5FC3B45480F8C5AFFFFFF8345F8018B45D80145F48B45F83B4"
    . "5500F8C38FFFFFFE9B60000008B451883C00169C0E80300008"
    . "94518C745F800000000E98F000000C745FC00000000EB748B4"
    . "5F483C0024863D0488B45284801D00FB6000FB6C069D02B010"
    . "0008B45F483C0014863C8488B45284801C80FB6000FB6C069C"
    . "04B0200008D0C028B45F44863D0488B45284801D00FB6000FB"
    . "6C06BC07201C83B451873108B45F04863D0488B45584801D0C"
    . "600318345FC018345F4048345F0018B45FC3B45487C848345F"
    . "8018B45D80145F48B45F83B45500F8C65FFFFFF8B45482B859"
    . "000000083C0018985900000008B45502B859800000083C0018"
    . "985980000008B45703945780F4D45788945D8C745F80000000"
    . "0E90B010000C745FC00000000E9EC0000008B45F80FAF45488"
    . "9C28B45FC01D08945F48B85800000008945E08B85880000008"
    . "945DCC745F000000000E9800000008B45F03B45707D368B45F"
    . "04898488D148500000000488B45604801D08B108B45F401D04"
    . "863D0488B45584801D00FB6003C31740A836DE001837DE0007"
    . "8778B45F03B45787D368B45F04898488D148500000000488B4"
    . "5684801D08B108B45F401D04863D0488B45584801D00FB6003"
    . "C30740A836DDC01837DDC00783C8345F0018B45F03B45D80F8"
    . "C74FFFFFF8B55388B45FC01C2488B85A000000089108B55408"
    . "B45F801C2488B85A80000008910B801000000EB2F90EB01908"
    . "345FC018B45FC3B85900000000F8C05FFFFFF8345F8018B45F"
    . "83B85980000000F8CE6FEFFFFB8000000004883C4405DC390"
    MCode(MyFunc, A_PtrSize=8 ? x64:x32)
  }
  return, DllCall(&MyFunc, "int",mode
    , "uint",color, "int",n, "ptr",Scan0, "int",Stride
    , "int",sx, "int",sy, "int",sw, "int",sh
    , "ptr",&ss, "ptr",&s1, "ptr",&s0
    , "int",len1, "int",len0, "int",err1, "int",err0
    , "int",w, "int",h, "int*",rx, "int*",ry)
}

xywh2xywh(x1,y1,w1,h1,ByRef x,ByRef y,ByRef w,ByRef h)
{
  SysGet, zx, 76
  SysGet, zy, 77
  SysGet, zw, 78
  SysGet, zh, 79
  left:=x1, right:=x1+w1-1, up:=y1, down:=y1+h1-1
  left:=left<zx ? zx:left, right:=right>zx+zw-1 ? zx+zw-1:right
  up:=up<zy ? zy:up, down:=down>zy+zh-1 ? zy+zh-1:down
  x:=left, y:=up, w:=right-left+1, h:=down-up+1
}

GetBitsFromScreen(x,y,w,h,ByRef Scan0,ByRef Stride,ByRef bits)
{
  VarSetCapacity(bits,w*h*4,0), bpp:=32
  Scan0:=&bits, Stride:=((w*bpp+31)//32)*4
  Ptr:=A_PtrSize ? "UPtr" : "UInt", PtrP:=Ptr . "*"
  win:=DllCall("GetDesktopWindow", Ptr)
  hDC:=DllCall("GetWindowDC", Ptr,win, Ptr)
  mDC:=DllCall("CreateCompatibleDC", Ptr,hDC, Ptr)
  ;-------------------------
  VarSetCapacity(bi, 40, 0), NumPut(40, bi, 0, "int")
  NumPut(w, bi, 4, "int"), NumPut(-h, bi, 8, "int")
  NumPut(1, bi, 12, "short"), NumPut(bpp, bi, 14, "short")
  ;-------------------------
  if hBM:=DllCall("CreateDIBSection", Ptr,mDC, Ptr,&bi
    , "int",0, PtrP,ppvBits, Ptr,0, "int",0, Ptr)
  {
    oBM:=DllCall("SelectObject", Ptr,mDC, Ptr,hBM, Ptr)
    DllCall("BitBlt", Ptr,mDC, "int",0, "int",0, "int",w, "int",h
      , Ptr,hDC, "int",x, "int",y, "uint",0x00CC0020|0x40000000)
    DllCall("RtlMoveMemory","ptr",Scan0,"ptr",ppvBits,"ptr",Stride*h)
    DllCall("SelectObject", Ptr,mDC, Ptr,oBM)
    DllCall("DeleteObject", Ptr,hBM)
  }
  DllCall("DeleteDC", Ptr,mDC)
  DllCall("ReleaseDC", Ptr,win, Ptr,hDC)
}

base64tobit(s)
{
  ListLines, Off
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  StringCaseSense, On
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:=(i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,A_LoopField,v)
  }
  StringCaseSense, Off
  s:=SubStr(s,1,InStr(s,"1",0,0)-1)
  s:=RegExReplace(s,"[^01]+")
  ListLines, On
  return, s
}

bit2base64(s)
{
  ListLines, Off
  s:=RegExReplace(s,"[^01]+")
  s.=SubStr("100000",1,6-Mod(StrLen(s),6))
  s:=RegExReplace(s,".{6}","|$0")
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:="|" . (i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,v,A_LoopField)
  }
  ListLines, On
  return, s
}

ASCII(s)
{
  if RegExMatch(s,"(\d+)\.([\w+/]{3,})",r)
  {
    s:=RegExReplace(base64tobit(r2),".{" r1 "}","$0`n")
    s:=StrReplace(StrReplace(s,"0","_"),"1","0")
  }
  else s=
  return, s
}

MCode(ByRef code, hex)
{
  ListLines, Off
  bch:=A_BatchLines
  SetBatchLines, -1
  VarSetCapacity(code, StrLen(hex)//2)
  Loop, % StrLen(hex)//2
    NumPut("0x" . SubStr(hex,2*A_Index-1,2), code, A_Index-1, "char")
  Ptr:=A_PtrSize ? "UPtr" : "UInt"
  DllCall("VirtualProtect", Ptr,&code, Ptr
    ,VarSetCapacity(code), "uint",0x40, Ptr . "*",0)
  SetBatchLines, %bch%
  ListLines, On
}

; You can put the text library at the beginning of the script,
; and Use Pic(Text,1) to add the text library to Pic()'s Lib,
; Use Pic("comment1|comment2|...") to get text images from Lib
Pic(comments, add_to_Lib=0) {
  static Lib:=[]
  if (add_to_Lib)
  {
    re:="<([^>]*)>[^$]+\$\d+\.[\w+/]{3,}"
    Loop, Parse, comments, |
      if RegExMatch(A_LoopField,re,r)
        Lib[Trim(r1)]:=r
  }
  else
  {
    text:=""
    Loop, Parse, comments, |
      text.="|" . Lib[Trim(A_LoopField)]
    return, text
  }
}


/***** C source code of machine code *****

int __attribute__((__stdcall__)) PicFind(int mode
  , unsigned int c, int n, unsigned char * Bmp
  , int Stride, int sx, int sy, int sw, int sh
  , char * ss, int * s1, int * s0
  , int len1, int len0, int err1, int err0
  , int w, int h, int * rx, int * ry)
{
  int x, y, o=sy*Stride+sx*4, j=Stride-4*sw, i=0;
  int r, g, b, rr, gg, bb, e1, e0;
  if (mode==0)  // Color Mode
  {
    rr=(c>>16)&0xFF; gg=(c>>8)&0xFF; bb=c&0xFF;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
      {
        r=Bmp[2+o]-rr; g=Bmp[1+o]-gg; b=Bmp[o]-bb;
        if (r<0) r=-r; if (g<0) g=-g; if (b<0) b=-b;
        if (r+g+b<=n) ss[i]='1';
      }
  }
  else  // Gray Threshold Mode
  {
    c=(c+1)*1000;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
        if (Bmp[2+o]*299+Bmp[1+o]*587+Bmp[o]*114<c)
          ss[i]='1';
  }
  w=sw-w+1; h=sh-h+1;
  j=len1>len0 ? len1 : len0;
  for (y=0; y<h; y++)
  {
    for (x=0; x<w; x++)
    {
      o=y*sw+x; e1=err1; e0=err0;
      for (i=0; i<j; i++)
      {
        if (i<len1 && ss[o+s1[i]]!='1' && (--e1)<0)
          goto NoMatch;
        if (i<len0 && ss[o+s0[i]]!='0' && (--e0)<0)
          goto NoMatch;
      }
      rx[0]=sx+x; ry[0]=sy+y;
      return 1;
      NoMatch:
      continue;
    }
  }
  return 0;
}

*/


;================= The End =================

;



F5::
 Exitapp
	
	
	
	
	
	
	
	
	
	

Exitapp

Esc::
Exitapp	