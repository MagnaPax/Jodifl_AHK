﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



;		http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR1B08BD00EA


;	msgbox % UrlDownloadToVar("http://www.autohotkey.com/download/CurrentVersion.txt", ByRef Result, UserAgent = "", Proxy = "", ProxyBypass = "")

;	msgbox % UrlDownloadToVar("http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR1B08BD00EA", ByRef Result, UserAgent = "", Proxy = "", ProxyBypass = "")

/*
;IE의 주소창의 주소를 사용해서 해당 페이지의 소스를 출력해 줍니다.
^RButton:: 
WinGetTitle,Title,A
SetTitleMatchMode,3
ControlGetText,url,Edit1,%Title%
res:=UrlDownloadToVar(url,src)
msgbox,%src%
reload
return
*/

/*
F10:: 
WinGetTitle,Title,A
SetTitleMatchMode,3
ControlGetText,url,Edit1,%Title%

;	res:=11UrlDownloadToVar("http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR1B08BD00EA", src)
res:=UrlDownloadToVar(url,src)

msgbox,%src%
reload
return
*/



/*
#SingleInstance, force

UrlDownloadToVar("http://www.autohotkey.com/download/CurrentVersion.txt", urldata)
filedelete, urldata.txt
fileappend, %urldata%,urldata.txt
msgbox, urldata:`n%urldata%
exitapp
*/




;UrlDownloadToVar("http://www.autohotkey.com/download/CurrentVersion.txt", urldata)
UrlDownloadToVar("http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR47CEFB0C", urldata)
clipboard := urldata
MsgBox, % urldata




/*
a := UrlDownloadToVar("http://jm.greatnow.com/ahk_test_20060618140900_517b.txt")
b := UrlDownloadToVar("http://jm.greatnow.com/ahk_spaces_20060618140900_224b.txt")
fileappend, %a%, a.txt
fileappend, %b%, b.txt
*/



;Esc::
;Exitapp















UrlDownloadToVar(URL, ByRef Result, UserAgent = "", Proxy = "", ProxyBypass = "") {
  ; Requires Windows Vista, Windows XP, Windows 2000 Professional, Windows NT Workstation 4.0,
  ; Windows Me, Windows 98, or Windows 95.
  ; Requires Internet Explorer 3.0 or later.
  
  hModule := DllCall("LoadLibrary", "Str", "wininet.dll") 

  AccessType := Proxy != "" ? 3 : 1
  ;INTERNET_OPEN_TYPE_PRECONFIG                    0   // use registry configuration 
  ;INTERNET_OPEN_TYPE_DIRECT                       1   // direct to net 
  ;INTERNET_OPEN_TYPE_PROXY                        3   // via named proxy 
  ;INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY  4   // prevent using java/script/INS 

	io := DllCall("wininet\InternetOpenA" 
	, "Str", UserAgent ;lpszAgent 
	, "UInt", AccessType 
	, "Str", Proxy 
	, "Str", ProxyBypass 
	, "UInt", 0) ;dwFlags 
	
	iou := DllCall("wininet\InternetOpenUrlA" 
	, "UInt", io 
	, "Str", url 
	, "Str", "" ;lpszHeaders 
	, "UInt", 0 ;dwHeadersLength 
	, "UInt", 0x80000000 ;dwFlags: INTERNET_FLAG_RELOAD = 0x80000000 // retrieve the original item 
	, "UInt", 0) ;dwContext 
	
	If (ErrorLevel != 0 or iou = 0) { 
	  DllCall("FreeLibrary", "UInt", hModule) 
	  return 0 
	} 
	
	VarSetCapacity(buffer, 10240, 0)
	VarSetCapacity(BytesRead, 4, 0)
	
	Loop 
	{ 
	  ;http://msdn.microsoft.com/library/en-us/wininet/wininet/internetreadfile.asp
	  irf := DllCall("wininet\InternetReadFile", "UInt", iou, "UInt", &buffer, "UInt", 10240, "UInt", &BytesRead) 
	  VarSetCapacity(buffer, -1) ;to update the variable's internally-stored length
	  
    BytesRead_ = 0 ; reset
	  Loop, 4  ; Build the integer by adding up its bytes. (From ExtractInteger-function)
      BytesRead_ += *(&BytesRead + A_Index-1) << 8*(A_Index-1) ;Bytes read in this very DllCall
    
    ; To ensure all data is retrieved, an application must continue to call the
    ; InternetReadFile function until the function returns TRUE and the lpdwNumberOfBytesRead parameter equals zero.
    If (irf = 1 and BytesRead_ = 0)
	    break
	  Else ; append the buffer's contents
	    Result .= SubStr(buffer, 1, BytesRead_) 
    
    /* optional: retrieve only a part of the file
    BytesReadTotal += BytesRead_
    If (BytesReadTotal >= 30000) ; only read the first x bytes
	    break                      ; (will be a multiple of the buffer size, if the file is not smaller; trim if neccessary)
	  */
	}
	
	DllCall("wininet\InternetCloseHandle",  "UInt", iou) 
	DllCall("wininet\InternetCloseHandle",  "UInt", io) 
	DllCall("FreeLibrary", "UInt", hModule)
}