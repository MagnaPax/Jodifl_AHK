;Prepare our WinHttpRequest object
HttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
;HttpObj.SetProxy(2,"localhost:8888") ;Send data through Fiddler
HttpObj.SetTimeouts(6000,6000,6000,6000) ;Set timeouts to 6 seconds
;HttpObj.Option(6) := False ;disable location-header rediects

;#Include function.ahk

autologin := "on"

username = customer3
password = Jo123456789

username =
password =
loginURL = http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR47CEFB0C

autologin := "on"


;Step 1
loginBody := "username=" username "&password=" password "&autologin=" autologin "&login=Login"

;Step 2/3
HttpObj.Open("POST",loginURL)
HttpObj.SetRequestHeader("Content-Type","application/x-www-form-urlencoded")
HttpObj.Send(loginBody)

;Step 4
If (InStr(HttpObj.ResponseText,"<title>Login"))
    MsgBox, The login failed!
Else
    MsgBox, Login was successfull!





;url := "http://www.google.com"

; Example: Download text to a variable:
;HttpRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
HttpRequest.Open("GET", loginURL)
HttpRequest.Send()
this_text := HttpRequest.ResponseText
html := ComObjCreate("htmlfile")
html.write(this_text)
MsgBox, % loginURL

; Loop through all links and add them to link_list variable with a new line
Loop % html.links.length
  ;link_list .= html.links[A_Index - 1].href . "`r"
  link_list .= html.links[A_Index - 1].href . "lblShipCompanyName"

; Certain links for relative and have text like 'about:/services/' replace the about with url
StringReplace, link_list, link_list, about:, %url%, A
msgbox link_list`n`n`n%link_list%


MsgBox, this_text`n%this_text%`n`n`nhtml`n%html%
;MsgBox, %html%
;FindInfoInFASHIONGO(this_text)
		MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nState : %State%`n`nPhone : %Phone%`n`nEmail : %Email%
