;Prepare our WinHttpRequest object
HttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
;HttpObj.SetProxy(2,"localhost:8888") ;Send data through Fiddler
HttpObj.SetTimeouts(6000,6000,6000,6000) ;Set timeouts to 6 seconds
;HttpObj.Option(6) := False ;disable location-header rediects

;Set our URLs
loginSiteURL := "http://www.autohotkey.com/board/index.php?app=core&module=global&section=login"
loginURL := "http://www.autohotkey.com/board/index.php?app=core&module=global&section=login&do=process"

;Set our login data
username := "Brutosozialprodukt"
password := "xxxxxxxxxxxxxx"
rememberMe := "1"

;Step 1
HttpObj.Open("GET",loginSiteURL)
HttpObj.Send()

;Step 2
RegExMatch(HttpObj.ResponseText,"<input\stype='hidden'\sname='auth_key'\svalue='(\w+)'\s/>",match)
auth_key := match1

;Step 3
loginBody := "auth_key=" auth_key "&ips_username=" username "&ips_password=" password "&rememberMe=" rememberMe

;Step 4/5
HttpObj.Open("POST",loginURL)
HttpObj.SetRequestHeader("Content-Type","application/x-www-form-urlencoded")
HttpObj.Send(loginBody)

;Step 6
If (InStr(HttpObj.ResponseText,"<title>Sign In"))
    MsgBox, The login failed!
Else
    MsgBox, Login was successfull!