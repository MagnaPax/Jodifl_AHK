#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



; 이미 존재하는 크롬 창 제어하기 위한 함수. 아래와 같은 코멘드 라인으로 열린 크롬이라야 이 함수로 컨트롤 가능하다
; Start Chrome with command line: chrome.exe --remote-debugging-port=9222
ChromeGet(IP_Port := "127.0.0.1:9222") {
	driver := ComObjCreate("Selenium.ChromeDriver")
	driver.SetCapability("debuggerAddress", IP_Port)
	driver.Start()
	
	;~ driver.AddArgument("--start-maximized")
	;~ driver.ExecuteScript("document.body.style.zoom = '100%';")
	;~ driver.setCapability("ignoreZoomSetting", true)
	
	return driver
}
