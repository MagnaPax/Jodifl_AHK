﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;ControlSend, Edit4, "CLOVER CHIC BOUTIQUE"f{Tab}, UPS WorldShip - Remote Workstation
;ControlSend, Edit4, {Tab}, UPS WorldShip - Remote Workstation

ControlSend, , {F2}, UPS WorldShip - Remote Workstation
Sleep 500

ControlSend, , {F5}, UPS WorldShip - Remote Workstation
Sleep 500


ControlSetText, Edit4, CLOVER CHIC BOUTIQUE, UPS WorldShip - Remote Workstation

ControlSetText, Edit5, ASHLEY HAZLEWOOD, UPS WorldShip - Remote Workstation



ExitApp
  
Esc::
 Exitapp