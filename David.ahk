#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.




MsgBox, 262144, Functions, [Alt+z] - ( available_allocate_qty > 0 )  `n`n`n`n`n`n`n`n[Alt+``] Dropped- (#                 ) - Jodifl`n`n[Alt+1] - Dropped - Jodifl`n`n[Alt+2]   - Change quantity - Jodifl`n`n[Alt+3] for damages & missing - Jodifl`n`n[Alt+4] for returns option - Jodifl`n`n[Alt+5] Dear Customer,`n`nThis email is to inform you that style (                      )  has been`n`n123`n`n[Alt+6]`nDear Customer,`nThis email is to inform you that style (                         ) has been`ndiscontinued due to a fabric shortage. We do apologize for this`ninconvenience this may have caused. If you have further questions`nor concerns, feel free to contact us and one  of our representative`nwill assist you.`nThank you,`nBest!`nJODIFL`n213-537-0578


return

ExitApp



!z::
SendInput, ( available_allocate_qty > 0 )  
return



!`::
Send, Dropped- ({#}                 ) - Jodifl{Enter}
return

!1::
Send,  - Dropped - Jodifl{Enter}
return


!2::
Send,   - Change quantity - Jodifl{Enter}
return

!3::
Send,  for damages & missing - Jodifl{Enter}
return


!4::
Send,  for returns option - Jodifl{Enter}
return



!5::
Send, Dear Customer,`n`nThis email is to inform you that style ({#}                      )  has been`n`n123{Enter}
return



!6::
Send, Dear Customer,`nThis email is to inform you that style ({#}                         ) has been`ndiscontinued due to a fabric shortage. We do apologize for this`ninconvenience this may have caused. If you have further questions`nor concerns, feel free to contact us and one  of our representative`nwill assist you.`nThank you,`nBest!`nJODIFL`n213-537-0578{Enter}
return



Esc::
ExitApp