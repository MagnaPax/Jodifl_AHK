#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;==================================================

#IfWinActive, ahk_class IEFrame ;internet explorer
^y:: ;internet explorer - get/set zoom
WinGet, hWnd, ID, A
oWB := WBGet("ahk_id " hWnd)
vLogicalXDPI := oWB.document.parentWindow.screen.logicalXDPI
vDeviceXDPI := oWB.document.parentWindow.screen.deviceXDPI
vZoom := Round(vDeviceXDPI / vLogicalXDPI * 100)

;JEE_InputBoxVPD(vZoom2, "zoom %:", vZoom)
InputBox, vZoom2, , % "zoom %:", , , , , , , , %vZoom%
;if (vZoom2 = "n")
;vZoom2 := 100
;if vZoom2 in 0,1,2,3,4,5,6,7,8,9
;vZoom2 := "1" vZoom2 "0"
if ErrorLevel OR (vZoom2 = vZoom)
{
oWB := ""
Return
}

oZoomNew := ComVar()
oWB.ExecWB(63, 2, vZoom2, oZoomNew.ref)
oWB := ""
MsgBox % oZoomNew[]
Return
#IfWinActive

;==================================================

; ComVar: Creates an object which can be used to pass a value ByRef.
;   ComVar[] retrieves the value.
;   ComVar[] := Val sets the value.
;   ComVar.ref retrieves a ByRef object for passing to a COM function.
ComVar(Type:=0xC)
{
    static base := { __Get: "ComVarGet", __Set: "ComVarSet", __Delete: "ComVarDel" }
    ; Create an array of 1 VARIANT.  This method allows built-in code to take
    ; care of all conversions between VARIANT and AutoHotkey internal types.
    arr := ComObjArray(Type, 1)
    ; Lock the array and retrieve a pointer to the VARIANT.
    DllCall("oleaut32\SafeArrayAccessData", "ptr", ComObjValue(arr), "ptr*", arr_data)
    ; Store the array and an object which can be used to pass the VARIANT ByRef.
    return { ref: ComObject(0x4000|Type, arr_data), _: arr, base: base }
}
ComVarGet(cv, p*) { ; Called when script accesses an unknown field.
    if !p.Length  ; No name/parameters, i.e. cv[]
        return cv._[0]
}
ComVarSet(cv, v, p*) { ; Called when script sets an unknown field.
    if !p.Length  ; No name/parameters, i.e. cv[]:=v
        return cv._[0] := v
}
ComVarDel(cv) { ; Called when the object is being freed.
    ; This must be done to allow the internal array to be freed.
    DllCall("oleaut32\SafeArrayUnaccessData", "ptr", ComObjValue(cv._))
}

;==================================================