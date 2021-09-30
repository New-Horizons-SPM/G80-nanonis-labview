
Func wait4anyKey()
   $sKeyboardState = _WinAPI_GetKeyboardState(1)
While _WinAPI_GetKeyboardState(1) = $sKeyboardState
    Sleep(100)
 WEnd
 EndFunc



; #FUNCTION# ====================================================================================================================
; Name...........:  _WinAPI_GetKeyboardState
; Description ...:  Returns the status of the 256 virtual keys
; Syntax.........:  _WinAPI_GetKeyboardState($iFlag=0)
; Parameters ....:  $iFlag   - Return Type
;                   0 Returns an array[256]
;                   1 Returns a string
; Return values .:  Success  - Array[256] or String containing status of 256 virtual keys
;                   Failure  - False
; Author ........:  Eukalyptus
; Modified.......:
; Remarks .......:  If the high-order bit is 1, the key is down; otherwise, it is up.
;                   If the key is a toggle key, for example CAPS LOCK, then the low-order bit is 1
;                   when the key is toggled and is 0 if the key is untoggled
; Related .......:  _IsPressed
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _WinAPI_GetKeyboardState($iFlag = 0)
    Local $aDllRet, $lpKeyState = DllStructCreate("byte[256]")
    $aDllRet = DllCall("User32.dll", "int", "GetKeyboardState", "ptr", DllStructGetPtr($lpKeyState))
    If @error Then Return SetError(@error, 0, 0)
    If $aDllRet[0] = 0 Then
        Return SetError(1, 0, 0)
    Else
        Switch $iFlag
            Case 0
                Local $aReturn[256]
                For $i = 1 To 256
                    $aReturn[$i - 1] = DllStructGetData($lpKeyState, 1, $i)
                Next
                Return $aReturn
            Case Else
                Return DllStructGetData($lpKeyState, 1)
        EndSwitch
    EndIf
EndFunc   ;==>_WinAPI_GetKeyboardState