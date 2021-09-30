#cs ----------------------------------------------------------------------------

    AutoIt Version: 3.2.13.7 (beta)
    Author:      dbzfanatic

    Script Function:
    Task Scheduler UDF (Task.au3)

#ce ----------------------------------------------------------------------------

Global $sRet

; #FUNCTION#;===============================================================================
;
; Name...........: _TaskSchedule
; Description ...: Adds a scheduled task.
; Syntax.........: _TaskSchedule($sDay, $sTime, $hProgram, $sName, $iID, $bInteractive, $iOccurrence)
; Parameters ....: $sDay - The day(s) to run the task. (ex. M, T, Th; 1, 2, 3, 15, 22)
;                   $sTime - The time to run the program.
;                   $hProgram - Path to the program/batch file to run.
;                   $sName - [Optional] The name of the computer on which to schedule the task.
;                   $iID - [Optional] The id to give the specified task.
;                   $bInteractive - [Optional] Boolean value to allow interaction with the desktop
;                   $iOccurrence - [Optional] Determines how often the program/batch file is run.
;~                      |1 - Every day
;~                      |2 - Once
; Return values .: Success - Deletes the specified task.
;                 Failure - Sets @Error:
;                 |1 - Invalid $sName
;                 |2 - Invalid $sTime
;                 |3 - Invalid $iID
;                 |4 - Invalid $bInteractive
;                 |5, 7 - Invalid $iOccurrence
;                 |6 - Invalid $hProgram
;                 |16 - Command failed
; Author ........: Michael Duane LaBruyere (dbzfanatic)
; Modified.......:
; Remarks .......:
; Related .......: _TaskDelete(), _TaskGetSchedule()
; Link ..........;
; Example .......; Yes
;
;;==========================================================================================

Func _TaskSchedule($sDay, $sTime, $hProgram, $sName = "", $iID = 1, $bInteractive = True, $iOccurrence = 1)

    Local $sInteract, $sOccur, $hRun

    If Not IsString($sName) Then
        Return SetError(1)
    EndIf

    If Not IsString($sTime) Then
        Return SetError(2)
    EndIf

    If Not IsInt($iID) Then
        Return SetError(3)
    EndIf

    If Not IsBool($bInteractive) Then
        Return SetError(4)
    EndIf

    If Not IsInt($iOccurrence) Then
        Return SetError(5)
    EndIf

    If $hProgram = "" Then
        Return SetError(6)
    EndIf

    If $bInteractive = True Then
        $sInteract = "/interactive"
    Else
        $sInteract = ""
    EndIf

    If $iOccurrence = 1 Then
        $sOccur = "/every:" & $sDay
    ElseIf $iOccurrence = 2 Then
        $sOccur = "/next:" & $sDay
    Else
        Return SetError(7)
    EndIf

    $hRun = Run("cmd.exe /c at " & $sName & " " & $iID & " " & " " & $sDay & " " & $sInteract & " " & $sOccur & " " & $hProgram, @SystemDir, @SW_HIDE)

    While ProcessExists("cmd.exe")
        $sRet = StdoutRead($hRun, True)
    WEnd
    If $hRun = 0 Then
        Return SetError(16)
    EndIf
    Return $sRet

EndFunc;==>_TaskSchedule

; #FUNCTION#;===============================================================================
;
; Name...........: _TaskGetScheduled
; Description ...: Returns a list of scheduled tasks.
; Syntax.........: _TaskGetScheduled()
; Parameters ....:
; Return values .: Success - Returns the list of scheduled tasks.
;                 Failure - Sets @Error:
;                 |1 - Command Failed
; Author ........: Michael Duane LaBruyere (dbzfanatic)
; Modified.......:
; Remarks .......:
; Related .......: _TaskSchedule(), _TaskDelete()
; Link ..........;
; Example .......; Yes
;
;;==========================================================================================
Func _TaskGetScheduled()

    Local $hAt

    $hAt = Run("cmd.exe /c at", @SystemDir, @SW_HIDE, 8)
    While ProcessExists("cmd.exe")
        $sRet = StdoutRead($hAt, True)
    WEnd
    If $hAt = 0 Then
        Return SetError(1)
    EndIf
    Return $sRet
EndFunc;==>_TaskGetScheduled

; #FUNCTION#;===============================================================================
;
; Name...........: _TaskDelete
; Description ...: Deletes a scheduled task.
; Syntax.........: _TaskDelete($sName, $iID, $iDelete)
; Parameters ....: $sName - The name of the computer to execute the deletion on.
;                   $iID - [Optional] The id of the task to delete. If blank will delete the first scheduled task.
;                   $iDelete - [Optional] Integer value to determine deletion type. If 0 will delete all tasks, 1 will delete on the specified task.
; Return values .: Success - Deletes the specified task.
;                 Failure - Sets @Error:
;                 |1 - Invalid $iDelete
;                 |2 - Invalid $iID
;                 |3 - Invalid $sName
;                 |4 - Command Failed
; Author ........: Michael Duane LaBruyere (dbzfanatic)
; Modified.......:
; Remarks .......:
; Related .......: _TaskSchedule(), _TaskGetSchedule()
; Link ..........;
; Example .......; Yes
;
;;==========================================================================================
Func _TaskDelete($sName, $iID = 1, $iDelete = 1)

    Local $hDel, $sDelete

    If Not IsInt($iDelete) Then
        Return SetError(1)
    EndIf

    If Not IsInt($iID) Then
        Return SetError(2)
    EndIf

    If Not IsString($sName) Then
        Return SetError(3)
    EndIf

    If $iDelete = 0 Then
        $sDelete = "/sDelete /yes"
    ElseIf $iDelete = 1 Then
        $sDelete = "/sDelete " & $iID & " /yes"
    EndIf

    $hDel = Run("cmd.exe /c at " & $sName & " " & $sDelete, @SystemDir, @SW_HIDE, 8)
    While ProcessExists("cmd.exe")
        $sRet = StdoutRead($hDel, True)
    WEnd
    If $hDel = 0 Then
        Return SetError(4)
    EndIf
    Return $sRet
EndFunc;==>_TaskDelete