if WinExists("Sponsored session") Then
;~    WinClose("Sponsored session")
WinActivate("Sponsored session")
Sleep(300)
$pos=WinGetPos("Sponsored session")
MouseClick("left",$pos[0]+430,$pos[1]+142,5)
EndIf