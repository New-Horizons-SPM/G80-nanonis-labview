


#include <_FileDragDrop.au3>

#include <_EnumChildWindows.au3>




Func closeWSXM()

Local $hWnd="",$hControl=0,$sTitle=0,$sClass=0,$aEnumList
Local $iCalcPID,$hNumber1Button=-1,$hNumber4Button=-1,$hPlusButton=-1,$hEqualButton=-1
$hWnd=WinGetHandle("[CLASS:AfxMDIFrame80s]")

WinActivate($hWnd)	; this seems to be a better alternative, the window seems fully created after this is called in my tests
WinWaitActive($hWnd)


;~ $aEnumList=_EnumChildWindows($hWnd,$hControl,$sTitle,$sClass)	;,2) for RegExp Title
$aEnumList=_EnumChildWindows($hWnd,$hControl,$sTitle,$sClass)

Local $first=""

ProgressOn("Closing channel windows in WxSM","")
For $i=$aEnumList[0][0] to 1 Step -1
   If StringInStr($aEnumList[$i][4], "sxm") <> 0 Then

   If StringInStr($aEnumList[$i][4],"Image Z (m).sxm") ==0 Then
   WinClose($aEnumList[$i][0])
   ElseIf $first <> $aEnumList[$i][4]  Then
   $first = $aEnumList[$i][4]
   Else
   WinClose($aEnumList[$i][0])
   EndIf
   EndIf
   $per=Round(100*($aEnumList[0][0] -$i)/$aEnumList[0][0] )
   ProgressSet($per, $per & " %")
Next
ProgressOff()
   EndFunc


Func dragdropFile()
$temp=clipget()
Send("^c")
$filename=clipget()
if StringLen($filename)>1 Then
$filepath="X:\STM Data\"
$session= "20"&StringLeft($filename,2)&StringMid($filename,3,2)&'\'
$hNPPlus=WinGetHandle("[CLASS:AfxMDIFrame80s]")
$aNPPPos=WinGetPos($hNPPlus)
EndIf
ConsoleWrite($filepath&$session&$filename&@CR)
clipput($filepath&$session&$filename)

$iRet=_FileDragDrop($hNPPlus,$filepath&$session&$filename&"|",0,0

,'|',True)


EndFunc

HotKeySet("{F6}","dragdropFile")
HotKeySet("{F7}","closeWSXM")

while 1
   Sleep(200)
   WEnd

;~ closeWSXM()
;~ Local $hWnd="",$hControl=0,$sTitle=0,$sClass=0,$aEnumList
;~ Local $iCalcPID,$hNumber1Button=-1,$hNumber4Button=-1,$hPlusButton=-1,$hEqualButton=-1
;~ $hWnd=WinGetHandle("[CLASS:AfxMDIFrame80s]")

;~ WinActivate($hWnd)	; this seems to be a better alternative, the window seems fully created after this is called in my tests
;~ WinWaitActive($hWnd)


;~ $aEnumList=_EnumChildWindows($hWnd,$hControl,$sTitle,$sClass)	;,2) for RegExp Title
;~ $aEnumList=_EnumChildWindows($hWnd,$hControl,$sTitle,$sClass)

;~ ConsoleWrite($aEnumList)

;~ Local $first=""

;~ ProgressOn("Closing channel windows in WxSM","")
;~ For $i=$aEnumList[0][0] to 1 Step -1
;~    If StringInStr($aEnumList[$i][4], "image") <> 0 Then

;~    If StringInStr($aEnumList[$i][4],"Image Z (m).sxm") ==0 Then
;~    WinClose($aEnumList[$i][0])
;~    ElseIf $first <> $aEnumList[$i][4]  Then
;~    $first = $aEnumList[$i][4]
;~    Else
;~    WinClose($aEnumList[$i][0])
;~    EndIf
;~    EndIf
;~    $per=Round(100*($aEnumList[0][0] -$i)/$aEnumList[0][0] )
;~    ProgressSet($per, $per & " %")
;~ Next
;~ ProgressOff()
