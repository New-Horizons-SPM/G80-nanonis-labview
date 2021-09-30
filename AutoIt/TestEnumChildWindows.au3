#include <_EnumChildWindows.au3>
; ===============================================================================================================================
; <TestEnumChildWindows.au3>
;
;	Test for _EnumChildWindows UDF.
;
; Author: Ascend4nt
; ===============================================================================================================================

; ===================================================================================================================
;	TEST
; ===================================================================================================================

#include <Array.au3>

Func closeWSXM()

Local $hWnd="",$hControl=0,$sTitle=0,$sClass=0,$aEnumList
Local $iCalcPID,$hNumber1Button=-1,$hNumber4Button=-1,$hPlusButton=-1,$hEqualButton=-1
$hWnd=WinGetHandle("[CLASS:AfxMDIFrame80s]")

WinActivate($hWnd)	; this seems to be a better alternative, the window seems fully created after this is called in my tests
WinWaitActive($hWnd)


;~ $aEnumList=_EnumChildWindows($hWnd,$hControl,$sTitle,$sClass)	;,2) for RegExp Title
$aEnumList=_EnumChildWindows($hWnd,$hControl,$sTitle,$sClass)

Local $first=True
For $i=1 to $aEnumList[0][0]
   If StringInStr($aEnumList[$i][4], "image") <> 0 Then

   If StringInStr($aEnumList[$i][4],"Image Z (m).sxm") ==0 Then
   WinClose($aEnumList[$i][0])
   ElseIf $first Then
   $first = False
   Else
   WinClose($aEnumList[$i][0])
   $first = True
   EndIf
   EndIf
Next

   EndFunc

closeWSXM()
