#include-once
#Obfuscator_Ignore_Funcs=__EnumChildWinProc
#include <WinAPI.au3>
; ===============================================================================================================================
; <_EnumChildWindows.au3>
;
;	Function to enumerate child windows (which usually consist of just controls) of a window.
;		Its flexibility allows it to search for only windows/controls of a given class, handle, or title (w/ 3 match-modes)
;
; Functions:
;	_EnumChildWindows()
;
; INTERNAL Functions:
;	__EnumChildWinProc
;
; See also:
;	<_WinWaitEx.au3>	; looks for window matching process ID (used to be part of the 'Test' example script)
;	<_EnumProcessWindows.au3>
;	<_WinGetAltTabWinList.au3>
;	<TestEnumChildWindows.au3>
;
; Author: Ascend4nt
; ===============================================================================================================================

; ===================================================================================================================
;	GLOBAL __EnumChildWinProc() CALLBACK VARIABLES
; ===================================================================================================================

Global $_iECArray2ndDim,$_aECWinList,$_vECControl,$_sECTitle,$_sECClass,$_bECGetTitle,$_iECTitleMatchMode

; ===================================================================================================================
; Func _EnumChildWindows($vWnd,$vControl=0,$sTitle=0,$sClass=0,$iTitleMatchMode=0,$bAllIterationCounts=True,$bGetTitle=True)
;
; Function to Enumerate child windows using API call (and callback routine)
;  Note: usually only 1 of the 3 parameters following $vWnd needs to be set. Also:
;	- If $vControl is set to the Control Handle (NOT ID) of a Control, ONLY that item is searched for
;	  (it will return the iteration and class, useful for identification in the future using [CLASS:name; INSTANCE:##]
;	- If $sTitle is set, it will search for controls that return something on WinGetTitle() that match
;	- If $sClass is set, it will only search for items of a given class. This can be combined with $sTitle, but
;	  it's not recommended
;
; NOTE: Iteration numbering is dependent on what is passed.  See parameters.
;
; $vWnd = Handle to window, or Title (see special definitions documentation for WinGetHandle())
; $vControl = (optional) Handle to control to find.  This will return only one match, but sets the iteration count and the rest
; $sTitle = (optional) Title to search for. This can be direct match, a substring, or a PCRE (see $iTitleMatchMode)
;	NOTES: setting this to "" will make the routine search for windows *without* titles. Set to non-string (default) to match any
;	  Also, no iteration count is kept for this unless $sClass is also set
; $sClass = (optional) Classname to look for.  This WILL produce iteration counts
; $iTitleMatchMode = mode of matching title:
;	0 = default match string (not case sensitive)
;	1 = match substring
;	2 = match regular expression
; $bAllIterationCounts = If True (default), get iteration counts for all classnames.
;	This will fail of course if $vControl, $sTitle, or $sClass is passed
; $bGetTitle = If True, grabs the title/text of the control/child window
;
; Returns:
;	Success: An array set up as follows (also @extended=#of colummns)
;		[0][0] = count of items (0 means nothing found), set up as follows:
;
;		[$i][0] = Child Window/Control Handle
;		[$i][1] = Class name
;		[$i][2] = Control ID - This is good to have for sending/posting messages.
;				  NOTE: This will be 0 for some controls, and -1 for others when they don't have an official ID #
;		[$i][3] = Instance count (if applicable - see above)
;	  Also, *IF* $bGetTitle=True, then:
;		[$i][4] = Title/Text from control
;
;	Failure: "" is returned with @error set:
;		@error = 2 = error setting up callback proc (@extended= error # from DllCallbackRegister)
;		@error = 3 = Enumeration API call failed (@extended = error returned from DLLCall)
;
; Author: Ascend4nt
; ===================================================================================================================

Func _EnumChildWindows($vWnd,$vControl=0,$sTitle=0,$sClass=0,$iTitleMatchMode=0,$bAllIterationCounts=True,$bGetTitle=True)
	Local $hCBReg,$aWinArray,$iErr,$i,$iIn,$iIteration,$sTmpClass

	If Not IsHWnd($vWnd) Then $vWnd=WinGetHandle($vWnd)
	; I don't know that there are empty-string classes, so for now, we'll treat empty strings as do-not-search
	If $sClass="" Then $sClass=0

	$_iECArray2ndDim=4
	If $bGetTitle Then $_iECArray2ndDim+=1

	; Setup the winlist array
	Dim $_aECWinList[10+1][$_iECArray2ndDim]
	$_aECWinList[0][0]=0
	; Current-array-size count (used in ReDim'ing array inside the callback)
	$_aECWinList[0][1]=10

	; Control Handle received? We will only find one match, so let's grab everything except the iteration count now
	If IsHWnd($vControl) Then
		; Grab classname now
		If Not IsString($sClass) Then $sClass=_WinAPI_GetClassName($vControl)
		$_aECWinList[1][0]=$vControl
		$_aECWinList[1][1]=$sClass	; no need to repeat this
		; Set the Control ID # (GWL_ID = -12)
		$_aECWinList[1][2]=_WinAPI_GetWindowLong($vControl,-12)
		$_aECWinList[1][3]=0	; Set the instance count (if searching for a specific handle)
		; Set the title/text of control/window (ControlGetText is an option but won't work for non-control child windows..)
		If $bGetTitle Then $_aECWinList[1][4]=WinGetTitle($vControl)	; or ControlGetText($vWnd,"",$vControl)
	EndIf

	; Set the required callback parameters (except $_iECArray2ndDim, set above due to initialization of array)
	$_bECGetTitle=$bGetTitle
	$_iECTitleMatchMode=$iTitleMatchMode
	$_sECClass=$sClass
	$_sECTitle=$sTitle
	$_vECControl=$vControl
	$_iECCtrlIteration=0

	$hCBReg=DllCallbackRegister("__EnumChildWinProc","int","hwnd;lparam")
	If $hCBReg=0 Then Return SetError(2,0,"")

;~ 	ConsoleWrite("On entry, $vWnd:"&$vWnd&" $vControl:"&$vControl&" $sTitle: '"&$sTitle&"' $sClass: '"&$sClass&"'"&@CRLF)

	; BOOL EnumChildWindows(_in HWND hWndParent,_in WNDENUMPROC lpEnumFunc,_in LPARAM lParam);
	$aRet=DllCall("user32.dll","int","EnumChildWindows","hwnd",$vWnd,"ptr",DllCallbackGetPtr($hCBReg),"lparam",101)
	$iErr=@error

	DllCallbackFree($hCBReg)

	If $iErr Then Return SetError(3,$iErr,"")

	ReDim $_aECWinList[$_aECWinList[0][0]+1][$_iECArray2ndDim]
	$aWinArray=$_aECWinList

	$aWinArray[0][1]=""
	$_aECWinList=""

	; Set all iterations? *Only if $bAllIterationCounts=True, AND *ALL* child windows are enumerated.*
	If $bAllIterationCounts And $aWinArray[0][0] And $sClass="" And Not IsString($sTitle) And Not IsHWnd($vControl) Then
		For $i=0 To $aWinArray[0][0]
			; Found an empty iteration count?
			If Not $aWinArray[$i][3] Then
				$iIteration=1
				$aWinArray[$i][3]=1
				$sTmpClass=$aWinArray[$i][1]
				For $iIn=$i+1 To $aWinArray[0][0]
					If $aWinArray[$iIn][1]=$sTmpClass Then
						$iIteration+=1
						$aWinArray[$iIn][3]=$iIteration
					EndIf
				Next
			EndIf
		Next
	EndIf
;~ 	ConsoleWrite("Child Window array size on exit:"&$aWinArray[0][0]&@CRLF)
	Return SetError(0,$_iECArray2ndDim,$aWinArray)
EndFunc


; ===================================================================================================================
; Func __EnumChildWinProc($hWnd,$lParam)
;
; Do *NOT* call this directly!!
;
; MSDN:  BOOL CALLBACK EnumChildProc(_in HWND hwnd,_in LPARAM lParam);
;
; Author: Ascend4nt
; ===================================================================================================================

Func __EnumChildWinProc($hWnd,$lParam)
	Local $sClassname,$sTitle

	$sClassname=_WinAPI_GetClassName($hWnd)
;~ 	ConsoleWrite("hWnd found:"&$hWnd&" Classname:"&$sClassname&" Title:"&WinGetTitle($hWnd)&@CRLF)

	; No match for our classname? Keep searching
	If IsString($_sECClass) And $_sECClass<>$sClassname Then Return True

	; Classname match AND we are searching for a handle?
	If IsHWnd($_vECControl) Then
		; Keep this at 1 always
		$_aECWinList[0][0]=1
		$_aECWinList[1][3]+=1	; increment iteration/instance count
		; Have we found it?!
		If $hWnd=$_vECControl Then Return False
		Return True
	EndIf

	If $_bECGetTitle Then $sTitle=WinGetTitle($hWnd)	; basically same as: ControlGetText($_hECMainWnd,"",$vControl)

	; Either searching only by title, or by title and classname.. check to see if there's a match. If not, skip!

	If IsString($_sECTitle) Then
		; Prevent duplicate title/text reads
		If Not $_bECGetTitle Then $sTitle=WinGetTitle($hWnd)
		; If no match in any mode, then continue search (Return True)
		Switch $_iECTitleMatchMode
			Case 0
				If $_sECTitle<>$sTitle Then Return True
			Case 1
				If StringInStr($sTitle,$_sECTitle)=0 Then Return True
			Case Else
				If StringRegExp($sTitle,$_sECTitle)=0 Then Return True
		EndSwitch
	EndIf

	; We are either grabbing all child windows, those matching a title, or a given classname..

	$_aECWinList[0][0]+=1
	If $_aECWinList[0][0]>$_aECWinList[0][1] Then
		$_aECWinList[0][1]+=10
		ReDim $_aECWinList[$_aECWinList[0][1]+1][$_iECArray2ndDim]
;~ 		ConsoleWrite("Resized array to "&$_aECWinList[0][1]&"+1 (count at bottom) elements, currently AT element #"&$_aECWinList[0][0]&@CRLF)
	EndIf
	$_aECWinList[$_aECWinList[0][0]][0]=$hWnd
	$_aECWinList[$_aECWinList[0][0]][1]=$sClassname
	; Set the Control ID # (GWL_ID = -12)
	$_aECWinList[$_aECWinList[0][0]][2]=_WinAPI_GetWindowLong($hWnd,-12)

	; Keep iteration counts if looking at a specific classname
	If IsString($_sECClass) Then
		If $_aECWinList[0][0]>1 Then
			$_aECWinList[$_aECWinList[0][0]][3]=$_aECWinList[$_aECWinList[0][0]-1][3]+1
		Else
			$_aECWinList[$_aECWinList[0][0]][3]=1
		EndIf
	EndIf
	If $_bECGetTitle Then $_aECWinList[$_aECWinList[0][0]][4]=$sTitle
	Return True
EndFunc

#cs
	; Alternate form of getting Control ID, matches what _WinAPI_GetWindowLong($hWnd,-12) gets:
	;int GetDlgCtrlID(_in HWND hwndCtl);
	$aRet=DllCall("user32.dll","int","GetDlgCtrlID","hwnd",$hWnd)
	ConsoleWrite("WindowLong ID:"&$_aECWinList[$_aECWinList[0][0]][2]&" GetDlgCtrlID:"&$aRet[0]&@CRLF)
#ce