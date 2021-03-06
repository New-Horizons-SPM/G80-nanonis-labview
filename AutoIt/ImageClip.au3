#include <ClipBoard.au3>
#include <GDIPlus.au3>

;===============================================================================
;
; Function Name:   _ImageToClip
; Description::    Copies all Image Files to ClipBoard
; Parameter(s):    $Path -> Path of image
; Requirement(s):  GDIPlus.au3
; Return Value(s): Success: 1
;                  Error: 0 and @error:
;                          1 -> Error in FileOpen
;                          2 -> Error when setting to Clipboard
; Author(s):
;
;===============================================================================
;
Func _ImageToClip($Path)
	_GDIPlus_Startup()
	Local $hImg = _GDIPlus_ImageLoadFromFile($Path)
	If $hImg = 0 Then Return SetError(1,0,0)
	Local $hBitmap = _GDIPlus_ImageCreateGDICompatibleHBITMAP($hImg)
    _GDIPlus_ImageDispose($hImg)
	_GDIPlus_Shutdown()
	Local $ret = _ClipBoard_SetHBITMAP($hBitmap)
	Return 1
EndFunc


Func _ImageToClipResize($Path,$x,$y)
	_GDIPlus_Startup()
	Local $hImg = _GDIPlus_ImageLoadFromFile($Path)
	If $hImg = 0 Then Return SetError(1,0,0)
	Local $hBitmap = _GDIPlus_ImageCreateGDICompatibleHBITMAPResize($hImg,$x,$y)
    _GDIPlus_ImageDispose($hImg)
	_GDIPlus_Shutdown()
	Local $ret = _ClipBoard_SetHBITMAP($hBitmap)
	Return 1
EndFunc

;===============================================================================
;
; Function Name:   _ClipBoard_SetHBITMAP
; Description::    Sets a HBITAMP as ClipBoardData
; Parameter(s):    $hBitmap -> Handle to HBITAMP from GDI32, NOT GDIPlus
; Requirement(s):  ClipBoard.au3
; Return Value(s): Success: 1 ; Error: 0 And @error = 1
; Author(s):       Prog@ndy
; Notes:           To use Images from GDIplus, convert them with _GDIPlus_ImageCreateGDICompatibleHBITMAP
;
;===============================================================================
;
Func _ClipBoard_SetHBITMAP($hBitmap,$Empty = 1)
	_ClipBoard_Open(0)
	If $Empty Then _ClipBoard_Empty()
	_ClipBoard_SetDataEx( $hBitmap, $CF_BITMAP)
	_ClipBoard_Close()
	If Not _ClipBoard_IsFormatAvailable($CF_BITMAP)  Then
		Return SetError(1,0,0)
	EndIf
EndFunc

;===============================================================================
;
; Function Name:   _GDIPlus_ImageCreateGDICompatibleHBITMAP
; Description::    Converts a GDIPlus-Image to GDI-combatible HBITMAP
; Parameter(s):    $hImg -> GDIplus Image object
; Requirement(s):  GDIPlus.au3
; Return Value(s): HBITMAP, compatible with ClipBoard
; Author(s):       Prog@ndy
;
;===============================================================================
;

Func _GDIPlus_ImageCreateGDICompatibleHBITMAP($hImg)
	Local $hBitmap2 = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hImg)
	Local $hBitmap = _WinAPI_CopyImage($hBitmap2,0,0,0,$LR_COPYDELETEORG+$LR_COPYRETURNORG)
	_WinAPI_DeleteObject($hBitmap2)
	Return $hBitmap
 EndFunc

 Func _GDIPlus_ImageCreateGDICompatibleHBITMAPResize($hImg,$x,$y)
	Local $hBitmap2 = _GDIPlus_ImageResize ( $hImg, $x, $y)
	Local $hBitmap3 = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hBitmap2)

	Local $hBitmap = _WinAPI_CopyImage($hBitmap3,0,0,0,$LR_COPYDELETEORG+$LR_COPYRETURNORG)
	_WinAPI_DeleteObject($hBitmap2)
	_WinAPI_DeleteObject($hBitmap3)

	Return $hBitmap
EndFunc

;===============================================================================
;
; Function Name:   _WinAPI_CopyImageLocal
; Description::    Copies an image, also makes GDIPlus-HBITMAP to GDI32-BITMAP
; Parameter(s):    $hImg -> HBITMAP Object, GDI or GDIPlus
; Requirement(s):  WinAPI.au3
; Return Value(s): Succes: Handle to new Bitmap, Error: 0
; Author(s):       Prog@ndy
;
;===============================================================================
;
Func _WinAPI_CopyImageLocal($hImg,$uType=0,$x=0,$y=0,$flags=0)
	Local $aResult

	$aResult = DllCall("User32.dll", "hwnd", "CopyImage", "hwnd", $hImg,"UINT",$uType,"int",$x,"int",$y,"UINT",$flags)
	_WinAPI_Check("_WinAPI_CopyImage", ($aResult[0] = 0), 0, True)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_CopyIcon

;===============================================================================
;
; Function Name:   _AutoItWinGetHandle
; Description::    Returns the Windowhandle of AutoIT-Window
; Parameter(s):    --
; Requirement(s):  --
; Return Value(s): Autoitwindow Handle
; Author(s):       Prog@ndy
;
;===============================================================================
;
Func _AutoItWinGetHandle()
	Local $oldTitle = AutoItWinGetTitle()
	Local $x = Random(1248578,1249780)
	AutoItWinSetTitle("qwrzu"&$x)
	Local $x = WinGetHandle("qwrzu"&$x)
	AutoItWinSetTitle($oldTitle)
	Return $x
EndFunc

; #####################################################################
; #####################################################################
; 	EXAMPLE
; #####################################################################

;~ 	$Path = "D:\Dokumente\Dateien von Andreas\Eigene Bilder\Avatar-MR-X.gif"
;~ 	_ImageToClip($Path)
;~ 	$paint = Run("mspaint")
;~ 	$oldMatch = Opt("WinTitleMatchMode",2)
;~ 	WinWait("- Paint")
;~ 	Opt("WinTitleMatchMode",$oldMatch)
;~ 	If WinExists("Untitled - Paint") Then
;~ 		WinWaitActive("Untitled - Paint")
;~ 	Else
;~ 		WinWaitActive("Unbenannt - Paint")
;~ 	EndIf
;~ 	Send("^v")
; #####################################################################
; #####################################################################