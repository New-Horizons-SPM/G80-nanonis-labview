

#include <GUIConstantsEx.au3>
#include <GDIPlus.au3>

Func drawCross($x,$y,$file,$outfile)
_GDIPlus_Startup()
Global $hImage=  _GDIPlus_ImageLoadFromFile($file)
If Not $hImage Then
    _GDIPlus_Shutdown()
    Exit
EndIf
Global $iW = _GDIPlus_ImageGetWidth($hImage), $iH = _GDIPlus_ImageGetHeight($hImage)
Global $hGUI = GUICreate("GDI+ Test", $iW, $iH) ;no check when image is larger than screen!
;~ GUISetState()
Global $hGfx = _GDIPlus_GraphicsCreateFromHWND($hGUI) ;only for display purposes
Global $hBitmap = _GDIPlus_BitmapCreateFromGraphics($iW, $iH, $hGfx)
Global $hImgContext = _GDIPlus_ImageGetGraphicsContext($hBitmap) ;this is needed to copy the image to the bitmap and draw on it
_GDIPlus_GraphicsDrawImageRect($hImgContext, $hImage, 0, 0, $iW, $iH) ;copy loaded image to bitmap (not display in GUI yet)
Global $hBrush = _GDIPlus_BrushCreateSolid(0xFF007EEF) ;create a brush AARRGGBB
_GDIPlus_GraphicsFillRect($hImgContext, $x - 1, 0, 2, $iH, $hBrush) ;draw filled rectangle of the image vertically
_GDIPlus_GraphicsFillRect($hImgContext, 0, $y - 1, $iW, 2, $hBrush) ;draw filled rectangle of the image horizontally

_GDIPlus_GraphicsDrawImageRect($hGfx, $hBitmap, 0, 0, $iW, $iH) ;copy modified image to graphic handle -> display it in the GUI

_GDIPlus_ImageSaveToFile($hBitmap, $outfile)

;clean up resources
_GDIPlus_BrushDispose($hBrush)
_GDIPlus_GraphicsDispose($hImgContext)
_GDIPlus_GraphicsDispose($hGfx)
_GDIPlus_BitmapDispose($hBitmap)
_GDIPlus_ImageDispose($hImage)
_GDIPlus_Shutdown()


   _GDIPlus_Shutdown()
EndFunc