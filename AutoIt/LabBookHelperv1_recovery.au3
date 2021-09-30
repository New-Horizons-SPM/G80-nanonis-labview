#include <Array.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include "ImageClip.au3"
#include <_EnumChildWindows.au3>

Global $DirPath= ""
Global $queue=0;

Func getSessionPath()
   Return GUICtrlRead($txtGRID)
   EndFunc

Func getSPMtemps()
   ShellExecuteWait("D:\Scripts\VIs\builds\WriteTempertures.exe")
   $hFile = Fileopen("D:\Scripts\temp\temperatures.csv",0)
   $tempCryo = FileReadLine ($hFile,1)
   $tempSTM = FileReadLine ($hFile,2)
   FileClose($hFile)
   Return String(Round(number($tempSTM),2))
   EndFunc

Func GetLatestFile($DirPath, $FileSpec)

; Shows the filenames of all files in the current directory.
$search = FileFindFirstFile($DirPath & "\" & $FileSpec)

$LatestFile = ""
$LatestTime = 0

; Check if the search was successful
If $search = -1 Then
    MsgBox(0, "Error", "No files/directories matched the search pattern")
Else
	While 1
		$File = FileFindNextFile($search)
		if @error then ExitLoop

		if StringInStr(FileGetAttrib($DirPath & "\" & $File), "D") > 0 Then
			; Skip directories
		Else
			$FileTime = FileGetTime($DirPath & "\" & $File,0,1)
			if $FileTime > $LatestTime Then
				$LatestFile = $File
				$LatestTime = $FileTime
			EndIf
		 EndIf

	WEnd
	Return $LatestFile
 EndIf
 EndFunc

Func GetLatestDir($DirPath)

; Shows the filenames of all files in the current directory.
$search = FileFindFirstFile($DirPath & "\*")

$LatestFile = ""
$LatestTime = 0

; Check if the search was successful
If $search = -1 Then
    MsgBox(0, "Error", "No files/directories matched the search pattern")
Else
	While 1
		$File = FileFindNextFile($search)
		if @error then ExitLoop

		if StringInStr(FileGetAttrib($DirPath & "\" & $File), "D") > 0 Then
			$FileTime = FileGetTime($DirPath & "\" & $File,0,1)
			if $FileTime > $LatestTime Then
				$LatestFile = $File
				$LatestTime = $FileTime
			EndIf
		 EndIf

	WEnd
	Return $LatestFile
 EndIf
 EndFunc

Func gotoFollowMe()
    WinActivate("Scan Control")
	Sleep(200)
;~    WinWaitActive("Scan Control")
   $pos=WinGetPos("Scan Control")
   MouseClick("",$pos[0]+145,$pos[1]+50,1,0)
   Sleep(200)
   EndFunc

Func gotoScanning()
   WinActivate("Scan Control")
   	Sleep(200)
;~    WinWaitActive("Scan Control")
   $pos=WinGetPos("Scan Control")
   MouseClick("",$pos[0]+45,$pos[1]+50,1,0)
   Sleep(400)
   EndFunc

Func grabImage($size)
   WinActivate("Scan Control")
;~    WinWaitActive("Scan Control")
	Sleep(200)
   $pos=WinGetPos("Scan Control")
   MouseClick("",$pos[0]+928+25,$pos[1]+88,1,0)
   Sleep(600)

    _ScreenCapture_Capture ( @MyDocumentsDir & "\temp.jpg" , $pos[0]+282  , $pos[1]+109 , $pos[0]+1003  , $pos[1]+828 , False )
    _ImageToClipResize(@MyDocumentsDir & "\temp.jpg",$size,$size)
   EndFunc

Func grabImageCursor($size)
   WinActivate("Scan Control")
;~    WinWaitActive("Scan Control")
	Sleep(200)
   $pos=WinGetPos("Scan Control")
   MouseClick("",$pos[0]+928+25,$pos[1]+88,1,0)
   	Sleep(200)
   MouseClick("",$pos[0]+145,$pos[1]+50,1,0)
   $posTip=PixelSearch($pos[0]+282  , $pos[1]+109 , $pos[0]+1003  , $pos[1]+828 , 0x007EEF )
   MouseMove($posTip[0]+1,$posTip[1]+4,1)
   Sleep(100)
    _ScreenCapture_Capture ( @MyDocumentsDir & "\temp.jpg" , $pos[0]+282  , $pos[1]+109 , $pos[0]+1003  , $pos[1]+828 , True )
    _ImageToClipResize(@MyDocumentsDir & "\temp.jpg",$size,$size)
   EndFunc


Func getTemperture()
   $filehwd=FileOpen("D:\Scripts\temp\temperatures.csv")
   $temp=Round(Number(FileReadLine($filehwd,2)),2)
   FileClose($filehwd)
   return $temp
   EndFunc

Func grabSTS()
     WinActivate("Bias Spectroscopy")
;~    WinWaitActive("Bias Spectroscopy")
	Sleep(200)
   $pos=WinGetPos("Bias Spectroscopy")
    _ScreenCapture_Capture ( @MyDocumentsDir & "\temp.jpg" , $pos[0]+20  , $pos[1]+450 , $pos[0]+610  , $pos[1]+730 , False )
    _ImageToClipResize(@MyDocumentsDir & "\temp.jpg",316,150)

   EndFunc


Func newTableLine()
   Send("^{END}")
   Sleep(100)
   Send("{APPSKEY}")
   Sleep(100)
   Send("a")
   Sleep(100)
   Send("b")
   EndFunc


Func writeTopo2Onenote()
$DirPath=getSessionPath()
;~ $STMtemp= getSPMtemps()
grabImage(250)
Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)
;~ Send("Filename")
;~ Send("{TAB}")
;~ Send("Image")
;~ Send("{TAB}")
;~ Send("Bias")
;~ Send("{TAB}")
;~ Send("Current")
;~ Send("{TAB}")
;~ Send("Comments")
;~ Send("{ENTER}")

;Find the last File
$FileSpec= "*.sxm"
Local $lastestFile = ClipGet()
;open and reading the file
$hFile = Fileopen($DirPath&"\"&$lastestFile,512)
;~ Local $keys[]= [":BIAS:",":Z-CONTROLLER:",":COMMENT:"]
;~ $i=1
;~ While 1
;~ $line = FileReadLine ($hFile,$i)
;~ For $str In $keys
;~ if $line= $str  Then
;~
;~ EndIf
;~ Next

;~ $i=$i+1
;~ WEnd
$SpeedStr=FileReadLine ($hFile,18)
$aSpeed =Stringsplit ($SpeedStr, " ",0)
$SizeStr = FileReadLine ($hFile,20)
$aSize =Stringsplit ($SizeStr, " ",0)
$Size=Round(Number($aSize[12])*1E9,2)&"x"& Round(Number($aSize[23])*1E9,2)&" nm2"
$BiasStr = FileReadLine ($hFile,28)
$Bias= Number(StringReplace($BiasStr, " ", ""))*1000
$CurrentStr = FileReadLine ($hFile,31)
$Speed=Round((Number($aSize[12])*1E9) /Number($aSpeed[14]),2)&" nm/s"
$Current= Stringsplit ($CurrentStr, Chr (9), 2)[3]
$CurrentnA=Number($Current)*1E9

FileClose($hFile)
newTableLine()
Send("TOPO")
Send("{ENTER}")
Send("{ENTER}")
Send ($lastestFile)
highlightOneNoteBlue()
Send("{TAB}")
send("^v")
Send($Size)
Send("{TAB}")
Send ($Bias&" mV")
Send("{TAB}")
Send ($CurrentnA&" nA")
Send("{TAB}")
;~ Send("T(STM)="&$STMtemp&" K")
Send("Scan Speed= "&$Speed&" ")
Send("{LEFT}")
Send("{ENTER}")
send("T_STM= "&getTemperture()&" K")
;~ Send("{TAB}")
;~ highlightOneNoteWhite()
;~ Sleep(450)
;~ Send("{LEFT 2}")
EndFunc

Func writeSTS2Onenote()
$DirPath=getSessionPath()
grabSTS()
Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)
;~ Send("Filename")
;~ Send("{TAB}")
;~ Send("Sweep")
;~ Send("{TAB}")
;~ Send("Current")
;~ Send("{TAB}")
;~ Send("Comments")
;~ Send("{ENTER}")

;Find the last File
$FileSpec= "*.dat"
Local $lastestFile = GetLatestFile($DirPath, $FileSpec)
;open and reading the file
$hFile = Fileopen($DirPath&"\"&$lastestFile)

;~ Local $keys[]= [":BIAS:",":Z-CONTROLLER:",":COMMENT:"]
;~ $i=1
;~ While 1
;~ $line = FileReadLine ($hFile,$i)
;~ For $str In $keys
;~ if $line= $str  Then
;~
;~ EndIf
;~ Next

;~ $i=$i+1
;~ WEnd
$StartStr = FileReadLine ($hFile,17)
$Start= Stringsplit ($StartStr, Chr (9), 2)[1]
$StartmV =Number($Start)

$expoPos=StringInStr($Start,"e")
If $expoPos <>0 Then
   $StartmV=$StartmV*10^(Number(StringMid($Start,$expoPos+1)))
   EndIf

$EndStr = FileReadLine ($hFile,18)
$End= Stringsplit ($EndStr, Chr (9), 2)[1]
$EndmV =Number($End)
$expoPos=StringInStr($End,"e")
If $expoPos <>0 Then
   $EndmV=$EndmV*10^(Number(StringMid($End,$expoPos+1)))
   EndIf
$CurrentStr = FileReadLine ($hFile,38)
$Current= Stringsplit ($CurrentStr, Chr (9), 2)[1]
$CurrentnA=Number($Current)*1E9

FileClose($hFile)
newTableLine()
Send("STS")
Send("{ENTER}")
Send("{ENTER}")
Send ($lastestFile)
highlightOneNoteRed()
Send("{TAB}")
send("^v")
Send("{TAB}")
Send("{TAB}")
Send($StartmV &"->"&$EndmV &" V")
Send("{TAB}")
Send (Round($CurrentnA,3)&" nA")
Send("{TAB}")
;~ gotoFollowMe()
grabImageCursor(150)
Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)
Sleep(200)
send("^v")
send("T_STM= "&getTemperture()&" K")




EndFunc

Func writeGrid2Onenote()
   $DirPath=getSessionPath()

grabImage(250)
Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)


;Find the last Dir
Local $lastestFile = GetLatestDir($DirPath)
;open and reading the file
$hFile = Fileopen($DirPath&"\"&$lastestFile&"\"&$lastestFile&"_drift.csv",128)
;~ Local $keys[]= [":BIAS:",":Z-CONTROLLER:",":COMMENT:"]
;~ $i=1
;~ While 1
;~ $line = FileReadLine ($hFile,$i)
;~ For $str In $keys
;~ if $line= $str  Then
;~
;~ EndIf
;~ Next

;~ $i=$i+1
;~ WEnd
$ParamStr=""
FileSetPos ($hFile,0,0)
for $ii = 1 to 14
$line =FileReadLine ($hFile)
$ParamStr = $ParamStr & $line & @CR
Switch $ii
Case 1
   $imgCurrent="Img:"& Round(Number(StringSplit($line,":")[2])*1E9,2) & " nA"
Case 2
   $STSCurrent="STS:"&Round(Number(StringSplit($line,":")[2])*1E9,2)& " nA"
Case 5
   $start=Round(Number(StringSplit($line,":")[2])*1000,2)& " mV"
Case 6
   $end=Round(Number(StringSplit($line,":")[2])*1000,2)& " mV"
EndSwitch
Next


FileClose($hFile)

newTableLine()
Send("GRID")
Send("{ENTER}")
Send("{ENTER}")
Send ($lastestFile)
highlightOneNoteGreen()
Send("{TAB}")
send("^v")
Send("{TAB}")
Send("{TAB}")
Send($start)
Send(" -> ")
Send($end)
Send("{TAB}")
Send($imgCurrent&@CR)
Send($STSCurrent)
Send("{TAB}")
;~ Send("!h")
;~ Sleep(100)
;~ Send("f")
;~ Send("s")
;~ Send("10")
;~ send("{ENTER}")
;~ Sleep(300)
;~ Send($ParamStr)
ClipPut($ParamStr)
send("^v")
send("T_STM= "&getTemperture()&" K")
;~ Send("{LEFT}")
;~ Send("!h")
;~ Sleep(100)
;~ Send("f")
;~ Send("s")
;~ Send("14")
;~ send("{ENTER}")
EndFunc

Func writeHeader2Onenote()
   Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)

Send("Filename")
Send("{TAB}")
Send("Data")
Send("{TAB}")
Send("Bias")
Send("{TAB}")
Send("Current")
Send("{TAB}")
Send("Comments")
EndFunc

Func highlightOneNote()
sleep(500)
Send("!n")
   sleep(100)
   Send("t")
   sleep(100)
   Send("{ENTER}")
   sleep(100)
   Send("!j")
   sleep(100)
   Send("g")
   Send("{DOWN}")
Send("{DOWN}")
Send("{DOWN}")
Send("{DOWN}")
Send("{DOWN}")
Send("{DOWN}")
Send("{RIGHT}")
   Send("{ENTER}")

   EndFunc

Func highlightOneNoteBlue()
   Send("!j")
   sleep(100)
   Send("g")
   Send("{DOWN}")
      Send("{DOWN}")
   Send("{DOWN}")

Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
   Send("{ENTER}")

   EndFunc

Func highlightOneNoteGreen()
   Send("!j")
   sleep(100)
   Send("g")
   Send("{DOWN}")
      Send("{DOWN}")
   Send("{DOWN}")

Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")

   Send("{ENTER}")

   EndFunc

Func highlightOneNoteRed()
   Send("!j")
   sleep(100)
   Send("g")
   Send("{DOWN}")
   Send("{DOWN}")
   Send("{DOWN}")

Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")
Send("{RIGHT}")

   Send("{ENTER}")

   EndFunc

Func highlightOneNoteWhite()
   Send("!j")
   sleep(100)
   Send("g")
   Send("{ENTER}")

   EndFunc

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
   If StringInStr($aEnumList[$i][4], "image") <> 0 Then

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

Func moveONpage()
ShellExecuteWait("D:\Scripts\UnLockOneNote.Bat" )
Sleep(300)
Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)
WinWaitActive($hWndOneNote)
Send("^e")
Sleep(100)
Send("Quick Notes")
Sleep(100)
Send("{ENTER}")
Send("^!m")
WinWaitActive("Move or Copy Pages")
Sleep(300)
Send(getCurrentSection())
Sleep(300)
while WinActive("Move or Copy Pages") <> 0
   Send("{ENTER}")
   Sleep(200)
   WEnd
;~ send("{SHIFTDOWN}")
;~ send("{F9}")
;~ send("{SHIFTUP}")
Return 15000;

   EndFunc


func getCurrentSection()
   $ret = @YEAR&" "
   Switch number(@MON)
   Case 1
	  $ret =$ret & "January"
   Case 2
	  $ret =$ret & "February"
   Case 3
	  $ret =$ret & "March"
   Case 4
	  $ret =$ret & "April"
   Case 5
	  $ret =$ret & "May"
   Case 6
	  $ret =$ret & "June"
   Case 7
	  $ret =$ret & "July"
   Case 8
	  $ret =$ret & "August"
   Case 9
	  $ret =$ret & "September"
   Case 10
	  $ret =$ret & "October"
   Case 11
	  $ret =$ret & "November"
   Case 12
	  $ret =$ret & "December"
	  EndSwitch

   Return $ret
   EndFunc


Func MouseMove1Pixel()
   $pos=MouseGetPos()
   MouseMove($pos[0]+1,$pos[1])
   MouseMove($pos[0],$pos[1])
   EndFunc


Func copyWSXM()
Sleep(300)
   $pos = _WinAPI_GetMousePos()

$hwnd = _WinAPI_WindowFromPoint($pos)

$rer=_WinAPI_GetParent($hwnd)

$name=WinGetTitle($rer)

MouseClick("left")
Sleep(200)
Send("^c")

WinActivate("[CLASS:PPTFrameClass]")

Sleep(300)
Send("^v")
Sleep(300)
send("!jp")
Sleep(100)
send("w")
send("10")
send("{ENTER}")
ClipPut($name)
Send("^v")

EndFunc



Func btnGRIDClick()
writeGrid2Onenote()
EndFunc
Func btnHeaderClick()
writeHeader2Onenote()
EndFunc
Func btnONENOTEClick()
$queue=moveONpage()
EndFunc
Func btnSTSClick()
writeSTS2Onenote()
EndFunc
Func btnTOPOClick()
writeTopo2Onenote()
EndFunc
Func Form1Close()
Exit 0
EndFunc


HotKeySet("{F9}","writeGrid2Onenote")
HotKeySet("{F6}","writeHeader2Onenote")
HotKeySet("{F7}","writeTopo2Onenote")
HotKeySet("{F8}","writeSTS2Onenote")
HotKeySet("#h","highlightOneNote")
HotKeySet("#c","closeWSXM")
HotKeySet("#g","copyWSXM")

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=D:\Scripts\AutoIt\Labookhelper.kxf
$Form1 = GUICreate("labbookhelper recovery", 533, 83, 1752, 1299)
GUISetOnEvent($GUI_EVENT_CLOSE, "Form1Close")
;~ GUISetOnEvent($GUI_EVENT_MINIMIZE, "Form1Minimize")
;~ GUISetOnEvent($GUI_EVENT_MAXIMIZE, "Form1Maximize")
;~ GUISetOnEvent($GUI_EVENT_RESTORE, "Form1Restore")
$btnHeader = GUICtrlCreateButton("Header F6", 8, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnHeaderClick")
$btnONENOTE = GUICtrlCreateButton("move ON page", 424, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnONENOTEClick")
$btnSTS = GUICtrlCreateButton("STS F8", 216, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnSTSClick")
$btnTOPO = GUICtrlCreateButton("TOPO F7", 112, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnTOPOClick")
$btnGRID = GUICtrlCreateButton("GRID F9", 320, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnGRIDClick")
Global $lblGRID = GUICtrlCreateLabel ("Session:", 5, 55, 200, 25)
Global $txtGRID = GUICtrlCreateInput ("D:\STM Data\082016",55, 51, 200, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

$screenoff=150
$mouseposition=MouseGetPos()

While 1
	Sleep(100)
	$screenoff-=1
	if $queue <> 0 Then
	   $queue = $queue -100
	   if $queue <= 0 Then
		  ShellExecuteWait("D:\Scripts\LockOneNote.Bat" )
		  $queue=0
		  EndIf

	   EndIf
	if $screenoff ==0 Then
   	   $screenoff =150
   	   $newMousePos=MouseGetPos()
	   if  $newMousePos[0] == $mouseposition[0] And $newMousePos[1] == $mouseposition[1] Then
	   MouseMove1Pixel()

	  EndIf
	  	   $mouseposition=$newMousePos

	  EndIf
WEnd
