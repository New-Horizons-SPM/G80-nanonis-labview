#include <Array.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include "ImageClip.au3"
#include <_EnumChildWindows.au3>
#include <GDIPlus.au3>
#include "img.au3"
#include <FileConstants.au3>
Opt("SendKeyDelay",2)
Opt("SendKeyDownDelay",2)

Global $DirPath= ""
Global $queue=0;

Func checkFilename($name)
   ConsoleWrite($name)

   $split=StringSplit($name,".")
   $paramFile=""

   if $split[2] == "sxm" Then
	  $paramFile='d:\scripts\temp\lastSxm.txt'
   EndIf
    if $split[2] == "dat" Then
	  $paramFile='d:\scripts\temp\lastDat.txt'
	EndIf

   $oldname=FileRead($paramFile)
   if $oldname <> $name Then
	  $f=FileOpen($paramFile,$FO_OVERWRITE)
	  FileWrite($f,$name)
	  Return 0

	  EndIf
   $button=MsgBox(260,"ERROR",$name &@CR&"has already been added to the labbook."&@CR&"Make sure you save the data you want add."&@CR&"Do you want to write the data to OneNote anyway?")
   if $button == 6 Then
	  return 0
	  endif
   return 1
   EndFunc

Func activateOneNote()
Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
;Local $hWndOneNote = WinWait("[CLASS:ApplicationFrameWindow]", "", 10)
WinActivate($hWndOneNote)
WinWaitActive($hWndOneNote)
;~ Send("^e")
;~ send("qui")
;~ Send("{ENTER}")
   EndFunc


Func copyLMCurve()

    if WinExists("Lateral_manipulation.vi") AND WinGetState("Lateral_manipulation.vi") <> 23 Then











;~ 	      Opt("MouseClickDelay",400)

;~
;~ 	   $pos=WinGetPos("Lateral_manipulation.vi")
;~ 	   	   MouseClick("left",$pos[0]+450,$pos[1]+215,3,0);
;~ send("^c")
;~ $latmanparam=StringReplace(clipget(),"^","{^}");


;~ 	   MouseClick("right",$pos[0]+400,$pos[1]+400,1,0);
;~ 	   sleep(100)
;~ 	   	   MouseClick("right",$pos[0]+400,$pos[1]+400,1,0);
;~ 	   sleep(100)
;~ send("e")
;~ 	   sleep(100)
;~ send("{ENTER}")

;~ 	   sleep(100)
;~ send("{UP}")
;~ 	   sleep(100)
;~ send("{ENTER}")
;~ 	   sleep(100)
;~ send("{ENTER}")
activateOneNote()
$file=FileOpen("D:\Scripts\temp\latman.txt")
$latmanparam=FileReadLine($file)
FileClose($file)
$latmanparam=StringReplace($latmanparam,"^","{^}");
send($latmanparam)
    _ImageToClipResize("D:\Scripts\temp\latman.bmp",612,445)

send("^v")

;~    Opt("MouseClickDelay",10)

	   EndIf

   EndFunc


Func SciNumber($line)
;~    $val=StringReplace($line,"e","E")
;~    return Number($val*10^Number(StringMid($val,StringInStr($val,"E")+1)))
   return Number($line,3)
   EndFunc


Func getSessionPath()
   $pathfile=FileOpen("D:\Scripts\temp\path.txt")
   $path= FileReadLine($pathfile)
   GUICtrlSetData($txtGRID,$path)

   Return $path
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
   $hwd= WinActivate("Scan Control")
   WinMove($hwd, "", 490, 1, 1283, 842)
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
;~    MouseMove($posTip[0]+1,$posTip[1]+4,1)
   $x=$posTip[0]+1-282-$pos[0]
   $y=$posTip[1]-109+4-$pos[1]
   Sleep(100)
    _ScreenCapture_Capture ( @MyDocumentsDir & "\temp.jpg" , $pos[0]+282  , $pos[1]+109 , $pos[0]+1003  , $pos[1]+828 , True )
    drawCross($x,$y,@MyDocumentsDir & "\temp.jpg",@MyDocumentsDir & "\temp2.jpg")
    _ImageToClipResize(@MyDocumentsDir & "\temp2.jpg",$size,$size)
   EndFunc


Func getTemperture()
   $nanonishelperWnd=WinGetHandle("NanonisHelper")
   $title=WinGetTitle($nanonishelperWnd)

   if StringInStr($title,"Front") == 0 Then
   $filehwd=FileOpen("D:\Scripts\temp\temperatures.csv")
   $temp=Round(Number(FileReadLine($filehwd,2)),2)
   FileClose($filehwd)
Else
   $temp="NA"
   EndIf

   return $temp
   EndFunc

Func grabSTS()
     WinActivate("Bias Spectroscopy")
;~    WinWaitActive("Bias Spectroscopy")
	Sleep(200)
   $pos=WinGetPos("Bias Spectroscopy")
    _ScreenCapture_Capture ( @MyDocumentsDir & "\temp.jpg" , $pos[0]+20  , $pos[1]+380 , $pos[0]+610  , $pos[1]+730 , False )
    _ImageToClipResize(@MyDocumentsDir & "\temp.jpg",316,220)

   EndFunc


Func grabzSpec()
     WinActivate("Z Spectroscopy")
;~    WinWaitActive("Bias Spectroscopy")
	Sleep(200)
   $pos=WinGetPos("Z Spectroscopy")
    _ScreenCapture_Capture ( @MyDocumentsDir & "\temp.jpg" , $pos[0]+20  , $pos[1]+380 , $pos[0]+610  , $pos[1]+730 , False )
    _ImageToClipResize(@MyDocumentsDir & "\temp.jpg",316,220)

   EndFunc

Func newTableLine()
;~    MouseWheel("down",350)
   sleep(50)
   sendDalay("^{END}",100)
   sendDalay("^{ENTER}",100)

   ;Send("{APPSKEY}")
EndFunc


Func writeTopo2Onenote()
   checkNanonisHelper()
$DirPath=getSessionPath()
;~ $STMtemp= getSPMtemps()
grabImage(250)
activateOneNote()

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
Local $lastestFile = GetLatestFile($DirPath, $FileSpec)
if checkFilename($lastestFile) == 1 Then
   Return 1
   EndIf

;open and reading the file
$hFile = Fileopen($DirPath&"\"&$lastestFile,512)
;~checkSTSFileFormat($hFile)
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
$PixelStr=FileReadLine ($hFile,14)
$aPixel = Stringsplit ($PixelStr," ",0)
$Pixel=Number($aPixel[8]) &"x"& Number($aPixel[8])&" Pixels2"
$Size=Round(Number($aSize[12])*1E9,2)&"x"& Round(Number($aSize[23])*1E9,2)&" nm2"
$BiasStr = FileReadLine ($hFile,28)
$Bias= Number(StringReplace($BiasStr, " ", ""))*1000

$AngleStr=FileReadLine ($hFile,24)
$Angle= Number(StringReplace($AngleStr, " ", ""))
if $Angle == 0 Then
   $Angle = ""
Else
   $Angle = "Rotation: "& String($Angle)& " deg"
   EndIf


$CurrentStr = FileReadLine ($hFile,31)
$Speed=Round((Number($aSize[12])*1E9) /Number($aSpeed[14]),2)&" nm/s"
$Current= Stringsplit ($CurrentStr, Chr (9), 2)[3]
$Current_pA=Number($Current)*1E9
$Amplitude=""
$lockin_LP_CF=""

if StringCompare(FileReadLine ($hFile,195),'TRUE')==0 Then
   $line=FileReadLine ($hFile,181)
   $value=SciNumber($line)
   $Amplitude="Amplitude: "&String($value*10^12)&" pm"&@CR

   EndIf

if StringCompare(FileReadLine ($hFile,227),'ON')==0 Then
   $line=FileReadLine ($hFile,233)
   $value=SciNumber($line)
   $Amplitude="Amplitude: "&String($value*10^3)&" mV"&@CR
   $line=FileReadLine ($hFile,253)
   $value=SciNumber($line)
   $lockin_LP_CF="LockIn LP CF: "& String(Round($value,2))&" Hz"&@CR

   EndIf

$tipLift=""
$tipLiftValue=SciNumber(FileReadLine ($hFile,109))

if $tipLiftValue <> 0 Then
   $tipLift="TipLift: "&String($tipLiftValue*10^12)&" pm"&@CR
   EndIf


FileClose($hFile)
BlockInput (1)
newTableLine()
sleep(100)
sendDalay("TOPO",20)
sendDalay("{ENTER}",20)
sendDalay("{ENTER}",20)
sendDalay ($lastestFile,20)
sendDalay("{TAB}",50)
sendDalay("^v",50)
sendDalay($Pixel,20)
sendDalay("{ENTER}",20)
sendDalay($Size,20)
sendDalay("{TAB}",20)
sendDalay ($Bias&" mV",20)
sendDalay("{TAB}",20)
sendDalay ($Current_pA&" nA",20)
sendDalay("{TAB}",20)
sendDalay("Scan Speed= "&$Speed&" ",20)
sendDalay("{LEFT}",20)
sendDalay("{ENTER}",20)
sendDalay("T_STM= "&getTemperture()&" K"&@CR,20)
sendDalay($Amplitude,20)
sendDalay($lockin_LP_CF,20)
sendDalay($tipLift,20)
sendDalay($Angle,20)

BlockInput (0)
copyLMCurve()
EndFunc

Func sendDalay($message,$delayTime)
   Send($message)
   sleep($delayTime)
EndFunc
Func writeSTS2Onenote()

checkNanonisHelper()
$DirPath=getSessionPath()

;Find the last File
$FileSpec= "*.dat"
Local $lastestFile = GetLatestFile($DirPath, $FileSpec)

if checkFilename($lastestFile) == 1 Then
   Return 1
   EndIf

$zSpec = StringInStr($lastestFile,"Z-Spectroscopy") <> 0





if Not $zSpec Then
grabSTS()
FileCopy("D:\Scripts\temp\lockin.txt",StringReplace($DirPath&"\"&$lastestFile,".dat","_lockin.txt"))
$lockin=FileOpen("D:\Scripts\temp\lockin.txt")
$str=FileRead($lockin)
FileClose($lockin)
$lockinActive = StringInStr($str,'"Lockin on/off status":false') == 0
Local $aResult = StringRegExp($str, '(?:LP Cutoff Frq \(Hz\)":)([0-9.]*).*(?:Frequency \(Hz\)":)([0-9]*).*(?:Amplitude \(mod. signal unit\)":)([0-9.]*)', $STR_REGEXPARRAYMATCH)
$LP="LowPass:"&String(round(number($aResult[0]),1)& " Hz")
$freq="Freq:"&String($aResult[1])&" Hz"
$mod="Modulation:"&String(Round(number($aResult[2])*1000,2))&" mV"
$StartmV=""
$EndmV=""
$sweeps=""
$lockinActive=False
$Current_pA=0
;open and reading the file
$hFile = Fileopen($DirPath&"\"&$lastestFile)
checkSTSFileFormat($hFile)
ConsoleWrite($lastestFile)
$StartStr = FileReadLine ($hFile,20)
$Start= Stringsplit ($StartStr, Chr (9), 2)[1]
$StartmV =Number($Start)

$expoPos=StringInStr($Start,"e")
If $expoPos <>0 Then
   $StartmV=$StartmV*10^(Number(StringMid($Start,$expoPos+1)))
   EndIf

$EndStr = FileReadLine ($hFile,21)
$End= Stringsplit ($EndStr, Chr (9), 2)[1]
$EndmV =Number($End)
$expoPos=StringInStr($End,"e")
If $expoPos <>0 Then
   $EndmV=$EndmV*10^(Number(StringMid($End,$expoPos+1)))
EndIf

$sweeps= " " &Stringsplit (FileReadLine ($hFile,33), Chr (9), 2)[1] & " sweeps"
$lockinActive = $lockinActive or StringInStr(FileReadLine($hFile,34),"TRUE") > 0

;~ $CurrentStr = FileReadLine ($hFile,38)
$CurrentStr = FileReadLine ($hFile,172)

$Current= Stringsplit ($CurrentStr, Chr (9), 2)[1]
$Current_pA=Round(Abs(Number($Current))*1E12,0)
FileClose($hFile)

Else
  $lockinActive = False
   grabzSpec()
   FileCopy("D:\Scripts\temp\zSpec.txt",StringReplace($DirPath&"\"&$lastestFile,".dat","_zSpec.txt"))
   FileCopy("D:\Scripts\temp\osci.txt",StringReplace($DirPath&"\"&$lastestFile,".dat","_osci.txt"))

$setpointFile=FileOpen("D:\Scripts\temp\setpoint.txt")
$str=FileRead($setpointFile)
FileClose($setpointFile)
Local $setPointResult = StringRegExp($str, '.*(?:Z Controller setpoint":)([0-9.e-]*).*(?:Bias \(V\)":)([0-9.e-]*).*(?:Z position \(m\)":)([0-9.e-]*).*(?:Amplitude \(m\)":)([0-9.e-]*)', $STR_REGEXPARRAYMATCH)
$Current_pA=Number($setPointResult[0])*1E9
$StartmV=Number($setPointResult[1])*1E3
$value=Round(Number($setPointResult[3])*1E12,3)
$Amplitude="Amplitude: "&String($value)&" pm"&@CR

$zSpecFile=FileOpen("D:\Scripts\temp\zSpec.txt")
$zSpecstr=FileRead($zSpecFile)
FileClose($zSpecFile)
Local $zSpecResult = StringRegExp($zSpecstr, '.*(?:Z offset \(m\)":)([0-9.e-]*)(?:,"Z distance \(m\)":)([0-9.e-]*)', $STR_REGEXPARRAYMATCH)
ConsoleWrite($zSpecstr)
$zoffset=Number($zSpecResult[0])*1E12
$sweep=Number($zSpecResult[1])*1E12
$RetractCondition=StringReplace(StringMid($zSpecstr,StringInStr($zSpecstr,",",0,-1)+1),'"','')
$RetractCondition=StringReplace($RetractCondition,'}','')

EndIf


; bias STS
activateOneNote()
newTableLine()
Send("STS")
Send("{ENTER}")
Send("{ENTER}")
Send($lastestFile)
highlightOneNoteRed()
Send("{TAB}")
send("^v")
Send("{TAB}")
Send("{TAB}")
if $zSpec Then
   Send (Round($StartmV,3)&" mV")
   Else
Send($StartmV &"->"&$EndmV &" V")
EndIf
Send("{TAB}")
Send (Round($Current_pA,3)&" pA")
Send("{TAB}")
;~ gotoFollowMe()
grabImageCursor(150)
activateOneNote()
Sleep(200)
send("^v")
if $lockinActive Then
   send(@CR&"T_STM= "&getTemperture()&" K"&@CR & $freq &@CR&$mod &@CR&$LP &@CR& $sweeps)
Elseif $zSpec Then
   send(@CR&"T_STM= "&getTemperture()&" K"&@CR & $Amplitude & @CR)
   send(Round($zoffset,3)&" pm ->"&Round($sweep+$zoffset,3) &" pm" & @CR)
   send($RetractCondition)
Else
   send(@CR&"T_STM= "&getTemperture()&" K"&@CR & $sweeps)
   endif



EndFunc


Func writeGridCloud2Onenote()
$DirPath=getSessionPath()

grabImage(250)
activateOneNote()

newTableLine()
Send("GRID")
Send("{ENTER}")
Send("{ENTER}")
Send ($DirPath)
highlightOneNoteGreen()
Send("{TAB}")
send("^v")
Send("{TAB}")
Send("{TAB}")

Send("{TAB}")

Send("{TAB}")

EndFunc

Func writeGrid2Onenote()

   $DirPath=getSessionPath()
if StringInStr($DirPath,"topo") Then
   $DirPath=StringLeft($DirPath,StringInStr($DirPath,'\',0,3))
   EndIf
grabImage(250)
activateOneNote()


;Find the last Dir
Local $lastestFile = GetLatestDir($DirPath)
ConsoleWrite($lastestFile)

FileCopy("D:\Scripts\temp\lockin.txt",$DirPath&"\"&$lastestFile&"\"&$lastestFile&"_lockin.txt")
FileCopy("D:\Scripts\temp\osci.txt",$DirPath&"\"&$lastestFile&"\"&$lastestFile&"_osci.txt")
FileCopy("D:\Scripts\temp\zSpec.txt",$DirPath&"\"&$lastestFile&"\"&$lastestFile&"_zSpec.txt")

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
$line =StringReplace($line,'"','')
$line =StringReplace($line,'{','')
$line =StringReplace($line,'}','')
$line =StringReplace($line,',',@CR)
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
send(@CR&"T_STM= "&getTemperture()&" K")
;~ Send("{LEFT}")
;~ Send("!h")
;~ Sleep(100)
;~ Send("f")
;~ Send("s")
;~ Send("14")
;~ send("{ENTER}")
EndFunc

Func writeHeader2Onenote()
activateOneNote()
sleep(50)
Send("Filename")
sleep(50)
Send("{TAB}")
sleep(50)
Send("Data")
sleep(50)
Send("{TAB}")
sleep(50)
Send("Bias")
sleep(50)
Send("{TAB}")
sleep(50)
Send("Current")
sleep(50)
Send("{TAB}")
sleep(50)
Send("Comments")
sleep(50)
EndFunc

Func writeEvent2Onenote()
activateOneNote()

Send("Time")
Send("{TAB}")
Send("Event")
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
activateOneNote()
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

Func copyImg2Canvas()
   $pos=MouseGetPos()
   MouseClick("left",$pos[0],$pos[1],2,0)
   $winpos=WinGetPos("Scan Inspector")
   MouseClick("left",$winpos[0]+196,$winpos[1]+110,1,0)
   MouseMove($pos[0],$pos[1],0)
   EndFunc

Func btnGRIDClick()
writeGrid2Onenote()
EndFunc
Func btnGRIDCloudClick()
writeGridCloud2Onenote()
EndFunc
Func checkTeamViewer()
  if WinExists("Sponsored session")  Then
   WinActivate("Sponsored session")
   Sleep(100)
   $pos=WinGetPos("Sponsored session")
   MouseClick("left",$pos[0]+430,$pos[1]+142,1,0)
EndIf
   EndFunc
Func btnHeaderClick()
writeHeader2Onenote()
EndFunc

Func btnEventClick()
writeEvent2Onenote()
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
Func checkNanonisHelper()

if StringInStr(WinGetTitle("NanonisHelper"),"Front") <> 0 Then
   $box= MsgBox(5,"ERROR","Please make sure the NanonisHelper vi is running")
   if $box == 4 Then
	  checkNanonisHelper()
	  Else
   Exit(1)
   endif
   EndIf
EndFunc

Func checkSTSFileFormat($hFile)
   $line=FileReadLine($hFile,170)
   If StringCompare("[DATA]",$line) <> 0 then
	  MsgBox(0,"Erong STS file Format","Double-check module parameters in Nanonis STS Save options, Dialog window"&@CR&" Closing now...")
	  exit(-1)
	  EndIf
EndFunc

HotKeySet("{F2}","copyImg2Canvas")
HotKeySet("{F9}","writeGrid2Onenote")
HotKeySet("{F6}","writeHeader2Onenote")
HotKeySet("{F7}","writeTopo2Onenote")
HotKeySet("{F8}","writeSTS2Onenote")
HotKeySet("{F5}","writeEvent2Onenote")
HotKeySet("#h","highlightOneNote")
HotKeySet("#c","closeWSXM")
HotKeySet("#g","copyWSXM")

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>


Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=D:\Scripts\AutoIt\Labookhelper.kxf
$Form1 = GUICreate("labbookhelper", 637, 83, 752, 1299)
GUISetOnEvent($GUI_EVENT_CLOSE, "Form1Close")
;~ GUISetOnEvent($GUI_EVENT_MINIMIZE, "Form1Minimize")
;~ GUISetOnEvent($GUI_EVENT_MAXIMIZE, "Form1Maximize")
;~ GUISetOnEvent($GUI_EVENT_RESTORE, "Form1Restore")
$btnEvent = GUICtrlCreateButton("Event F5", 8, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnEventClick")
$btnHeader = GUICtrlCreateButton("Header F6", 112, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnHeaderClick")
$btnONENOTE = GUICtrlCreateButton("move ON page", 528, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnONENOTEClick")
$btnSTS = GUICtrlCreateButton("STS F8", 320, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnSTSClick")
$btnTOPO = GUICtrlCreateButton("TOPO F7", 216, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnTOPOClick")
$btnGRID = GUICtrlCreateButton("GRID F9", 424, 16, 97, 25)
GUICtrlSetOnEvent(-1, "btnGRIDClick")
$btnGRIDcould = GUICtrlCreateButton("GRID zSpec", 424, 51, 97, 25)
GUICtrlSetOnEvent(-1, "btnGRIDCloudClick")

Global $lblGRID = GUICtrlCreateLabel ("Session:", 5, 55, 200, 25)
Global $lblTimer = GUICtrlCreateLabel ("0", 528, 55, 200, 25)



Global $txtGRID = GUICtrlCreateInput ("D:\STM Data\"&@YEAR&@MON,55, 51, 200, 25)
GUISetState(@SW_SHOW)
WinSetOnTop($Form1, "", 1)

#EndRegion ### END Koda GUI section ###

$screenoff=150
$mouseposition=MouseGetPos()
checkNanonisHelper()
getSessionPath()
While 1
	Sleep(300)
	checkTeamViewer()
	$screenoff-=1
	if $queue <> 0 Then
	   $queue = $queue -100
	   GUICtrlSetData($lblTimer,$queue/1000)
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
