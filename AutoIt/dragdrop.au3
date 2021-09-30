#include <_FileDragDrop.au3>

$filename=clipget()
$filepath="\\storage.erc.monash.edu\shares\R-SCI-Schiffrin-lab\instrument\STM Data\"
$session= "20"&StringLeft($filename,2)&StringMid($filename,3,2)&'\'



$hNPPlus=WinGetHandle("[CLASS:AfxMDIFrame80s]")
$aNPPPos=WinGetPos($hNPPlus)
;~ WinActivate($hNotePad)



$iRet=_FileDragDrop($hNPPlus,$filepath&$session&$filename,$aNPPPos[2]/2,$aNPPPos[3]/2,'|',False)
ConsoleWrite("Return: "&$iRet&" @error="&@error&", @extended="&@extended&@CRLF)
ConsoleWrite($filepath&$session&$filename&@CRLF)

