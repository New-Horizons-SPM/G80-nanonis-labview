#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

Func SciNumber($line)
   $val=StringReplace($line,"e","E")
   ConsoleWrite($val&@CR)
   ConsoleWrite(Number(StringMid($val,StringInStr($val,"E")+1)))
   return Number($val)*10^Number(StringMid($val,StringInStr($val,"E")+1))
EndFunc

ConsoleWrite(ProcessList ()[0][0])


   $setpointFile=FileOpen("D:\Scripts\temp\setpoint.txt")
$str=FileRead($setpointFile)
Local $aResult = StringRegExp($str, '.*(?:Z Controller setpoint":)([0-9.e-]*).*(?:Bias \(V\)":)([0-9.e-]*).*(?:Z position \(m\)":)([0-9.e-]*).*(?:Amplitude \(m\)":)([0-9.e-]*)', $STR_REGEXPARRAYMATCH)
If Not @error Then
    MsgBox($MB_OK, "SRE Example 6 Result",Number($aResult[3]))
EndIf

;~ $str='{"Harmonic":1,"Phase Reference":-46.352001190185547,"LP Filter Order":1,"LP Cutoff Frq (Hz)":19.428092956542969,"Sync Filter on/off status":false,"HP Filter Order":0,"HP Cutoff Frq (Hz)":9.7140464782714844}{"Harmonic":2,"Phase Reference":180,"LP Filter Order":1,"LP Cutoff Frq (Hz)":19.428092956542969,"Sync Filter on/off status":false,"HP Filter Order":0,"HP Cutoff Frq (Hz)":9.7140464782714844}{"Frequency (Hz)":440,"Amplitude (mod. signal unit)":0.001500000013038516,"Lockin on/off status":false}'

;~ Local $aResult = StringRegExp($str, '(?:LP Cutoff Frq \(Hz\)":)([0-9.]*).*(?:Frequency \(Hz\)":)([0-9]*).*(?:Amplitude \(mod. signal unit\)":)([0-9.]*)', $STR_REGEXPARRAYMATCH)
;~ If Not @error Then
;~     MsgBox($MB_OK, "SRE Example 6 Result", number($aResult[0]))  .*(?:Bias (V)).*([0-9.]*).*(?:Z position (m)":).*([0-9.]*).*
;~ EndIf