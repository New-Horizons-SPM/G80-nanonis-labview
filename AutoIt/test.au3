#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

$file=FileOpen("D:\Scripts\temp\latman.txt")
ConsoleWrite($file)
$latmanparam=FileReadLine($file)
FileClose($file)
ConsoleWrite($latmanparam)