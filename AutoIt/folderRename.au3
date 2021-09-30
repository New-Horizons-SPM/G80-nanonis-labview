#include <MsgBoxConstants.au3>



FileChangeDir("

D:\STM Data")
Local $hSearch = FileFindFirstFile("*")
    ; Check if the search was successful, if not display a message and return False.
    If $hSearch = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
        Return False
    EndIf

    ; Assign a Local variable the empty string which will contain the files names found.
    Local $sFileName = "", $iResult = 0

    While 1
        $sFileName = FileFindNextFile($hSearch)
        ; If there is no more file matching the search.
        If @error Then ExitLoop


        ; Display the file name.
		If @extended = 1  then
if StringLen($sFileName)=6 then
		   		DirMove($sFileName,StringRight($sFileName,4) &StringLeft($sFileName,2))

ConsoleWrite($sFileName&"->"&StringRight($sFileName,4) &StringLeft($sFileName,2)& @CR)
EndIf
		endif
    WEnd

    ; Close the search handle.
    FileClose($hSearch)
