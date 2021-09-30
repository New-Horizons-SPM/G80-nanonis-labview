;~ Backup reminder
;~ This script display a reminder to start a manual backup

#include <TaskScheduler.au3>

#include <MsgBoxConstants.au3>

$backupYESNO=MsgBox($MB_OKCANCEL+$MB_ICONQUESTION+$MB_DEFBUTTON2+$MB_TOPMOST, "Backup Data ?", "Do you want to backup the data from this Computer ?"&@LF&"Bear in mind that this could interrupt current measurements.", 30)

if $backupYESNO == $IDOK Then
   ShellExecute("d:\scripts\backup.bat")
   EndIf