func moveCurrentPage()
ShellExecuteWait("D:\Scripts\UnLockOneNote.Bat" )
Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)
Send("^e")
Sleep(100)
Send("Quick Notes")
Sleep(100)
Send("{ENTER}")
Send("^!m")
Send(getCurrentSection())
Sleep(100)
Send("{ENTER}")
send("{SHIFTDOWN}")
send("{F9}")
send("{SHIFTUP}")
Sleep(5000)
ShellExecuteWait("D:\Scripts\LockOneNote.Bat" )
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

moveCurrentPage()
