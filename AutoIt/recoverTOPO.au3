
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

Func newTableLine()
   Send("^{END}")
   Sleep(100)
   Send("{APPSKEY}")
   Sleep(100)
   Send("a")
   Sleep(100)
   Send("b")
EndFunc
$i=1
While $i<24

MouseClick("left",2733,24,1)
MouseClick("left",3822,555,1) ;open export

MouseClick("left",3877,782,1)
MouseClick("left",3877,782,1)

Send("{CTRLDOWN}c{CTRLUP}")
$name=ClipGet()
MouseClick("left",3921,925,1)

MouseClick("left",2789, 399,1)
MouseClick("left",2789, 399,1)
Send("{CTRLDOWN}c{CTRLUP}")
$Bias=ClipGet()
MouseClick("left",2791, 450,1)
MouseClickDrag("left",2791, 450, 2698, 447)
Sleep(100)
Send("{CTRLDOWN}c{CTRLUP}")
$CurrentnA=ClipGet()


MouseClick("left",2779, 304,1)
MouseClick("left",2779, 304,1)

Send("{CTRLDOWN}c{CTRLUP}")
$Size=ClipGet()
MouseClick("left",2733,24,1)
MouseClick("left",3562, 74,1)
Send("{CTRLDOWN}b{CTRLUP}")


Local $hWndOneNote = WinWait("[CLASS:Framework::CFrame]", "", 10)
WinActivate($hWndOneNote)

newTableLine()
Send("TOPO")
Send("{ENTER}")
Send("{ENTER}")
Send ($name)
highlightOneNoteBlue()
Send("{TAB}")
send("^v")
Send($Size)
Send("{TAB}")
Send ($Bias&"")
Send("{TAB}")
Send ($CurrentnA&"")
Send("{TAB}")
Send("{LEFT}")
Send("{ENTER}")
$i+=1
MouseClick("left",2824, 131,1)
Sleep(300)
wend



