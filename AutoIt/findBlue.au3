$pos=WinGetPos("Scan Control")
$posTip=PixelSearch($pos[0]+282  , $pos[1]+109 , $pos[0]+1003  , $pos[1]+828 , 0x007EEF )
MouseMove($posTip[0]+1,$posTip[1]+4,1)

