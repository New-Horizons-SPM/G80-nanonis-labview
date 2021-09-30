

;~ If $oStatusCode == 200 then
;~  $file = FileOpen("Received.html", 2) ; The value of 2 overwrites the file if it already exists
;~  FileWrite($file, $oReceived)
;~  FileClose($file)


;~ EndIf


;~ #	Website URL	Label	API Key
Func getScopusPage()

$query="Real-space imaging of molecular structure and chemical bonding by single-molecule inelastic tunneling probe"
$query="10.1039/c5nr00302d"
send("^c")

$query=ClipGet()
$query=StringReplace($query,"DOI","")
$query=StringReplace($query,":","")
$query=StringReplace($query," ","")

$scopusAPI="22c98dc9183264ad495f85ecc80ce058"
$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
;~ $oHTTP.SetRequestHeader("X-ELS-APIKey",$scopusAPI)
;~ $oHTTP.Open("GET", "http://api.elsevier.com/content/search/scopus?query=title("&$query&")&apiKey="&$scopusAPI, False)
$oHTTP.Open("GET", "http://api.elsevier.com/content/search/scopus?query=DOI("&$query&")&apiKey="&$scopusAPI, False)

$oHTTP.Send()
$oReceived = $oHTTP.ResponseText
$oStatusCode = $oHTTP.Status

$response=StringSplit($oReceived,"{")
$tmp=StringSplit(StringSplit($response[10],",")[3],":")
$url=$tmp[2]&":"&$tmp[3]
$url=StringReplace($url,"}","")
$url=StringReplace($url,'"',"")


;~ $eid="2-s2.0-84901207665"
;~ $title="Real-space+imaging+of+molecular+structure+and+chemical+bonding+by+single-molecule+inelastic+tunneling+probe"
;~ $url="https://www.scopus.com/record/display.uri?eid="&$eid&"&origin=resultslist&sort=plf-f&src=s&st1="&$title&"&st2"
ShellExecute($url)
EndFunc

HotKeySet("{F4}","getScopusPage")

While 1
   Sleep(100)
   WEnd