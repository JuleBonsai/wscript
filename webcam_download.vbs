Set wshShell = WScript.CreateObject("WScript.Shell")
 Set wshSysEnv = wshShell.Environment("PROCESS")
 strHomePath = wshSysEnv("HOMEPATH")


' wait till 10:00 Uhr
If time<"10:00" Then 
	wscript.echo time & " waiting a little bit longer"
	WScript.Sleep 1800000 ' 30 minutes
else

' #########  Set your settings  #################
' set your websource
    strFileURL = "http://www.rosenhof.eu/webcam09/wc000M.jpg"
    strHDLocation = strHomePath&"\Desktop\webcam\Webcam_"&year(date)&right("0" & month(date),2)&right("0" & day(date),2)&"_"&right("0" & hour(time),2)&right("0" &minute(time),2)&right("0"&second(time),2)&".jpg"
wscript.echo strHDLocation


' Fetch the file
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

    objXMLHTTP.open "GET", strFileURL, false
    objXMLHTTP.send()


If objXMLHTTP.Status = 200 Then
Set objADOStream = CreateObject("ADODB.Stream")
objADOStream.Open
objADOStream.Type = 1 'adTypeBinary

objADOStream.Write objXMLHTTP.ResponseBody
objADOStream.Position = 0    'Set the stream position to the start

Set objFSO = Createobject("Scripting.FileSystemObject")
If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
Set objFSO = Nothing

objADOStream.SaveToFile strHDLocation
objADOStream.Close
Set objADOStream = Nothing
End if

Set objXMLHTTP = Nothing
' 30 seconds delay
WScript.Sleep 30000
end if
