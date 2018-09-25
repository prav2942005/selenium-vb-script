On Error Resume Next

Set restReq = CreateObject("Msxml2.ServerXMLHTTP")
Set objshell = CreateObject("WScript.Shell")

url = "http://localhost:9515/"
chrome_capabilities = "{""capabilities"": {""alwaysMatch"": {""browserName"": ""chrome"", ""platform"": ""ANY"",   ""version"": """"},  ""firstMatch"": []}, ""desiredCapabilities"": {""browserName"": ""chrome"",  ""platform"": ""ANY"",  ""version"": """"}}"
session_id = ""
test_url = "{""url"":""https://www.google.com/""}"

'The below codes will be executed only if the chrome driver is already running
'Based on WebDriver specification as in W3C webdriver specifications


'GET STATUS FROM CHROME DRIVER ---
restReq.open "GET", url & "status", false
restReq.setRequestHeader "Content-Type", "application/json"
restReq.send("")

'!! Revisit the below!!!
'Need to check the logic to start the chrome driver if its not already running 
If Err.Number <> 0 Then
	objShell.Run("""C:\D-Drive\HP ALM\TCS\Selenium\chromedriver_win32.exe""")
	Set objShell = Nothing
	objshell.sleep 1000
End If


'POST SESSION IN CHROME DRIVER TO START CHROME BROWSER ---
restReq.open "POST", url & "session", false
restReq.setRequestHeader "Content-Type", "application/json"
restReq.send(chrome_capabilities)
response = Left(restReq.responseText, Instr(restReq.responseText,",") - 1)
response = Right(response, Len(response) - Instr(response, ":"))
response = Left(response, Len(response) - 1)
session_id = Right(response, Len(response) - 1)


'POST URL TO OPEN IN CHROME BROWSER ---
restReq.open "POST", url & "session/" & session_id & "/url", false
restReq.setRequestHeader "Content-Type", "application/json"
restReq.send(test_url)

'GET URL TO TAKE SCREEN SHOT IN CHROME BROWSER ---
restReq.open "GET", url & "session/" & session_id & "/screenshot", false
restReq.setRequestHeader "Content-Type", "application/json"
restReq.send("")
response = Right(restReq.responseText, Len(restReq.responseText) - Instr(restReq.responseText,"""value"":") - 8)
base64stringimage = Left(response, Len(response) - 2)
Base64Decode base64stringimage, "C:\D-Drive\Screenshots\test1.png"


'DELETE TO CLOSE CHROME BROWSER ---
restReq.open "DELETE", url & "session/" & session_id & "/window", false
restReq.setRequestHeader "Content-Type", "application/json"
restReq.send("")



'FUNCTIONS :

Sub Base64Decode(vCode, file)

	On Error Resume Next
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Stream_BinaryToString file, oNode.nodeTypedValue
    Set oNode = Nothing
    Set oXML = Nothing
	
	If Err.Number <> 0 Then
		Msgbox Err.Description
	End if	
	
End Sub

Sub Stream_BinaryToString(file, datawrite)

  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = 1
  BinaryStream.Open
  BinaryStream.Write datawrite
  
  BinaryStream.SaveToFile file, 2

  Set BinaryStream = Nothing
  
  If Err.Number <> 0 Then
		Msgbox Err.Description
	End if
  
  
End Sub




