'---------------------------
' Author: Stephen Millard
' Version: 1.0
' Date: 2023-05-19
'---------------------------

' Constants
Const adTypeBinary = 1


' Initialise
Dim objStream
Set objStream = CreateObject("ADODB.Stream")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")


' Read image from the path passed in the first parameter
objStream.Open
objStream.Type = adTypeBinary
objStream.LoadFromFile WScript.Arguments.Item(0)


' Convert to base64
Dim bytes: bytes = objStream.Read
Dim dom: Set dom = CreateObject("Microsoft.XMLDOM")
Dim elem: Set elem = dom.createElement("tmp")
elem.dataType = "bin.base64"
elem.nodeTypedValue = bytes


' Write conversion to file
strFileExtension = LCase(objFSO.GetExtensionName(WScript.Arguments.Item(0)))

select case strFileExtension
	case "jpg"
		strImageType = "jpeg"
	case "jpeg"
		strImageType = "jpeg"
	case "png"
		strImageType = "png"
end select 

strB64 = "data:image/" & strImageType & ";base64," & Replace(elem.text, vbLf, "")
strB64Path = WshShell.ExpandEnvironmentStrings("%temp%") & "\b64.txt"
Set objFile = objFSO.CreateTextFile(strB64Path,True)
objFile.Write strB64
objFile.Close


' Copy to clipboard from temporary file
WshShell.Run "cmd.exe /c type """ & strB64Path & """ | clip", 0, TRUE


' Comment out the next line if you don't want a visual confirmation.
msgbox "Completed for" & vbCrLf & objFSO.GetFileName(WScript.Arguments.Item(0)), vbOKOnly & vbInformation, "Base 64 Conversion"
