Set WshShell = CreateObject("WScript.Shell" ) 
WshShell.Run chr(34) & ".\ppt_ndi.exe" & Chr(34) & " --bg", 0 
Set WshShell = Nothing 
