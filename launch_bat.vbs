Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "C:\Users\werterzz\Desktop\2_run_ngrok2.bat" & Chr(34), 0
Set WshShell = Nothing
Set WshShell2 = CreateObject("WScript.Shell") 
WshShell2.Run chr(34) & "C:\Users\werterzz\Desktop\3_nextGenStartServer.bat" & Chr(34), 0
Set WshShell2 = Nothing