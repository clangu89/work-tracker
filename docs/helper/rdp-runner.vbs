Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then WScript.Quit

Set objShell = CreateObject("WScript.Shell")
strArg = objArgs(0)
targetDir = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\RDPManager"
strCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & targetDir & "\rdp-handler.ps1"" """ & strArg & """"
objShell.Run strCommand, 0, False
