Dim shell,command
command = "powershell.exe -nologo -command "".\himawari.ps1"""
Set shell = CreateObject("WScript.Shell")
shell.Run command,0