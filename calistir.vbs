Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Mevcut script'in bulunduğu klasörü al
strCurrentPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = strCurrentPath

' Python yolunu ayarla
Dim pythonPath
pythonPath = "python"

' Python programını çalıştır
objShell.Run """" & pythonPath & """ """ & strCurrentPath & "\main.py""", 0, False 