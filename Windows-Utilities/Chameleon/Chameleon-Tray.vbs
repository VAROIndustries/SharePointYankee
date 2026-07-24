' Chameleon (Tray) silent launcher.
' Double-click to start the tray app with no PowerShell console window.
' Put a shortcut to this file in your Startup folder to run it at logon:
'   shell:startup  (paste into the Run dialog / File Explorer address bar)
Dim shell, here
Set shell = CreateObject("WScript.Shell")
here = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
shell.Run "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & here & "Chameleon-Tray.ps1""", 0, False
