Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "pythonw "" & scriptDir & "\folderlimiter.py""
CreateObject("Wscript.Shell").Run cmd, 0, False
