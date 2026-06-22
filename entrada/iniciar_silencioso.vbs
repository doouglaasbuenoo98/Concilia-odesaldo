Dim pasta, cmd
pasta = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
cmd = "pythonw """ & pasta & "\servidor.py"" --no-browser"
CreateObject("WScript.Shell").Run cmd, 0, False
