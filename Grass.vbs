Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
Set WshShell = CreateObject("WScript.Shell")
Command = "cmd /c shutdown /s /f"
intMessage = MsgBox( "Go touch grass ", _
	vbYesNo + vbExclamation + vbDefaultButton2, "NOW")

If intMessage = vbYes Then
	Set WshShellExec = WshShell.Exec(Command)

End if