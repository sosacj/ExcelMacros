set shell = CreateObject("WScript.Shell")
shell.AppActivate("CONN")

Dim index, count
'Number of commands to execute
count = 48
For index = 1 To count
	shell.SendKeys "{ENTER}a{ENTER}{DOWN}", True
Next