set shell = CreateObject("WScript.Shell")
shell.AppActivate("CONN")

Dim index, count
'Number of commands to execute
count = 288
For index = 1 To count
	shell.SendKeys "{ENTER}A{ENTER}{DOWN}", True
Next