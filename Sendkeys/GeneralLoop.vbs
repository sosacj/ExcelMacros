set shell = CreateObject("WScript.Shell")
shell.AppActivate("CONN")

Dim index, count
'Number of commands to execute
count = 427
For index = 1 To count
	shell.SendKeys "nn{ENTER}{DOWN}", True
Next