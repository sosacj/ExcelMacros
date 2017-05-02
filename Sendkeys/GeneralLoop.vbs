set shell = CreateObject("WScript.Shell")
shell.AppActivate("Microsoft Excel")

Dim index, count
'Number of commands to execute
count = 100
For index = 1 To count
	shell.SendKeys "{END}{DOWN}{DOWN}(^v)", True
Next