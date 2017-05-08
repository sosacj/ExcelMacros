set shell = CreateObject("WScript.Shell")
shell.AppActivate("Programa Zona 4 Este.xslx")

Dim index, count
'Number of commands to execute
count = 2
For index = 1 To count
	shell.SendKeys "{F2}(+{HOME})(^c){ENTER}{UP}{END}{RIGHT}{F2}{HOME}{+}(^v){+}{ENTER}{END}{RIGHT}", True
Next

