set shell = CreateObject("WScript.Shell")
shell.AppActivate("Unificación de costeo - Bianchi.xlsm")

shell.SendKeys "(^c)", True
shell.SendKeys "{DOWN}", True
shell.SendKeys "(^v)", True
shell.SendKeys "{F2}", True
shell.SendKeys "{.}", True
shell.SendKeys "1", True
shell.SendKeys "{ENTER}", True
shell.SendKeys "{UP}", True
shell.SendKeys "{END}", True
shell.SendKeys "{DOWN}", True