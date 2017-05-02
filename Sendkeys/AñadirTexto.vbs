set shell = CreateObject("WScript.Shell")
shell.AppActivate("CONN")

shell.SendKeys "{F2}", True
shell.SendKeys "{HOME}", True
shell.SendKeys "(^{RIGHT})", True
shell.SendKeys "{LEFT}", True
shell.SendKeys "(^v)", True
shell.SendKeys "{ENTER}", True
shell.SendKeys "{DOWN}", True
