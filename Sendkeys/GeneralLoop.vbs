

Option Explicit
 Dim x
 set x=createobject("wscript.shell")
 x.AppActivate("CONN")

 x.sendkeys"PASSWORD"

 x.sendkeys"+{CLICK LEFT;50;60}"
