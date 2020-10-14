Set WshShell=WScript.CreateObject("WScript.shell")
WshShell.AppActivate"嘿嘿"
for i=1 to 99
WScript.Sleep 299
WshShell.SendKeys"^v"
WshShell.SendKeys i
WshShell.SendKeys"%s"
next