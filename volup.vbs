WScript.Sleep(200000)
Do
Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys(chr(&hAF))
WScript.Sleep(7500)
Loop
