Set WshShell = WScript.CreateObject("WScript.Shell")
j = Hour(Now())
do while j < 22
  WScript.Sleep 30000

WshShell.SendKeys "{ESC}"
j = Hour(Now())
loop
