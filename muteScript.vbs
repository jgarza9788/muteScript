
'' this toggle's mute
' CreateObject("WScript.Shell").SendKeys(chr(173))

''take volume to zero
set WshShell = CreateObject("WScript.Shell")
WshShell.Run "%SystemRoot%\System32\SndVol.exe"
WScript.Sleep 500
For i = 1 To 10
    WshShell.SendKeys "{PGDN}"
    ' WScript.Sleep 250
Next
WshShell.SendKeys "%{F4}"  ' Alt+F4