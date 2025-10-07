Dim objShell
Set objShell = CreateObject("WScript.Shell")
a=MsgBox("Are you sure you want to run this dangerous piece of software?", vbYesNo+vbQuestion, "Confirm")

if a = vbNo Then
   objShell.run "script.bat"
End If   

b=MsgBox("LAST WARNING! Are you going to run this script for real?!? Think about the damages!", vbYesNo+vbCritical, "LAST WARNING")

if b = vbNo Then
   objShell.run "script.bat"
End If

WScript.Sleep 500

objShell.run "main.bat"
