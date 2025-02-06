Dim objResult

Set objShell = WScript.CreateObject("WScript.Shell")    
MsgBox "starting keyboard jiggler"

Do While True
  objResult = objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
  Wscript.Sleep (6000)
Loop
