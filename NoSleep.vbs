Set objShell = WScript.CreateObject("WScript.Shell")
Do While True
  objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
  Wscript.Sleep (5 * 60 * 1000)
Loop
