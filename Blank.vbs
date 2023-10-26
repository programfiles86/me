do
Set wshshell = wscript.CreateObject("WScript.Shell")
Wshshell.run "Notepad"
wscript.sleep 100
wshshell.sendkeys "B"
wscript.sleep 100
wshshell.sendkeys "L"
wscript.sleep 100
wshshell.sendkeys "A"
wscript.sleep 100
wshshell.sendkeys "N"
wscript.sleep 100
wshshell.sendkeys "K"
loop
