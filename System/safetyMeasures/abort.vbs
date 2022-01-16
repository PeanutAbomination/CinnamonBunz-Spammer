dim start
abort = msgbox("   Please select yes if you want to abort, select no to ignore" + Chr(13) + Chr(10) + "   this message", 4 , "Abort")

If abort=vbYes then
CreateObject("WScript.Shell").run("forceKillSpam.bat")
else
CreateObject("WScript.Shell").Popup "You ignored the abort selection box, you can't abort anymore now.", 3, "Abort"
end if