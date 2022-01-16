 Dim InputC, title, InputNo, ErrOverNo, ErrNaN, loopNo, loopTimes, abortLoop, ErrLow, InputDelay
  Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("errMsg/CB.E.001.cr", 1)
err001 = file.ReadAll
Set file = fso.OpenTextFile("errMsg/CB.E.002.cr", 1)
err002 = file.ReadAll
Set file = fso.OpenTextFile("errMsg/CB.E.003.cr", 1)
err003 = file.ReadAll
Set file = fso.OpenTextFile("endAbort/prop/titleBody.cr", 1)
titleBody = file.ReadAll
Set file = fso.OpenTextFile("endAbort/prop/version.cr", 1)
version = file.ReadAll

   Set WshShell = WScript.CreateObject("WScript.Shell")
title = titleBody + " " + version
InputC = InputBox("Enter the message you wanna spam below: ", title + "","the message you wanna spam here")
InputNo = InputBox("Enter the amount of time you want to spam here: ", title + "","amount")
InputDelay = InputBox("Enter your delay between each messages!" + Chr(13) + Chr(10) +  Chr(13) + Chr(10) + "   example: 1000 = I second", title, "1000")

if IsNumeric(InputNo) Then
'check numeric then start'
if InputNo = "40" Then
'its 40 then start'
WshShell.Run("abort.lnk")
for loopTimes = 1 to InputNo
WshShell.SendKeys InputC
WshShell.SendKeys "{ENTER}"
WScript.Sleep(InputDelay)
next
WshShell.Run("delay.lnk")
end if
if InputNo < "40" Then
if InputNo < "1" Then
ErrLow = MsgBox(err001 + Chr(13) + Chr(10) + "   Exit Code: CB.E.002", 16, title + "")
end if 
if InputNo > "1" Then
WshShell.Run("abort.lnk")
for loopTimes = 1 to InputNo
WshShell.SendKeys InputC
WshShell.SendKeys "{ENTER}"
WScript.Sleep(InputDelay)
next
WshShell.Run("delay.lnk")
end if
end if
if (InputNo > "40") Then
ErrOverNo = MsgBox(err002 + Chr(13) + Chr(10) + "   Exit Code: CB.E.001", 16, title + "")
end if
end if
if not IsNumeric(InputNo) Then
ErrNaN = MsgBox(err003 + Chr(13) + Chr(10) + "   Exit Code: CB.E.003", 16, title + "")
end if