  Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("prop/titleBody.cr", 1)
head = file.ReadAll
Set file = fso.OpenTextFile("prop/version.cr", 1)
body = file.ReadAll
title = head + " " + body
Dim done
done = MsgBox("   I have done spamming!", 32, "Cinnamo")