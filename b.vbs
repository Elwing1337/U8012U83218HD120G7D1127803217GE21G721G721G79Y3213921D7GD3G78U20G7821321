Set WshShell = CreateObject("WScript.Shell")
Set WshNetwork = CreateObject("WScript.Network")

' Bilgi mesajı göster
MsgBox "PRESS DONE WHEN ALL THE THINGS CLOSED AND COMPLETED.", vbInformation + vbOKOnly, "DONE"

' Bilgisayarı yeniden başlat
WshShell.Run "shutdown /r /t 10", 0, False

' Belleği temizle
Set WshShell = Nothing
Set WshNetwork = Nothing
Set fso = Nothing

