Set WshShell = CreateObject("WScript.Shell")
' baslat.bat dosyasını penceresiz (0) modunda çalıştırır
WshShell.Run "cmd.exe /c baslat.bat", 0, False
Set WshShell = Nothing