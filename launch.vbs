Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "D:\Dev\stenoai\app"
WshShell.Run "cmd /c pnpm start", 0, False
