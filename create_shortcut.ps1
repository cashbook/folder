$ws = New-Object -ComObject WScript.Shell
$lnkPath = [System.IO.Path]::Combine($env:USERPROFILE, "Desktop", "LinkSorter.lnk")
$sc = $ws.CreateShortcut($lnkPath)
$sc.TargetPath = "C:\Users\user\AppData\Local\Programs\Python\Python311\pythonw.exe"
$sc.Arguments = '"c:\Users\user\Documents\GitHub\folder\link_sorter.py"'
$sc.WorkingDirectory = "c:\Users\user\Documents\GitHub\folder"
$sc.Save()
Write-Host "Created: $lnkPath"
