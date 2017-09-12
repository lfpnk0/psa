$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Running Powershell Version " + $PSVersionTable.PSVersion.major,0,"Done",0x1)
