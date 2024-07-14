SFC /ScanNow
DISM /Online /Cleanup-Image /CheckHealth
DISM /Online /Cleanup-Image /ScanHealth
DISM /Online /Cleanup-Image /RestoreHealth
DISM /Online /cleanup-image /startcomponentcleanup /resetbase
C:\PowerShell\ResetWU.ps1
C:\PowerShell\Cate.ps1
Import-Module PSWindowsUpdate
Get-WindowsUpdate -AcceptAll -ForceDownload -ForceInstall -AutoReboot -MicrosoftUpdate -RecurseCycle 2 -Verbose

