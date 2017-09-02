. "$PSScriptRoot\Get-HandlesAndProcessIds.ps1"

$fileName = "c:\inetpub\ftproot\cdr\temp\"

$handles = & handle.exe -a -u -accepteula "$fileName"
$handleCount = (Get-HandlesAndProcessIds -HandleInfo $handles | Measure-Object)
$openHandles = Get-HandlesAndProcessIds -HandleInfo $handles
Write-Host "Count:" $handleCount.count
$openHandles



#$openHandles | Get-Member -MemberType Property