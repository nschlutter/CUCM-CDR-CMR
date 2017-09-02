<#
.Description
Converts the text output of the SysInternals tool "handle.exe" into an object array
 
.Example
 
$handles = & handle.exe -a  -u -accepteula "MY_LOCKED_ASSEMBLY.dll"
$allHandles = Get-HandlesAndProcessIds -HandleInfo $handles
$allHandles
 
---------------------------
 
ProcessName : dllhost.exe
ProcessId   : 6960
Type        : File
User        : DOMAIN\user
HandleId    : B34
Name        : MY_LOCKED_ASSEMBLY.dll
AppPoolName : 
 
#>

Add-Type -Language CSharp @"
public class Handle{
    public string ProcessName { get;set;}
    public int ProcessId {get;set;}
    public string Type {get;set;}
    public string User {get;set;}
    public string HandleId { get;set;}
    public string Name { get; set;}
    public string AppPoolName {get;set;}
}
"@;
 
function Get-HandlesAndProcessIds {
    [OutputType([Handle[]])]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [object[]]$HandleInfo)

    $handle = new-object Handle
    foreach ($line in $handles) {
        $line=$Line.trim()
        if($line -like "*pid:*") {
            # $result = $line | Select-String –pattern '^([^ ]*)\s*pid: ([0-9]*)\s*type: ([^ ]*)\s*([^ ]*)\s*(.*?): (.*)'
            $result = $line | Select-String –pattern '^(.*)\s*pid: ([0-9]*)\s*type: ([^ ]*)\s*([^ ]*)\s*(.*?): (.*)'

            $handle.ProcessName =  $result.Matches[0].Groups[1].Value
            $handle.ProcessId = [int]::Parse( $result.Matches[0].Groups[2].Value)
            $handle.Type = $result.Matches[0].Groups[3].Value
            $handle.User = $result.Matches[0].Groups[4].Value
            $handle.HandleId =  $result.Matches[0].Groups[5].Value
            $handle.Name = $result.Matches[0].Groups[6].Value
 
            if($handle.ProcessName -like "w3wp.exe") {
                $process = Get-AppPoolProcesses | where { $_.ProcessId -eq $handle.ProcessId }
                if($process) {
                    $handle.AppPoolName = $Process.AppPoolName
                }
            }
 
            $handle
        }
    }
}
 
function Get-AppPoolProcesses {
    Get-WmiObject -NameSpace 'root\WebAdministration' -class 'WorkerProcess' -ComputerName 'LocalHost' |
        select AppPoolName, ProcessId |
            sort -Property AppPoolName
}