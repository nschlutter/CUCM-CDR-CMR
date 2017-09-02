# CUCM-CDR-CMR
Cisco Unified Communications Manager (CallManager) CDR and CMR to MySQL import using PowerShell  

# Requirements:  

MySQL .NET Connector  
SysInternals handle.exe (handle.exe needs to be in the executing users %PATH%)


# Files and functions:

CDR Scheduled Task.xml - Windows Scheduled Task for CDR Processing  
CMR Scheduled Task.xml - Windows Scheduled Task for CMR Processing  
CDR-MySQL-Import.ps1 - PowerShell script for processing CDR files  
CMR-MySQL-Import.ps1 - PowerShell script for processing CMR files  
Get-HandlesAndProcessIds.ps1 - Helper script used by handle.ps1 script (requires handle.exe from SysInternals Suite)  
handle.ps1 - Helper script called to determine if data files are open by another process and should be skipped
