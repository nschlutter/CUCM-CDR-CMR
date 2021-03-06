<#
Purpose: Import Cisco CallManager Call Detail Records (CDR) in to MySql
Created By: Nathan Schlutter
https://github.com/nschlutter/CUCM-CDR-CMR
Date: September 2017
Version: 0.2
#>
[CmdletBinding()]
Param()

# Scripts will stop execuction if an error occurs
$ErrorActionPreference = 'Stop'

# Set required variables
# Set CDR file path
$cdrPath = "C:\Inetpub\ftproot\cdr"
# Set CDR file arhive path where files are moved after process
$archivePath = "C:\Inetpub\ftproot\cdr\_archive_batch"

# Set database information
$mySqlhost = "localhost"
$dbName = "call_detail_records"
$dbUsername = "database username"
$dbPassword = "database password"
$dbCdrTbl = "cdr"
$dbCmrTbl = "cmr"

#
# You should not need to change anything below this line
#

# Include function to determine if cdr files are open by another program
Try {
    . "$PSScriptRoot\Get-HandlesAndProcessIds.ps1"
} Catch [System.Exception] {
    Write-Verbose $_.Exception|format-list -force
    Write-Output $_.Exception.Message
}

$logFileName = "CDR_Import_" + [System.DateTime]::Now.ToString("M_d_yyyy") + ".log"
$logFile = "$PSScriptRoot\logs\$logFileName"
Try {
    $numCols=120
    mode con cols=$numCols
    Start-Transcript -Path $logFile -Append -NoClobber
} Catch {
}

# Set up MySql connection string and MySql object
Try {
    [void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data") 
} Catch [System.Exception] {
    Write-Verbose $_.Exception|format-list -force
    Write-Output $_.Exception.Message
}

$connStr="server=" + $mySqlhost + ";database=" + $dbName + ";Persist Security Info=false;user id=" + $dbUsername + ";pwd=" + $dbPassword + ";default command timeout=30;"
$conn = New-Object MySql.Data.MySqlClient.MySqlConnection($connStr)

# Check that CDR file and archive paths are valid 
If (!(Test-Path -Path $cdrPath)) {
    Write-Output "ERROR - Invalid CDR file path '$cdrPath'" 
} ElseIf (!(Test-Path -Path $archivePath)) {
    Write-Output "ERROR - Invalid CDR archive file path '$archivePath'"
} Else {
    # If both paths are OK then we continue
}

# Loop through all files starting with prefix 'cdr_'

$totalRows = 0
$totalFiles = 0

Get-ChildItem $cdrPath -Filter cdr_* | `
Foreach-Object{

    #Write-Output "$cdrPath\$($_.name)"

    # Check if the file we are going to process is already open
    $inputFile = $cdrPath + "\" + $($_.name)

    $handles = & handle.exe -a -u -accepteula "$inputFile"
    $openHandle = (Get-HandlesAndProcessIds -HandleInfo $handles | Measure-Object)
      
    # If file is open log an error
    If ( $openHandle.Count -gt 0) {
        Write-Output "ERROR - '$inputFile' is open"
    } Else {
        # Convert back slash in file path to forward slash per MySql requirements
        $cdrFile = $cdrPath -replace "\\", "/"
        $cdrFile = $cdrFile + "/" + $($_.name) -replace "\\", "/"
            
        # Open MySql connection and attempt to insert data
        Try {
            # Attempt to open MySql Connection
            Try {
                $conn.Open()
            } Catch [System.Exception] {
                Write-Verbose $_.Exception|format-list -force
                Write-Output $_.Exception.Message
                Write-Output "ERROR - Could not connect to database server '$mySqlhost'"
                Break
            }
            $cmd = New-Object MySql.Data.MySqlClient.MySqlCommand
            $cmd.Connection = $conn
            $cmd.CommandText = "LOAD DATA LOCAL INFILE '$cdrFile' INTO TABLE $dbName.$dbCdrTbl FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '""' LINES TERMINATED BY '\n' IGNORE 2 LINES (cdrRecordType, globalCallID_callManagerId, globalCallID_callId, origLegCallIdentifier, dateTimeOrigination, origNodeId, origSpan, origIpAddr, callingPartyNumber, callingPartyUnicodeLoginUserID, origCause_location, origCause_value, origPrecedenceLevel, origMediaTransportAddress_IP, origMediaTransportAddress_Port, origMediaCap_payloadCapability, origMediaCap_maxFramesPerPacket, origMediaCap_g723BitRate, origVideoCap_Codec, origVideoCap_Bandwidth, origVideoCap_Resolution, origVideoTransportAddress_IP, origVideoTransportAddress_Port, origRSVPAudioStat, origRSVPVideoStat, destLegIdentifier, destNodeId, destSpan, destIpAddr, originalCalledPartyNumber, finalCalledPartyNumber, finalCalledPartyUnicodeLoginUserID, destCause_location, destCause_value, destPrecedenceLevel, destMediaTransportAddress_IP, destMediaTransportAddress_Port, destMediaCap_payloadCapability, destMediaCap_maxFramesPerPacket, destMediaCap_g723BitRate, destVideoCap_Codec, destVideoCap_Bandwidth, destVideoCap_Resolution, destVideoTransportAddress_IP, destVideoTransportAddress_Port, destRSVPAudioStat, destRSVPVideoStat, dateTimeConnect, dateTimeDisconnect, lastRedirectDn, pkid, originalCalledPartyNumberPartition, callingPartyNumberPartition, finalCalledPartyNumberPartition, lastRedirectDnPartition, duration, origDeviceName, destDeviceName, origCallTerminationOnBehalfOf, destCallTerminationOnBehalfOf, origCalledPartyRedirectOnBehalfOf, lastRedirectRedirectOnBehalfOf, origCalledPartyRedirectReason, lastRedirectRedirectReason, destConversationId, globalCallId_ClusterID, joinOnBehalfOf, comment, authCodeDescription, authorizationLevel, clientMatterCode, origDTMFMethod, destDTMFMethod, callSecuredStatus, origConversationId, origMediaCap_Bandwidth, destMediaCap_Bandwidth, authorizationCodeValue, outpulsedCallingPartyNumber, outpulsedCalledPartyNumber, origIpv4v6Addr, destIpv4v6Addr, origVideoCap_Codec_Channel2, origVideoCap_Bandwidth_Channel2, origVideoCap_Resolution_Channel2, origVideoTransportAddress_IP_Channel2, origVideoTransportAddress_Port_Channel2, origVideoChannel_Role_Channel2, destVideoCap_Codec_Channel2, destVideoCap_Bandwidth_Channel2, destVideoCap_Resolution_Channel2, destVideoTransportAddress_IP_Channel2, destVideoTransportAddress_Port_Channel2, destVideoChannel_Role_Channel2, IncomingProtocolID, IncomingProtocolCallRef, OutgoingProtocolID, OutgoingProtocolCallRef, currentRoutingReason, origRoutingReason, lastRedirectingRoutingReason, huntPilotPartition, huntPilotDN, calledPartyPatternUsage, IncomingICID, IncomingOrigIOI, IncomingTermIOI, OutgoingICID, OutgoingOrigIOI, OutgoingTermIOI, outpulsedOriginalCalledPartyNumber, outpulsedLastRedirectingNumber, wasCallQueued, totalWaitTimeInQueue, callingPartyNumber_uri, originalCalledPartyNumber_uri, finalCalledPartyNumber_uri, lastRedirectDn_uri, mobileCallingPartyNumber, finalMobileCalledPartyNumber, origMobileDeviceName, destMobileDeviceName, origMobileCallDuration, destMobileCallDuration, mobileCallType, originalCalledPartyPattern, finalCalledPartyPattern, lastRedirectingPartyPattern, huntPilotPattern);"
            $rowsInserted = $cmd.ExecuteNonQuery()
            Write-Output "Parsed file '$cdrFile' - Added $rowsInserted CDR Records"
            $conn.Close()
            $totalRows = $totalRows + $rowsInserted
            $totalFiles = $totalFiles + 1
             
            # Move processed cdr file to archive folder
            Try {
                Move-Item -path $cdrFile -Destination $archivePath -Force
                Write-Output "Archived file '$cdrFile'"
            } Catch {
                Write-Verbose $_.Exception|format-list -force
                Write-Output $_.Exception.Message
                Write-Output "ERROR - Failed to archive file '$cdrFile'"
            }
        } Catch [System.Exception] {
            Write-Verbose $_.Exception|format-list -force
            Write-Output $_.Exception.Message
            Write-Output "ERROR - Failed to insert file '$cdrFile'"
        }
    }           
}

Write-Output "Total Rows Inserted: $totalRows"
Write-Output "Total Files Processed: $totalFiles"

Try {
    Stop-Transcript
} Catch {
}
