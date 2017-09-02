<#
Purpose: Import Cisco CallManager Call Management (CMR) Records in to MySql
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
# Set CMR file path
$cmrPath = "C:\Inetpub\ftproot\cdr"
# Set CMR file arhive path where files are moved after process
$archivePath = "C:\Inetpub\ftproot\cdr\_archive_batch"

# Set database information
$mySqlhost = "localhost"
$dbName = "call_detail_records"
$dbUsername = "database username"
$dbPassword = "database password"
$dbCmrTbl = "cmr"

#
# You should not need to change anything below this line
#

# Include function to determine if cmr files are open by another program
Try {
    . "$PSScriptRoot\Get-HandlesAndProcessIds.ps1"
} Catch [System.Exception] {
    Write-Verbose $_.Exception|format-list -force
    Write-Output $_.Exception.Message
}

$logFileName = "CMR_Import_" + [System.DateTime]::Now.ToString("M_d_yyyy") + ".log"
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
If (!(Test-Path -Path $cmrPath)) {
    Write-Output "ERROR - Invalid CMR file path '$cmrPath'" 
} ElseIf (!(Test-Path -Path $archivePath)) {
    Write-Output "ERROR - Invalid CDR archive file path '$archivePath'"
} Else {
    # If both paths are OK then we continue
}

# Loop through all files starting with prefix 'cmr_'

$totalRows = 0
$totalFiles = 0

Get-ChildItem $cmrPath -Filter cmr_* | `
Foreach-Object{

    #Write-Output "$cmrPath\$($_.name)"

    # Check if the file we are going to process is already open
    $inputFile = $cmrPath + "\" + $($_.name)

    $handles = & handle.exe -a -u -accepteula "$inputFile"
    $openHandle = (Get-HandlesAndProcessIds -HandleInfo $handles | Measure-Object)
      
    # If file is open log an error
    If ( $openHandle.Count -gt 0) {
        Write-Output "ERROR - '$inputFile' is open"
    } Else {
        # Convert back slash in file path to forward slash per MySql requirements
        $cmrFile = $cmrPath -replace "\\", "/"
        $cmrFile = $cmrFile + "/" + $($_.name) -replace "\\", "/"
            
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
            $cmd.CommandText = "LOAD DATA LOCAL INFILE '$cmrFile' INTO TABLE $dbName.$dbCmrTbl FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '""' LINES TERMINATED BY '\n' IGNORE 2 LINES (cdrRecordType, globalCallID_callManagerId, globalCallID_callId, nodeId, directoryNum, callIdentifier, dateTimeStamp, numberPacketsSent, numberOctetsSent, numberPacketsReceived, numberOctetsReceived, numberPacketsLost, jitter, latency, pkid, directoryNumPartition, globalCallId_ClusterID, deviceName, varVQMetrics, duration, videoContentType, videoDuration, numberVideoPacketsSent, numberVideoOctetsSent, numberVideoPacketsReceived, numberVideoOctetsReceived, numberVideoPacketsLost, videoAverageJitter, videoRoundTripTime, videoOneWayDelay, videoReceptionMetrics, videoTransmissionMetrics, videoContentType_channel2, videoDuration_channel2, numberVideoPacketsSent_channel2, numberVideoOctetsSent_channel2, numberVideoPacketsReceived_channel2, numberVideoOctetsReceived_channel2, numberVideoPacketsLost_channel2, videoAverageJitter_channel2, videoRoundTripTime_channel2, videoOneWayDelay_channel2, videoReceptionMetrics_channel2, videoTransmissionMetrics_channel2);"
            $rowsInserted = $cmd.ExecuteNonQuery()
            Write-Output "Parsed file '$cmrFile' - Added $rowsInserted CMR Records"
            $conn.Close()
            $totalRows = $totalRows + $rowsInserted
            $totalFiles = $totalFiles + 1
             
            # Move processed cdr file to archive folder
            Try {
                Move-Item -path $cmrFile -Destination $archivePath -Force
                Write-Output "Archived file '$cmrFile'"
            } Catch {
                Write-Verbose $_.Exception|format-list -force
                Write-Output $_.Exception.Message
                Write-Output "ERROR - Failed to archive file '$cmrFile'"
            }
        } Catch [System.Exception] {
            Write-Verbose $_.Exception|format-list -force
            Write-Output $_.Exception.Message
            Write-Output "ERROR - Failed to insert file '$cmrFile'"
        }
    }           
}

Write-Output "Total Rows Inserted: $totalRows"
Write-Output "Total Files Processed: $totalFiles"

Try {
    Stop-Transcript
} Catch {
}
