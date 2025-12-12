<#

.DESCRIPTION
		This PowerShell script can be used for removing Distribution lists and Room lists syncing from On-prem.

.NOTES
		Author: Avadhoot Dalavi

1) Removing On-prem DL
2) Logging enabled

#>

$now = get-date -f "yyyy-MM-dd hh:mm:ss"
$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
$TranscriptLogPath = "C:\Temp\DLsmigrate\Log\Transcript\Remove_DL-transcript_" + $nowfiledate + ".txt"

# Start transcript with custom path
Start-Transcript -Path $TranscriptLogPath

[System.GC]::Collect()

# Make Windows negotiate higher TLS version:
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


ConnectExchangeonPrem

# Path to CSV file with the list of DLs to remove
$csvPath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv"

# Path to the output CSV file for removed DLs
$DLs_removed = "C:\Temp\DLsmigrate\Export\Export_DLs_removed\Export_DLs_removed.csv"

# Path to the log file
$LogFile = "C:\Temp\DLsmigrate\Log\Export_DLs_removed\Export_DLs_removed_" + $nowfiledate + ".txt"

# Function to log messages
function Log-Message {
    param (
        [string]$Message
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$TimeStamp [$Type] $Message"
    Write-Host $LogEntry
    
    # Output the log entry to the log file
    Add-Content -Path $Logfile -Value $LogEntry
}

# Start logging
Log-Message "Starting the script to remove Distribution Lists (DLs)."

# Import the CSV file containing the list of DLs
try {
    $DLs = Import-Csv -Path $csvPath
    Log-Message "Successfully imported CSV file from $csvPath."
} catch {
    Log-Message "Failed to import CSV file. Error: $_" -Type "ERROR"
    exit
}

# Array to store successfully removed DLs
$removedDLs = @()

# Iterate through each row and remove the distribution group
foreach ($dl in $DLs) {
    # You need to use Primarysmtpaddress to remove the DL
    $primaryEmail = $dl.Primarysmtpaddress
    $displayName = $dl.displayName

    try {
        # Remove the On-prem synced Distribution Group
        Remove-DistributionGroup -Identity $primaryEmail -Confirm:$false
               
        # If successful, add to the removedDLs array
        $removedDLs += [PSCustomObject]@{
            DisplayName  = $displayName
            PrimarySMTPaddress = $primaryEmail
            Status       = "Removed"
        }

        Log-Message "Successfully removed: $displayName ($primaryEmail)."
    }
    catch {
        Log-Message "Failed to remove: $displayName ($primaryEmail). Error: $_" -Type "ERROR"
    }
}

# Export the removed DLs to a CSV file
if ($removedDLs.Count -gt 0) {
    try {
        $removedDLs | Export-Csv -Path $DLs_removed -NoTypeInformation
        Log-Message "Removed DLs written to: $DLs_removed."
    }
    catch {
        Log-Message "Failed to write removed DLs to $DLs_removed. Error: $_" -Type "ERROR"
    }
} else {
    Log-Message "No DLs were removed."
}

# End logging
Log-Message "DL's are removed. Script execution is completed."

Get-PSSession | Remove-PSSession
Log-Message "Disconnected from Exchange On-prem."

Stop-Transcript
