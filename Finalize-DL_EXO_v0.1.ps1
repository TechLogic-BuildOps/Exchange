<# 

.DESCRIPTION 
        This PowerShell script is used for finalizing Distribution lists and Room lists on O365, including renaming placeholder DLs and configuring message acceptance settings.

.NOTES 
        Author: Avadhoot Dalavi

1) Renaming placeholder DL
2) Added AcceptMessagesOnlyFrom and AcceptMessagesOnlyFromSendersOrMembers  
3) Enabled logging

#>

$now = get-date -f "yyyy-MM-dd hh:mm:ss"
$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
$TranscriptLogPath = "C:\Temp\DLsmigrate\Log\Transcript\Finalize_DL-transcript_" + $nowfiledate + ".txt"

# Start transcript with custom path
Start-Transcript -Path $TranscriptLogPath

$whoweare = $ENV:USERNAME

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
#Start-Transcript -Outputdirectory "J:\PowershellEARLlogs\Transcript"
#Write-Output $ENV:USERNAME

$global:nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"


[System.GC]::Collect()

# Make Windows negotiate higher TLS version:
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Get-PSSession | Remove-PSSession
Connect-ExchangeOnline

# Path to the CSV file containing distribution group data
$CsvFilePath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv"
# Path for the log file
$LogFile = "C:\Temp\DLsmigrate\Log\Finalize_DL\Finalize_DL_" + $nowfiledate + ".txt"

# Path for the successfully migrated DLs CSV
$SuccessCsvPath = "C:\Temp\DLsmigrate\Export\Finalized_DLs\Finalized_DLs.csv"

# Array to store successfully migrated DL details
$MigratedDLs = @()

# Function to log messages
function Log-Message {
    param (
        [string]$Message
    )
    $Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $LogEntry = "$TimeStamp [$Type] $Message"
    Write-Host $LogEntry

    # Output the log entry to the log file
    Add-Content -Path $Logfile -Value $LogEntry
}

try {
    # Import CSV file
    $GroupsData = Import-Csv -Path $CsvFilePath
    Log-Message "CSV file imported successfully."
} catch {
    Log-Message "Failed to import CSV file. Error: $_" -Type "ERROR"
    exit
}

$MigratedDLs = @()

# Iterate through each row of the CSV file
foreach ($GroupData in $GroupsData) {
    # Check if the distribution group already exists
    $ExistingGroup = Get-DistributionGroup -Identity ("Cloud-" + $GroupData.DisplayName) -ErrorAction SilentlyContinue

    if ($ExistingGroup) {
        Log-Message "Distribution group $($ExistingGroup.DisplayName) exists. Processing further..."

        # Construct parameters for Set-DistributionGroup cmdlet
        $SetGroupParams = @{
            Identity                      = $ExistingGroup.DisplayName
            DisplayName                   = $GroupData.DisplayName.Replace("Cloud-", "")
            Name                          = $GroupData.Name.Replace("Cloud-", "")
            Alias                         = $GroupData.Alias.Replace("Cloud-", "") 
            EmailAddresses                = ($GroupData.EmailAddresses -split ',')
            HiddenFromAddressListsEnabled = [System.Convert]::ToBoolean($GroupData.HiddenFromAddressLists)
        }

        try {
            # Set additional properties for the distribution group
            Set-DistributionGroup @SetGroupParams -ErrorAction Stop
            Log-Message "Distribution group $($ExistingGroup.DisplayName) updated successfully to $($GroupData.DisplayName)."
            # Store the successfully migrated DL details
            $MigratedDLs += [PSCustomObject]@{
                DisplayName = $SetGroupParams.DisplayName
                Name        = $SetGroupParams.Name
                Identity    = $ExistingGroup.Identity
                Alias       = $SetGroupParams.Alias
             }

        } catch {
            Log-Message "Error updating distribution group $($ExistingGroup.DisplayName): $_" -Type "ERROR"
        }
    }
    else {
        Log-Message "Distribution group Cloud-$($GroupData.DisplayName) does not exist. Skipping update." 
    }
}
    
# Export the successfully renamed groups to CSV
if ($MigratedDLs.Count -gt 0) {
    $MigratedDLs | Export-Csv -Path $SuccessCsvPath -NoTypeInformation -Force
    Log-Message "Successfully migrated DLs exported to: $SuccessCsvPath"
}

Sleep -Seconds 10

# Path to the CSV file containing distribution group data
$Csvfile = "C:\Temp\DLsmigrate\Export\Export_DLAcceptMessagesFrom\Export_DLAcceptMessagesFrom.csv"
# Path for the log file
$LogFile = "C:\Temp\DLsmigrate\Log\Import_DLAcceptMessagesFrom\Import_DLAcceptMessagesFrom_" + $nowfiledate + ".txt"

try {
    # Import CSV file
    $GroupsData = Import-Csv -Path $Csvfile
    Log-Message "CSV file imported successfully."
} catch {
    Log-Message "Failed to import CSV file. Error: $_" -Type "ERROR"
    exit
}

# Iterate through each row of the CSV file
foreach ($GroupData in $GroupsData) {
    try {
        # Check if the distribution group already exists
        $ExistingGroup = Get-DistributionGroup -Identity $GroupData.DisplayName -ErrorAction SilentlyContinue

        if ($ExistingGroup) {
            Log-Message "Distribution group $($ExistingGroup.DisplayName) exists. Processing further..."

            # Handle AcceptMessagesOnlyFrom
            if ($GroupData.AcceptMessagesOnlyFrom) {
                $EmailAddresses = $GroupData.AcceptMessagesOnlyFrom -split ','
                foreach ($Email in $EmailAddresses) {
                    Set-DistributionGroup -Identity $GroupData.DisplayName -AcceptMessagesOnlyFrom @{Add = $Email}
                    Log-Message "Set AcceptMessagesOnlyFrom for $($GroupData.DisplayName)."
                }
            }

            # Handle AcceptMessagesOnlyFromSendersOrMembers
            if ($GroupData.AcceptMessagesOnlyFromSendersOrMembers) {
                $SenderAddresses = $GroupData.AcceptMessagesOnlyFromSendersOrMembers -split ','
                foreach ($Sender in $SenderAddresses) {
                    Set-DistributionGroup -Identity $GroupData.DisplayName -AcceptMessagesOnlyFromSendersOrMembers @{Add = $Sender}
                    Log-Message "Set AcceptMessagesOnlyFromSendersOrMembers for $($GroupData.DisplayName)."
                }
            }
        } else {
            Log-Message "Distribution group $($GroupData.DisplayName) does not exist. Skipping update." 
        }
    } catch {
        Log-Message "Error updating distribution group $($GroupData.DisplayName): $_" -Type "ERROR"
    }

}

Disconnect-ExchangeOnline

Log-Message "Disconnected from Exchange onlinemanagement module."

Stop-Transcript
