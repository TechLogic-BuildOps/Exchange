<#

.DESCRIPTION
		This PowerShell script can be used for creating placeholder Distribution lists and Room lists.

.NOTES
		Author: Avadhoot Dalavi

1) Export AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromSendersOrMembers and AcceptMessagesOnlyFromDLMembers.
2) Create placeholder groups on EXO
3) Adding members in DL
4) Exports results to CSV with added members count
5) Assign Owners and Co-owners for Distribution lists and Room lists
6) Enabled logging

#>

$now = get-date -f "yyyy-MM-dd hh:mm:ss"
$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
$TranscriptLogPath = "C:\Temp\DLsmigrate\Log\Transcript\Import_DL-transcript_" + $nowfiledate + ".txt"

# Start transcript with custom path
Start-Transcript -Path $TranscriptLogPath

$whoweare = $ENV:USERNAME

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
$global:nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
[System.GC]::Collect()

# Make Windows negotiate higher TLS version:
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$now = get-date -f "yyyy-MM-dd hh:mm:ss"
$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
$logfilepath = "C:\Temp\DLsmigrate\Log\DLsMigrate-" + $nowfiledate + ".log"

Set-Variable -Name logfilepath -Value $logfilepath -Option ReadOnly -Scope Script -Force



Get-PSSession | Remove-PSSession
Connect-ExchangeOnline

# Define the log file
$Logfile = "C:\Temp\DLsmigrate\Log\Export_DLAcceptMessagesFrom\DLAcceptmessagesFrom_" + $nowfiledate + ".txt"

# Function to log messages
function Log-Message {
    param (
        [string]$Message
    )
    $Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $LogEntry = "$Timestamp [$Type] $Message"
    Write-Host $LogEntry
    Add-Content -Path $Logfile -Value $LogEntry
}

# Start logging
Log-Message "Starting script execution..."

# Import the CSV file
$CsvFilePath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv"
Log-Message "Importing CSV file: $CsvFilePath"
$DLList = Import-Csv -Path $CsvFilePath

# Prepare output array
$ExportData = @()



# Process each DL in the list
foreach ($DL in $DLList) {
    $PrimarySmtpAddress = $DL.PrimarySmtpAddress
    Log-Message "Fetching properties for DL: $PrimarySmtpAddress"

    try {
        # Get DL properties
        $DLProperties = Get-DistributionGroup -Identity $PrimarySmtpAddress
        
        # Resolve AcceptMessagesOnlyFrom
        $AcceptMessagesOnlyFrom = (Get-DistributionGroup -Identity $PrimarySmtpAddress).AcceptMessagesOnlyFrom | ForEach-Object {
            try {
                (Get-Recipient $_).PrimarySmtpAddress
            } catch {
                Log-Message "Failed to resolve PrimarySmtpAddress for $_" -Type "WARNING"
                $_  # Fallback to DN
            }
        }

        # Resolve AcceptMessagesOnlyFromDLMembers
        $AcceptMessagesOnlyFromDLMembers = (Get-DistributionGroup -Identity $PrimarySmtpAddress).AcceptMessagesOnlyFromDLMembers | ForEach-Object {
            try {
                (Get-Recipient $_).PrimarySmtpAddress
            } catch {
                Log-Message "Failed to resolve PrimarySmtpAddress for $_" -Type "WARNING"
                $_  # Fallback to DN
            }
        }

        # Resolve AcceptMessagesOnlyFromSendersOrMembers
        $AcceptMessagesOnlyFromSendersOrMembers = (Get-DistributionGroup -Identity $PrimarySmtpAddress).AcceptMessagesOnlyFromSendersOrMembers | ForEach-Object {
            try {
                (Get-Recipient $_).PrimarySmtpAddress
            } catch {
                Log-Message "Failed to resolve PrimarySmtpAddress for $_" -Type "WARNING"
                $_  # Fallback to DN
            }
        }


        # Add to export array
        $ExportData += [PSCustomObject]@{
            DisplayName                     = $DLProperties.DisplayName
            Name                            = $DLProperties.Name
            PrimarySmtpAddress              = $DLProperties.PrimarySmtpAddress
            AcceptMessagesOnlyFrom          = ($AcceptMessagesOnlyFrom -join ',')
            AcceptMessagesOnlyFromDLMembers = ($AcceptMessagesOnlyFromDLMembers -join ',')
            AcceptMessagesOnlyFromSendersOrMembers = ($AcceptMessagesOnlyFromSendersOrMembers -join ',')
        }
        Log-Message "Successfully fetched properties for $PrimarySmtpAddress"
    } catch {
        Log-Message "Error fetching properties for $PrimarySmtpAddress $_" -Type "ERROR"
    }
}

# Export the data to a CSV
$OutputFile = "C:\Temp\DLsmigrate\Export\Export_DLAcceptMessagesFrom\Export_DLAcceptMessagesFrom.csv"
Log-Message "Exporting data to CSV file: $OutputFile"
$ExportData | Export-Csv -Path $OutputFile -NoTypeInformation -Force

# End logging
Log-Message "AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromSendersOrMembers and AcceptMessagesOnlyFromDLMembers is successfully exported and script execution completed."

sleep -Seconds 10

# Define paths for the CSV file and log file
$CsvFilePath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv"
$Logfile = "C:\Temp\DLsmigrate\Log\Import_DL\Import_DL_" + $nowfiledate + ".txt"

# Log script start
Log-Message "Creating Place-holder DL's. Script execution started."

# Import CSV file
try {
    Log-Message "Importing CSV file from path: $CsvFilePath."
    $GroupsData = Import-Csv -Path $CsvFilePath
    Log-Message "CSV file imported successfully."
} catch {
    Log-Message "Error importing CSV file: $_" -LogLevel "ERROR"
    exit
}

# Iterate through each row of the CSV file
foreach ($GroupData in $GroupsData) {
    try {
        $GroupName = "Cloud-" + $GroupData.DisplayName
        Log-Message "Processing distribution group: $GroupName."

        # Check if the distribution group already exists
        $ExistingGroup = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue
        if ($ExistingGroup) {
            Log-Message "Distribution group $GroupName already exists." -LogLevel "WARNING"
        }
        else {
            # Handle different RecipientTypeDetails
            if ($GroupData.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                Log-Message "Creating Distribution Group: $GroupName."

                $NewGroupParams = @{
                    DisplayName        = $GroupName
                    Name               = "Cloud-" + $GroupData.Name
                    Alias              = "Cloud-" + $GroupData.Alias
                    PrimarySMTPAddress = "Cloud-" + $GroupData.PrimarySmtpAddress
                }

                # Create the new group
                $NewGroup = New-DistributionGroup @NewGroupParams
                Log-Message "Distribution group $GroupName created."

                # Set additional properties
                $SetGroupParams = @{
                    Identity                           = $NewGroup.DisplayName
                    HiddenFromAddressListsEnabled      = $True
                    MemberJoinRestriction              = "Closed"
                    MemberDepartRestriction            = "Closed"
                    RequireSenderAuthenticationEnabled = [System.Convert]::ToBoolean($GroupData.RequireSenderAuthenticationEnabled)
                    #ManagedBy                          = $GroupData.ManagedBy -split ','
                }

                Set-DistributionGroup @SetGroupParams
                Log-Message "Common Properties like Displayname, Alias, PrimarySMTPAddress, HiddenFromAddressListsEnabled, MemberJoin/DepartRestriction, RequireSenderAuthenticationEnabled is set for distribution group $GroupName."

                # Update Notes field using Set-Group
                if ($GroupData.Notes) {
                    try {
                        $SetGroupParams1 = @{
                            Identity = $NewGroup.DisplayName
                            Notes    = $GroupData.Notes
                            Description   = $GroupsData.description
                        }
                         if ($GroupData.Notes) {
                        Set-Group -Identity $NewGroup.DisplayName -Notes $GroupData.Notes
                        }
                        if ($GroupData.Description) {
                        Set-Group -Identity $NewGroup.DisplayName -Description $GroupData.Description
                        }
                        Log-Message "Notes and Description field updated for $GroupName."
                    } catch {
                        Log-Message "Failed to update Notes and Description for $GroupName. Error: $_" -LogLevel "ERROR"
                    }
                } else {
                    Log-Message "No Notes or Description found in CSV for $GroupName. Skipping Notes update."
                }
        
            } elseif ($GroupData.RecipientTypeDetails -eq "RoomList") {
                Log-Message "Creating RoomList: $GroupName."

                $NewGroupParams = @{
                    DisplayName        = $GroupName
                    Name               = "Cloud-" + $GroupData.Name
                    Alias              = "Cloud-" + $GroupData.Alias
                    PrimarySMTPAddress = "Cloud-" + $GroupData.PrimarySmtpAddress
                    Roomlist           = $True
                }

                # Create the new group
                $NewGroup = New-DistributionGroup @NewGroupParams
                Log-Message "RoomList $GroupName created."

                # Set additional properties
                $SetGroupParams = @{
                    Identity                           = $NewGroup.DisplayName
                    HiddenFromAddressListsEnabled      = $True
                    MemberJoinRestriction              = "Closed"
                    MemberDepartRestriction            = "Closed"
                    RequireSenderAuthenticationEnabled = [System.Convert]::ToBoolean($GroupData.RequireSenderAuthenticationEnabled)
                    #ManagedBy                          = $GroupData.ManagedBy -split ','
                }

                Set-DistributionGroup @SetGroupParams
                Log-Message "Common Properties like Displayname, Alias, PrimarySMTPAddress, HiddenFromAddressListsEnabled, MemberJoin/DepartRestriction, RequireSenderAuthenticationEnabled is set for RoomList $GroupName."
                                # Update Notes field using Set-Group
                if ($GroupData.Notes) {
                    try {
                        $SetGroupParams1 = @{
                            Identity = $NewGroup.DisplayName
                            Notes    = $GroupData.Notes
                            Description   = $GroupsData.description
                        }
                        if ($GroupData.Notes) {
                        Set-Group -Identity $NewGroup.DisplayName -Notes $GroupData.Notes
                        }
                        if ($GroupData.Description) {
                        Set-Group -Identity $NewGroup.DisplayName -Description $GroupData.Description
                        }

                        Log-Message "Notes and Description field updated for $GroupName."
                    } catch {
                        Log-Message "Failed to update Notes and Description for $GroupName. Error: $_" -LogLevel "ERROR"
                    }
                } else {
                    Log-Message "No Notes or Description found in CSV for $GroupName. Skipping Notes update."
                }

            }
        }
    } catch {
        Log-Message "Error processing group $GroupName $_" -LogLevel "ERROR"
    }
}

# Log script completion
Log-Message "Created Placeholder DL's. Script execution completed."

sleep -Seconds 10

# Define the log file path
$LogFile = "C:\Temp\DLsmigrate\Log\Add_Members\AddMembers_" + $nowfiledate + ".txt"

# Define the output CSV file path
$OutputCsvFile = "C:\Temp\DLsmigrate\Export\Add_Members\AddedMembers_" + $nowfiledate + ".csv"

# Path to the CSV file containing distribution group data
$CsvFilePath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv"

# Log the start of CSV import
Log-Message "Starting importing members from $CsvFilePath."

try {
    # Import CSV file
    $GroupsData = Import-Csv -Path $CsvFilePath
    Log-Message "CSV file imported successfully."
} catch {
    Log-Message "Failed to import CSV file. Error: $_"
    exit
}

# Initialize an array to store the results
$Results = @()

# Iterate through each row of the CSV file
foreach ($GroupData in $GroupsData) {
    Log-Message "Processing group: $($GroupData.PrimarySmtpAddress)."

    # Check if the distribution group already exists
    $ExistingGroup = Get-DistributionGroup -Identity ("Cloud-" + $GroupData.DisplayName) -ErrorAction SilentlyContinue

    if ($null -eq $ExistingGroup) {
        Log-Message "Distribution group $($GroupData.PrimarySmtpAddress) does not exist."
    }
    else {
        Log-Message "Distribution group $($ExistingGroup.PrimarySmtpAddress) found."

        # Initialize a counter for the number of members added
        $MembersAddedCount = 0

        # Check if MemberPrimarySmtpAddress is provided and not empty
        if (-not [string]::IsNullOrEmpty($GroupData.MembersPrimarySmtpAddress)) {
            Log-Message "MemberPrimarySmtpAddress found: $($GroupData.MembersPrimarySmtpAddress)"
            
            # Split the member email addresses if there are multiple addresses
            $Members = $GroupData.MembersPrimarySmtpAddress -split ","
            foreach ($Member in $Members) {
                # Trim whitespace from each member's email address
                $Member = $Member.Trim()

                Log-Message "Attempting to add member $Member to $($ExistingGroup.PrimarySmtpAddress)."
                
                # Add member to the distribution group
                try {
                    Add-DistributionGroupMember -Identity $ExistingGroup.PrimarySmtpAddress -Member $Member -BypassSecurityGroupManagerCheck -ErrorAction Stop
                    Log-Message "Member $Member added to $($ExistingGroup.PrimarySmtpAddress) successfully."

                    # Increment the count immediately after a successful addition
                    $MembersAddedCount++
                }
                catch {
                    Log-Message "Failed to add member $Member to $($ExistingGroup.PrimarySmtpAddress). Error: $_"
                }
            }

            # Add details to the results array
            $Results += [PSCustomObject]@{
                PrimarySmtpAddress        = $ExistingGroup.PrimarySmtpAddress
                MembersPrimarySmtpAddress = $GroupData.MembersPrimarySmtpAddress
                MembersAddedCount         = $MembersAddedCount
            }
        }
        else {
            Log-Message "No member provided for $($ExistingGroup.DisplayName)." -Type "WARNING"
        }
    }

    # Log completion of processing for the group
    Log-Message "Finished processing group: $($GroupData.DisplayName)."
}

# Export results to CSV
$Results | Export-Csv -Path $OutputCsvFile -NoTypeInformation

# Log script completion
Log-Message "Members are added to Place-holder DL's. Script execution is completed and results exported to $OutputCsvFile."

sleep -Seconds 10

# Path to the CSV file containing distribution group data
$CsvFilePath = "C:\Temp\DLsmigrate\Export\Export_Owner_Co-Owner\Export_Owner_Co-Owner.csv"
# Path for the log file
$LogFile = "C:\Temp\DLsmigrate\Log\Import_Owner_Co-Owner\Import_Owner_Co-Owner_" + $nowfiledate + ".txt"

# Function to log messages to a file
function Log-Message {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp [$LogLevel] $Message"
    Write-Host $LogEntry
    Add-Content -Path $Logfile -Value $LogEntry
}

try {
    # Import CSV file
    $GroupsData = Import-Csv -Path $Csvfilepath
    Log-Message "CSV file imported successfully."
} catch {
    Log-Message "Failed to import CSV file. Error: $_" -Type "ERROR"
       exit
}

# Start logging
Log-Message "Starting to add Owners and Co-owners. Script execution is started."


# Iterate through each row in the CSV file
foreach ($Group in $GroupsData) {
    $groupPrimarySMTP = $Group.GroupPrimarySMTPaddress
    $coOwnerSamAccountName = $Group.CoOwnerSamAccountName
    $coOwnerPrimarySMTP = $Group.CoOwnerPrimarySMTPaddress
    $ownerSamAccountName = $Group.OwnerSamAccountName
    $ownerPrimarySMTP = $Group.OwnerPrimarySMTPaddress

    try {
        $GroupName = "Cloud-" + $groupPrimarySMTP
        Log-Message "Processing distribution group: $GroupName."
        $group = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop

        # Resolve owner
        if ($ownerSamAccountName -and $ownerSamAccountName -ne "N/A" -and $ownerSamAccountName -ne "NULL") {
            $owner = Get-User -Identity $ownerSamAccountName -ErrorAction Stop
        } elseif ($ownerPrimarySMTP -and $ownerPrimarySMTP -ne "N/A" -and $ownerPrimarySMTP -ne "NULL") {
            $owner = Get-User -Identity $ownerPrimarySMTP -ErrorAction Stop
        } else {
            Log-Message "Skipping owner assignment due to missing data for group $GroupName"
            $owner = $null
        }

        # Resolve co-owner
        if ($coOwnerSamAccountName -and $coOwnerSamAccountName -ne "N/A" -and $coOwnerSamAccountName -ne "NULL") {
            $coOwner = Get-User -Identity $coOwnerSamAccountName -ErrorAction Stop
        } elseif ($coOwnerPrimarySMTP -and $coOwnerPrimarySMTP -ne "N/A" -and $coOwnerPrimarySMTP -ne "NULL") {
            $coOwner = Get-User -Identity $coOwnerPrimarySMTP -ErrorAction Stop
        } else {
            Log-Message "Skipping co-owner assignment due to missing data for group $GroupName"
            $coOwner = $null
        }

        $managedByList = @()

        if ($owner) {
            $managedByList += $owner.DistinguishedName
        }

        if ($coOwner) {
            $managedByList += $coOwner.DistinguishedName
        }

        if ($managedByList.Count -gt 0) {
            Set-DistributionGroup -Identity $group -ManagedBy @{Add = $managedByList} -ErrorAction Stop
            Log-Message "Successfully updated ManagedBy for $GroupName with: $($managedByList -join ', ')"
        } else {
            Log-Message "No valid owners found to add for group $GroupName"
        }

    }
    catch {
        Log-Message "Failed to update ManagedBy for distribution list '$groupPrimarySMTP'. Error: $_" "ERROR"
    }
}

Log-Message "Owners and Co-owners are added and Script is completed."

sleep -Seconds 10

Disconnect-ExchangeOnline

Log-Message "Disconnected from Exchange onlinemanagement module."

Stop-Transcript
