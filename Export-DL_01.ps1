<#

.DESCRIPTION
		This PowerShell script can be used for exporting the co-owners of Distribution lsts and Rooms lists. Once this data is exported then script uses this export to capture all the different properties of DL like DisplayName, Name, EmailAddresses,GroupType,Members, Memberscount etc. for DL which has active Owner/Co-owner.

.NOTES
		Author: Avadhoot Dalavi

1) Export Owner/Co-owners of Distribution lists and Rooms lists. 
2) Capture all the different properties of DL like DisplayName, Name, EmailAddresses,GroupType,Members, Memberscount etc. for DL which has active Owner/Co-owner.

#>
$now = get-date -f "yyyy-MM-dd hh:mm:ss"
$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
$TranscriptLogPath = "C:\Temp\DLsmigrate\Log\Transcript\Export_DL-transcript_" + $nowfiledate + ".txt"

# Start transcript with custom path
Start-Transcript -Path $TranscriptLogPath

[System.GC]::Collect()

# Make Windows negotiate higher TLS version:
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

ConnectExchangeonPrem

Import-Module ActiveDirectory

# Define the log file path
$logFile = "C:\Temp\DLsmigrate\Log\Export_Owner_Co-Owner\Export_Owner_Co-Owner_" + $nowfiledate + ".txt"

# Define the output CSV file path
$exportFile = "C:\Temp\DLsmigrate\Export\Export_Owner_Co-Owner\Export_Owner_Co-Owner.csv"

# Function to log messages to a file
function Log-Message {
    param (
        [string]$Message
        
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $logFile -Value $LogEntry
}

# Start logging
Log-Message "Capturing Owner and Co-owners script execution started."

# Path to your CSV file
$csvPath = "C:\Temp\DLsmigrate\Import\TestDL.csv" # CSV file should have column name as PrimarySmtpAddress.

# Import the CSV file
$groups = Import-Csv -Path $csvPath

# Initialize an empty array to store all results
$results = @()

# Loop through each row in the CSV
foreach ($groupRow in $groups) {
    # Get the group's primary email address from the CSV
    $primaryEmail = $groupRow.Primarysmtpaddress

    # Get the group object
    Log-Message "Attempting to retrieve group with email: $primaryEmail"
    $group = Get-ADGroup -Filter { mail -eq $primaryEmail } -Properties mail, DistinguishedName, ManagedBy, proxyAddresses

    # Check if group is found
    if (-not $group) {
        Log-Message "Group with email $primaryEmail not found. Skipping to next group." -Level "ERROR"
        continue
    }
    Log-Message "Successfully retrieved group: $($group.Name)"

    # Get the group's Distinguished Name
    $grpDN = $group.DistinguishedName

    # Retrieve the ACL for the group
    Log-Message "Retrieving ACL for group: $($group.Name)"
    try {
        $acl = Get-ACL -Path "AD:\$($grpDN)"
    }
    catch {
        Log-Message "Failed to retrieve ACL for group: $($group.Name). Error: $_" -Level "ERROR"
        continue
    }

    # Filter ACL entries for co-owners
    $outdlACL = $acl.Access | Where-Object {
        ($_.IsInherited -ne $true) -and
        ($_.IdentityReference.Value -like "Domain\*") -and    #Domain should be replaced with actual domain name#
        ($_.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::GenericWrite)
    } | Select-Object -ExpandProperty IdentityReference

    $countofacls = $outdlACL.Count
    Log-Message "Group '$($group.Name)' has $countofacls CoOwner(s)."

    # Get the primary owner from ManagedBy attribute
    if ($group.ManagedBy) {
        Log-Message "Retrieving primary owner information from ManagedBy attribute."
        try {
            $primaryOwner = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName, SamAccountName, proxyAddresses
            $OwnerDisplayName = $primaryOwner.DisplayName
            $OwnerSamAccountName = $primaryOwner.SamAccountName
            $OwnerPrimarySMTPAddress = ($primaryOwner.proxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1) -replace '^SMTP:', ''
            if (-not $OwnerPrimarySMTPAddress) { $OwnerPrimarySMTPAddress = "NULL" }
            Log-Message "Primary owner retrieved: $OwnerDisplayName ($OwnerSamAccountName), SMTP: $OwnerPrimarySMTPAddress"
        }
        catch {
            Log-Message "Failed to resolve ManagedBy for group: $($group.Name). Error: $_" -Level "WARNING"
            $OwnerDisplayName = "N/A"
            $OwnerSamAccountName = "N/A"
            $OwnerPrimarySMTPAddress = "N/A"
        }
    }
    else {
        Log-Message "No primary owner set for group: $($group.Name)." -Level "WARNING"
        $OwnerDisplayName = "N/A"
        $OwnerSamAccountName = "N/A"
        $OwnerPrimarySMTPAddress = "N/A"
    }

    # If there are co-owners, export each co-owner separately
    if ($countofacls -gt 0) {
        foreach ($aclOwner in $outdlACL) {
            $aclOwnerSamAccount = ($aclOwner.Value -split '\\')[1]

            # Get the co-owner user object
            try {
                $coOwner = Get-ADUser -Identity $aclOwnerSamAccount -Properties DisplayName, SamAccountName, proxyAddresses
                $CoOwnerDisplayName = $coOwner.DisplayName
                $CoOwnerSamAccountName = $coOwner.SamAccountName
                $CoOwnerPrimarySMTPAddress = ($coOwner.proxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1) -replace '^SMTP:', ''
                if (-not $CoOwnerPrimarySMTPAddress) { $CoOwnerPrimarySMTPAddress = "NULL" }
            }
            catch {
                Log-Message "Failed to find user: $aclOwnerSamAccount. Error: $_" -Level "WARNING"
                $CoOwnerDisplayName = "N/A"
                $CoOwnerSamAccountName = "N/A"
                $CoOwnerPrimarySMTPAddress = "N/A"
            }

            # Add the record to the results array
            $results += [PSCustomObject]@{
                GroupName                 = $group.Name
                GroupDN                   = $grpDN
                Groupsamaccountname       = $group.SamAccountName
                GroupMail                 = $group.Mail
                GroupPrimarySMTPaddress = ($group.ProxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1) -replace '^SMTP:', ''
                OwnerDisplayName          = $OwnerDisplayName
                OwnerSamAccountName       = $OwnerSamAccountName
                OwnerPrimarySMTPaddress   = $OwnerPrimarySMTPAddress
                CoOwnerDisplayName        = $CoOwnerDisplayName
                CoOwnerSamAccountName     = $CoOwnerSamAccountName
                CoOwnerPrimarySMTPaddress = $CoOwnerPrimarySMTPAddress
            }
        }
    }
    else {
        # If no co-owners exist but there is an owner, still add a record
        $results += [PSCustomObject]@{
            GroupName                 = $group.Name
            GroupDN                   = $grpDN
            Groupsamaccountname       = $group.SamAccountName
            GroupMail                 = $group.Mail
            GroupPrimarySMTPaddress = ($group.ProxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1) -replace '^SMTP:', ''
            OwnerDisplayName          = $OwnerDisplayName
            OwnerSamAccountName       = $OwnerSamAccountName
            OwnerPrimarySMTPaddress   = $OwnerPrimarySMTPAddress
            CoOwnerDisplayName        = "N/A"
            CoOwnerSamAccountName     = "N/A"
            CoOwnerPrimarySMTPaddress = "N/A"
        }
    }
}

# Export the collected data to a CSV file
Log-Message "Exporting data to CSV file: $exportFile"
try {
    $results | Export-Csv -Path $exportFile -NoTypeInformation -Encoding UTF8
    Log-Message "Data successfully exported to $exportFile."
}
catch {
    Log-Message "Failed to export data to CSV. Error: $_" -Level "ERROR"
}

Log-Message "Script execution completed."

Log-Message "Capturing Owner and Co-owners script execution completed."

##################### DL Owner/Co-owner export is completed ############################

sleep -Seconds 10


# Define the path for the log file
$logFile = "C:\Temp\DLsmigrate\Log\Export_DL_Properties\Export_DL_Properties_" + $nowfiledate + ".txt"

# Define the paths for input and output CSV files
$inputCsvPath = "C:\Temp\DLsmigrate\Export\Export_Owner_Co-Owner\Export_Owner_Co-Owner.csv"
$outputCsvPath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv"

# Start processing
Log-Message "Starting with DL properties export which has active Owner or Co-owner"

try {
    # Import the input CSV and filter unique PrimarySMTPaddress values
    $inputData = Import-Csv -Path $inputCsvPath
    $uniqueDLs = $inputData | Where-Object {
        -not (
            ($_.OwnerPrimarySMTPaddress -eq "NULL" -or $_.OwnerPrimarySMTPaddress -eq "N/A") -and
            ($_.CoOwnerPrimarySMTPaddress -eq "NULL" -or $_.CoOwnerPrimarySMTPaddress -eq "N/A")
        )
    } | Select-Object -Unique GroupMail

    # Initialize an array to store export data
    $Output = @()

    foreach ($dl in $uniqueDLs) {
        $primarySMTPAddress = $dl.GroupMail

        try {
            # Retrieve the distribution group by PrimarySMTPaddress
            $group = Get-DistributionGroup -Identity $primarySMTPAddress -ErrorAction Stop
            $groupDN = $group.DistinguishedName

            Log-Message "Processing distribution group: $primarySMTPAddress"

            # Get members of the distribution group
            $members = Get-DistributionGroupMember -Identity $groupDN -ResultSize Unlimited

            # Fetch Primary SMTP Addresses for ManagedBy users
            $managedBySMTP = @()
            foreach ($manager in $group.ManagedBy.name) {
                $recipient = Get-Recipient -Identity $manager -ErrorAction SilentlyContinue
                if ($recipient) {
                    $managedBySMTP += $recipient.PrimarySmtpAddress
                }
            }

            # Collect DL properties
            $Output += [PSCustomObject]@{
                DisplayName                            = $group.DisplayName
                Name                                   = $group.Name
                PrimarySmtpAddress                     = $group.PrimarySmtpAddress
                EmailAddresses                         = ($group.EmailAddresses -join ',')
                Domain                                 = $group.PrimarySmtpAddress.ToString().Split("@")[1]
                Alias                                  = $group.Alias
                GroupType                              = $group.GroupType
                RecipientTypeDetails                   = $group.RecipientTypeDetails
                Members                                = $members.Name -join ','
                MembersPrimarySmtpAddress              = $members.PrimarySmtpAddress -join ','
                MemberCount                            = $members.Count
                ManagedBy                              = ($managedBySMTP -join ',') # Capturing ManagedBy's primary SMTP addresses
                HiddenFromAddressLists                 = $group.HiddenFromAddressListsEnabled
                MemberJoinRestriction                  = $group.MemberJoinRestriction
                MemberDepartRestriction                = $group.MemberDepartRestriction
                RequireSenderAuthenticationEnabled     = $group.RequireSenderAuthenticationEnabled
                AcceptMessagesOnlyFrom                 = ($group.AcceptMessagesOnlyFrom.Name -join ',')
                AcceptMessagesOnlyFromDLMembers        = ($group.AcceptMessagesOnlyFromDLMembers.Name -join ',')
                AcceptMessagesOnlyFromSendersOrMembers = ($group.AcceptMessagesOnlyFromSendersOrMembers.Name -join ',')
                ModeratedBy                            = ($group.ModeratedBy -join ',')
                BypassModerationFromSendersOrMembers   = ($group.BypassModerationFromSendersOrMembers.Name -join ',')
                ModerationEnabled                      = $group.ModerationEnabled
                SendModerationNotifications            = $group.SendModerationNotifications
                GrantSendOnBehalfTo                    = ($group.GrantSendOnBehalfTo.Name -join ',')
                Notes                                  = (Get-Group $groupDN).Notes
                MemberOf                               = (Get-Recipient -Filter "Members -eq '$groupDN'").PrimarySmtpAddress -join ','
                Description                            = (Get-ADGroup $groupDN -Properties description).description
                DistinguishedName                      = $group.DistinguishedName
            }

            Log-Message "Successfully processed distribution group: $primarySMTPAddress"
        }
        catch {
            Log-Message "Failed to process distribution group: $primarySMTPAddress. Error: $_"
        }
    }

    # Export all collected data to CSV
    $Output | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8

    Log-Message "Export completed successfully. Data saved to $outputCsvPath"
}
catch {
    Log-Message "An error occurred: $_"
}

Log-Message "DL properties are exported which has active Owner or Co-owner. Script execution is completed."

Get-PSSession | Remove-PSSession
Log-Message "Disconnected from Exchange On-prem."

# Stop transcript logging
Stop-Transcript