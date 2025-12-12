$now = get-date -f "yyyy-MM-dd hh:mm:ss"
$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
$TranscriptLogPath = "C:\Temp\DLsmigrate\Log\Transcript\DL_Migration_Master-transcript_" + $nowfiledate + ".txt"

$CsvFilePath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv"
$SuccessCsvPath = "C:\Temp\DLsmigrate\Export\Finalized_DLs\Finalized_DLs.csv"
$TempCsvPath = "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Pending_DL_Properties.csv"

# Start transcript with custom path
Start-Transcript -Path $TranscriptLogPath

# Define script paths
$ExportDLScript = "C:\Temp\DLsmigrate\Export-DL_01.ps1"
$CreatePlaceholderDLScript = "C:\Temp\DLsmigrate\Import-DL_EXO_v0.1.ps1"
$RemoveDLScript = "C:\Temp\DLsmigrate\Remove-DLs_0.1.ps1"
$FinalizeDLScript = "C:\Temp\DLsmigrate\Finalize-DL_EXO_v0.1.ps1"

# Run Export-DL_01.ps1
Write-Host "Executing Export-DL_01.ps1..."
& $ExportDLScript
Write-Host "Export-DL_01.ps1 execution completed. Waiting for 300 seconds..."
Start-Sleep -Seconds 300

# Run Placeholder DL creation script
Write-Host "Executing Import-DL_EXO_v0.1.ps1..."
& $CreatePlaceholderDLScript
Write-Host "Import-DL_EXO_v0.1.ps1 execution completed. Waiting for 300 seconds..."
Start-Sleep -Seconds 300

# Run Remove-DLs_0.1.ps1 with sleep intervals
Write-Host "Executing Remove-DLs_0.1.ps1.."
& $RemoveDLScript
Write-Host "Remove-DLs_0.1.ps1 execution completed. Waiting for 300 seconds..."
Start-Sleep -Seconds 300

# Wait for 30 minutes to allow sync completion
Write-Host "Waiting for 30 minutes for sync to complete..."
Start-Sleep -Seconds 2400

# Retry Logic for Finalizing DLs
$MaxRetries = 10
$RetryInterval = 300  # 5 minutes
$RetryCount = 0

while ($RetryCount -lt $MaxRetries) {
    # Filter only DLs that still need renaming
    if (Test-Path $SuccessCsvPath) {
        $AllDLs = Import-Csv -Path $CsvFilePath
        $RenamedDLs = Import-Csv -Path $SuccessCsvPath
        $PendingDLs = $AllDLs | Where-Object { $_.DisplayName -notin $RenamedDLs.DisplayName }
        
        if ($PendingDLs.Count -eq 0) {
            Write-Host "All DLs have been renamed successfully. Exiting..."
            $now = get-date -f "yyyy-MM-dd hh:mm:ss"
            $nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
            Rename-Item -Path "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv" -NewName ("C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties_" + $nowfiledate + ".csv")
            Rename-Item -Path "C:\Temp\DLsmigrate\Export\Finalized_DLs\Finalized_DLs.csv" -NewName ("C:\Temp\DLsmigrate\Export\Finalized_DLs\Finalized_DLs_" + $nowfiledate + ".csv")
            Rename-Item -Path "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Pending_DL_Properties.csv" -NewName ("C:\Temp\DLsmigrate\Export\Export_DL_Properties\Pending_DL_Properties_" + $nowfiledate + ".csv")
            Rename-Item -Path "C:\Temp\DLsmigrate\Export\Export_DLs_removed\Export_DLs_removed.csv" -NewName ("C:\Temp\DLsmigrate\Export\Export_DLs_removed\Export_DLs_removed_" + $nowfiledate + ".csv")
            Rename-Item -Path "C:\Temp\DLsmigrate\Export\Export_DLAcceptMessagesFrom\Export_DLAcceptMessagesFrom.csv" -NewName ("C:\Temp\DLsmigrate\Export\Export_DLAcceptMessagesFrom\Export_DLAcceptMessagesFrom_" + $nowfiledate + ".csv")
            Rename-Item -Path "C:\Temp\DLsmigrate\Export\Export_Owner_Co-Owner\Export_Owner_Co-Owner.csv" -NewName ("C:\Temp\DLsmigrate\Export\Export_Owner_Co-Owner\Export_Owner_Co-Owner_" + $nowfiledate + ".csv")
            break
        }

        # Save only pending DLs for retry
        $PendingDLs | Export-Csv -Path $TempCsvPath -NoTypeInformation -Force
    } else {
        # First attempt, all DLs need processing
        Copy-Item -Path $CsvFilePath -Destination $TempCsvPath -Force
    }

    Write-Host "Executing Finalize-DL_EXO_v0.1.ps1 - Attempt #$($RetryCount+1)..."
    & $FinalizeDLScript
    Start-Sleep -Seconds $RetryInterval
    $RetryCount++
}

#$now = get-date -f "yyyy-MM-dd hh:mm:ss"
#$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
#Rename-Item -Path "C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties.csv" -NewName ("C:\Temp\DLsmigrate\Export\Export_DL_Properties\Export_DL_Properties_" + $nowfiledate + ".csv")

# Stop transcript logging
Stop-Transcript
