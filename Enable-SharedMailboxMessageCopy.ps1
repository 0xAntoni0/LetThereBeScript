# Requirements: ExchangeOnlineManagement module installed

# 1. Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connection established successfully." -ForegroundColor Green
} catch {
    Write-Host "Critical connection error: $_" -ForegroundColor Red
    exit
}

# 2. Get Shared Mailboxes (Fast Mode)
Write-Host "`nRetrieving shared mailboxes (Fast Mode)..." -ForegroundColor Cyan
try {
    # Using Get-EXOMailbox (REST API) for better performance
    $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    $totalMailboxes = $sharedMailboxes.Count
    Write-Host "Found $totalMailboxes shared mailboxes." -ForegroundColor Yellow
} catch {
    Write-Host "Error retrieving mailboxes: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Check if any mailboxes were found
if ($totalMailboxes -eq 0) {
    Write-Host "No shared mailboxes found in the tenant." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# 3. User Confirmation
$confirmation = Read-Host "`nEnable 'MessageCopy' (SentAs/SendOnBehalf) for these $totalMailboxes mailboxes? (Y/N)"
if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# 4. Processing Loop
$counter = 0
$results = @()

foreach ($mailbox in $sharedMailboxes) {
    $counter++
    $percent = [math]::Round(($counter / $totalMailboxes) * 100, 2)
    
    # Progress Bar
    Write-Progress -Activity "Configuring Shared Mailboxes" -Status "Processing: $($mailbox.UserPrincipalName)" -PercentComplete $percent
    
    try {
        # Enabling both settings
        Set-Mailbox -Identity $mailbox.UserPrincipalName `
            -MessageCopyForSentAsEnabled $true `
            -MessageCopyForSendOnBehalfEnabled $true `
            -ErrorAction Stop
        
        $status = "Success"
        $errorMsg = ""
    } catch {
        $status = "Failed"
        $errorMsg = $_.Exception.Message
    }

    # Add result to list
    $results += [PSCustomObject]@{
        UPN = $mailbox.UserPrincipalName
        Status = $status
        ErrorDetails = $errorMsg
    }
}

# 5. Export Results and Cleanup
Write-Progress -Activity "Configuring Shared Mailboxes" -Completed

$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$csvPath = ".\SharedMailbox_Report_$timestamp.csv"

# Exporting to CSV (UTF8 ensures compatibility if names have special chars, though the script text is English)
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`n================ SUMMARY ================" -ForegroundColor Cyan
Write-Host "Process completed."
Write-Host "Report saved to: $csvPath" -ForegroundColor Yellow

Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
