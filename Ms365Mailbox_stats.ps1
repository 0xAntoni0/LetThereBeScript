# Check and install ExchangeOnlineManagement module if not present
try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module is not installed. Installing..." -ForegroundColor Green
        Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
    } else {
        Write-Host "ExchangeOnlineManagement module is already installed."
    }

    Import-Module ExchangeOnlineManagement
} catch {
    Write-Error "Error installing or importing the module: $_"
    exit
}

# Prompt for user
Write-Host "🔑 Enter tenant admin user or exo admin user" -ForegroundColor Green
$user = Read-Host "➡️"
if ([string]::IsNullOrWhiteSpace($user)) {
    Write-Error "🚩No valid user was entered."
    exit
}

# Connect to Exchange Online using modern authentication
try {
    Connect-ExchangeOnline -UserPrincipalName $user -ShowBanner:$false
} catch {
    Write-Error "🚩Error connecting to Exchange Online: $_"
    exit
}

# Get export path in Documents
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$exportPath = [Environment]::GetFolderPath("MyDocuments") + "\Exchange_Mailbox_Stats_$timestamp.csv"

# Define minimum % to fetch mailboxes meeting or exceeding that storage usage
$MinPercent = Read-Host "✉️Enter minimum % of storage. All mailboxes exceeding this will be returned"

# Get mailboxes
$mailboxes = Get-Mailbox
$total =$mailboxes.Count
$counter = 0

$results = foreach ($mb in $mailboxes) {
    $counter++

    # Show progress bar
    Write-Progress -Activity "Processing mailboxes..." `
                   -Status "Proccessing: $($mb.DisplayName)" `
                   -PercentComplete (($counter / $total) * 100)

    $primarySmtp = $mb.PrimarySmtpAddress.ToString()
    $stats = Get-MailboxStatistics -Identity $primarySmtp
    $quotaRaw = (Get-Mailbox -Identity $primarySmtp).ProhibitSendQuota

    # Extract bytes from TotalItemSize
    $sizeBytes = ($stats.TotalItemSize.ToString() -split '[()]')[1] -replace '[^\d]'
    $sizeGB = [math]::Round($sizeBytes / 1GB, 2)

    # Tamaño del archivo online
    $hasArchive = ($mb.ArchiveStatus -eq "Active")

    if ($hasArchive) {
        $archiveStats = Get-MailboxStatistics -Identity $primarySmtp -Archive
        $archiveSizeBytes = ($archiveStats.TotalItemSize.ToString() -split '[()]')[1] -replace '[^\d]'
        $archiveSizeGB = [math]::Round($archiveSizeBytes / 1GB, 2)
    } else {
        $archiveSizeGB = 0
    }

    # Extract bytes from Quota
    $quotaBytes = ($quotaRaw.ToString() -split '[()]')[1] -replace '[^\d]'
    $quotaGB = [math]::Round($quotaBytes / 1GB, 2)

    # Avoid division by zero
    if ($quotaGB -eq 0) { continue }

    # Calculate usage percentage
    $usagePercent = [math]::Round(($sizeGB / $quotaGB) * 100, 2)

    # Filter if usage exceeds the value of $MinPercent
    if ($usagePercent -gt $MinPercent) {
        [PSCustomObject]@{
            DisplayName     = $mb.DisplayName
            TotalSizeGB     = $sizeGB
            QuotaGB         = $quotaGB
            UsagePercent    = "$usagePercent%"
            ArchiveSizeGB    = $archiveSizeGB
        }
    }
}

# Export to CSV
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "✅ CSV file generated with mailboxes exceeding $MinPercent% usage at: $exportPath"
}
else {
Write-Host "⚠️ CSV export skipped — no mailboxes exceeded the specified threshold." -ForegroundColor Yellow
}
