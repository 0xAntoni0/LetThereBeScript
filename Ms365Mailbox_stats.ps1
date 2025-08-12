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
$exportPath = [Environment]::GetFolderPath("MyDocuments") + "\Exchange_Mailbox_Stats_$timestamp.html"

# Define minimum % to fetch mailboxes meeting or exceeding that storage usage
$MinPercent = Read-Host "✉️Enter minimum % of storage. All mailboxes exceeding this will be returned"

# Get mailboxes
$mailboxes = Get-Mailbox
$total = $mailboxes.Count
$counter = 0

$results = foreach ($mb in $mailboxes) {
    $counter++

    Write-Progress -Activity "Scanning mailboxes..." `
                   -Status "Now checking: $($mb.DisplayName)" `
                   -PercentComplete (($counter / $total) * 100)

    $primarySmtp = $mb.PrimarySmtpAddress.ToString()
    $stats = Get-MailboxStatistics -Identity $primarySmtp
    $quotaRaw = (Get-Mailbox -Identity $primarySmtp).ProhibitSendQuota

    $sizeBytes = ($stats.TotalItemSize.ToString() -split '[()]')[1] -replace '[^\d]'
    $sizeGB = [math]::Round($sizeBytes / 1GB, 2)

    $hasArchive = ($mb.ArchiveStatus -eq "Active")
    if ($hasArchive) {
        $archiveStats = Get-MailboxStatistics -Identity $primarySmtp -Archive
        $archiveSizeBytes = ($archiveStats.TotalItemSize.ToString() -split '[()]')[1] -replace '[^\d]'
        $archiveSizeGB = [math]::Round($archiveSizeBytes / 1GB, 2)
    } else {
        $archiveSizeGB = 0
    }

    $quotaBytes = ($quotaRaw.ToString() -split '[()]')[1] -replace '[^\d]'
    $quotaGB = [math]::Round($quotaBytes / 1GB, 2)

    if ($quotaGB -eq 0) { continue }

    $usagePercent = [math]::Round(($sizeGB / $quotaGB) * 100, 2)

    if ($usagePercent -gt $MinPercent) {
        [PSCustomObject]@{
            DisplayName      = $mb.DisplayName
            TotalSizeGB      = $sizeGB
            QuotaGB          = $quotaGB
            UsagePercent     = "$usagePercent%"
            ArchiveSizeGB    = $archiveSizeGB
        }
    }
}

# Export to HTML
if ($results.Count -gt 0) {
    $htmlHeader = @"
    <html>
    <head>
        <title>Exchange Mailbox Usage Report</title>
        <style>
            body { font-family: Arial; margin: 20px; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
        </style>
    </head>
    <body>
        <h2>📊 Exchange Mailbox Usage Report</h2>
        <p>Generated on: $(Get-Date)</p>
        <p>Threshold: $MinPercent%</p>
"@

    $htmlFooter = @"
    </body>
    </html>
"@

    $tableHtml = $results | ConvertTo-Html -Property DisplayName,TotalSizeGB,QuotaGB,UsagePercent,ArchiveSizeGB -Fragment
    $fullHtml = $htmlHeader + $tableHtml + $htmlFooter

    $fullHtml | Out-File -FilePath $exportPath -Encoding UTF8
    Write-Host "✅ HTML report generated at: $exportPath"
} else {
    Write-Host "⚠️ HTML export skipped — no mailboxes exceeded the specified threshold." -ForegroundColor Yellow
}
