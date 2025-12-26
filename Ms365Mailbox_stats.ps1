<#
    .SYNOPSIS
    Script to export MailBox stats with Corporate Aesthetic HTML Report.
    Includes Conditional Formatting (Colors) based on Usage %.
#>

# Check and install ExchangeOnlineManagement module if not present
try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module is not installed. Installing..." -ForegroundColor Green
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    }

    # Import without clobbering to avoid errors if already loaded
    if (-not (Get-Module -Name ExchangeOnlineManagement)) {
        Import-Module ExchangeOnlineManagement
    }
} catch {
    Write-Error "Error installing or importing the module: $_"
    exit
}

# Connect to Exchange Online
try {
    $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $connection) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
} catch {
    Write-Error "Error connecting to Exchange Online: $_"
    exit
}

# --- HELPER FUNCTION FOR SIZE CONVERSION ---
function Get-ExSizeGB {
    param($Value)
    
    if ($null -eq $Value) { return 0 }
    
    # 1. If the object is "Live" (has the ToGB method), use it.
    if ($Value.PSObject.Methods.Name -contains "ToGB") {
        return [math]::Round($Value.ToGB(), 2)
    }
    
    # 2. If Deserialized (Text only), extract bytes using Regex
    # Typical format: "10 GB (10,737,418,240 bytes)"
    $str = $Value.ToString()
    if ($str -match "\(([\d,]+)\s*bytes\)") {
        try {
            # Clean commas and convert to number
            $bytes = [int64]($matches[1] -replace '[^\d]', '')
            return [math]::Round($bytes / 1GB, 2)
        } catch {
            return 0
        }
    }
    
    return 0
}
# -------------------------------------------

$timestamp = Get-Date -Format "yyyyMMdd_HHmm"

# Define base path and date-specific folder (C:\Scripts\yyyy-MM-dd)
$basePath = "C:\Scripts\Reportes"
$dateFolderName = $now.ToString("yyyy-MM-dd")
$targetFolder = Join-Path -Path $basePath -ChildPath $dateFolderName

# Create the full directory structure if it doesn't exist
if (-not (Test-Path -Path $targetFolder)) {
    Write-Host "Creating folder: $targetFolder" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
} else {
    Write-Host "Using existing folder: $targetFolder" -ForegroundColor Cyan
}

$exportPath = $basePath + "\Exchange_Mailbox_Stats_$timestamp.html"

$MinPercent = Read-Host "✉️ Enter minimum % of storage. All mailboxes exceeding this will be returned"

Write-Host "Fetching Mailbox list..." -ForegroundColor Cyan
$mailboxes = Get-Mailbox -ResultSize Unlimited
$total = $mailboxes.Count
$counter = 0

$results = foreach ($mb in $mailboxes) {
    $counter++

    Write-Progress -Activity "Scanning mailboxes..." `
                   -Status "[$counter / $total] Checking: $($mb.DisplayName)" `
                   -PercentComplete (($counter / $total) * 100)

    $primarySmtp = $mb.PrimarySmtpAddress.ToString()

    # --- ALIAS LOGIC ---
    # 1. Filter addresses starting with smtp: (or SMTP:)
    # 2. Remove 'smtp:' prefix to leave just the email
    # 3. Exclude primary address so only alternative aliases remain
    $aliasList = $mb.EmailAddresses | Where-Object { 
        $_ -match "^smtp:" 
    } | ForEach-Object {
        $_ -replace "^(?i)smtp:", "" # Remove 'smtp:' ignoring case
    } | Where-Object { 
        $_ -ne $primarySmtp 
    }

    # Join with <br> for separate lines in HTML
    $aliasString = $aliasList -join "<br>"
    if (-not $aliasString) { $aliasString = "-" } # If no alias, put a dash
    
    try {
        $stats = Get-MailboxStatistics -Identity $primarySmtp -ErrorAction Stop
    } catch {
        Write-Warning "Could not retrieve stats for $primarySmtp"
        continue
    }

    $sizeGB = Get-ExSizeGB -Value $stats.TotalItemSize.Value

    $archiveSizeGB = 0
    if ($mb.ArchiveStatus -eq "Active") {
        try {
            $archiveStats = Get-MailboxStatistics -Identity $primarySmtp -Archive -ErrorAction SilentlyContinue
            if ($archiveStats) {
                $archiveSizeGB = Get-ExSizeGB -Value $archiveStats.TotalItemSize.Value
            }
        } catch {
            $archiveSizeGB = 0
        }
    }

    $quotaRaw = if ($mb.ProhibitSendQuota.Value) { $mb.ProhibitSendQuota.Value } else { $mb.ProhibitSendQuota }
    
    if ("$quotaRaw" -eq "Unlimited") {
        $quotaGB = 100 
    } else {
        $quotaGB = Get-ExSizeGB -Value $quotaRaw
    }

    if ($quotaGB -eq 0) { continue }

    $usagePercent = [math]::Round(($sizeGB / $quotaGB) * 100, 2)

    if ($usagePercent -ge $MinPercent) {
        [PSCustomObject]@{
            DisplayName   = $mb.DisplayName
            Correo        = $primarySmtp
            Alias         = $aliasString # New property
            Tipo          = $mb.RecipientTypeDetails
            TotalSizeGB   = $sizeGB
            QuotaGB       = $quotaGB
            UsagePercent  = "$usagePercent%"
            ArchiveSizeGB = $archiveSizeGB
            RawPercent    = $usagePercent
        }
    }
}

# Export to HTML
if ($results.Count -gt 0) {
    
    # Sort by Usage Percent Descending
    $results = $results | Sort-Object -Property RawPercent -Descending

    $htmlHeader = @"
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informe de Uso de Buzones Exchange</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; background-color: #fff; margin: 20px; }
        h2, h3 { color: #2F4050; margin-bottom: 10px; margin-top: 20px; }
        h3 { font-size: 1.1em; font-weight: bold; border-bottom: 2px solid #2F4050; padding-bottom: 5px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 0.9em; }
        th { background-color: #2F4050; color: white; padding: 8px 10px; text-align: left; font-weight: bold; }
        td { padding: 8px 10px; border-bottom: 1px solid #eee; vertical-align: top; }
        /* Alternating color only for rows WITHOUT their own style */
        tr:not([style]):nth-child(even) { background-color: #f9f9f9; }
        .timestamp { font-size: 0.8em; color: #666; margin-bottom: 10px; }
        
        /* Alias Styles */
        .alias-cell { font-size: 0.85em; color: #555; } 
        /* FIX: If the row has a style (color), force alias to inherit that text color instead of being grey */
        tr[style] .alias-cell { color: inherit; }
    </style>
</head>
<body>
    <div class="timestamp">Generado: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</div>
    <h3>Informe de Uso de Almacenamiento (Umbral: $MinPercent%)</h3>
"@

    # Manual table construction
    $tableHtml = "<table>
        <thead>
            <tr>
                <th>Nombre</th>
                <th>Correo</th>
                <th>Alias</th> <!-- New Column -->
                <th>Tipo</th>
                <th>Tamaño (GB)</th>
                <th>Cuota (GB)</th>
                <th>Uso (%)</th>
                <th>Archivo (GB)</th>
            </tr>
        </thead>
        <tbody>"

    foreach ($row in $results) {
        $rowStyle = ""
        
        # Requested RGB colors converted to Hex
        # > 85%: Red Background (255,199,206) #FFC7CE | Red Text (156,0,6) #9C0006
        # >= 75%: Yellow Background (255,242,204) #FFF2CC | Ochre Text (156,87,0) #9C5700
        
        if ($row.RawPercent -gt 85) {
            $rowStyle = 'style="background-color: #FFC7CE; color: #9C0006; font-weight:bold;"'
        } elseif ($row.RawPercent -ge 75) {
            $rowStyle = 'style="background-color: #FFF2CC; color: #9C5700; font-weight:bold;"'
        }
        
        $tableHtml += "
            <tr $rowStyle>
                <td>$($row.DisplayName)</td>
                <td>$($row.Correo)</td>
                <td class='alias-cell'>$($row.Alias)</td>
                <td>$($row.Tipo)</td>
                <td>$($row.TotalSizeGB)</td>
                <td>$($row.QuotaGB)</td>
                <td>$($row.UsagePercent)</td>
                <td>$($row.ArchiveSizeGB)</td>
            </tr>"
    }

    $tableHtml += "</tbody></table>"

    $htmlFooter = @"
    <p style="font-size: 0.8em; color: #666;">Total de buzones reportados: $($results.Count)</p>
    <p style="font-size: 0.8em; color: #666;">*En rojo se colorean los buzones con el 85% o más de su capacidad ocupada</p>
    <p style="font-size: 0.8em; color: #666;">**En amarillo se colorean los buzones que se sitúen entre el 75% y el 85% de su capacidad ocupada</p>
    
</body>
</html>
"@

    $fullHtml = $htmlHeader + $tableHtml + $htmlFooter

    $fullHtml | Out-File -FilePath $exportPath -Encoding UTF8
    
    Write-Host "HTML report generated at: $exportPath" -ForegroundColor Green
    Start-Process $exportPath
    
} else {
    Write-Host "HTML export skipped — no mailboxes exceeded the specified threshold." -ForegroundColor Yellow
}
