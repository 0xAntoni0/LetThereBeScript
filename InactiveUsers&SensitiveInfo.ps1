<#
.SYNOPSIS
  Script para enumerar los objetos usuarios con >90 días sin validarse.
.DESCRIPTION
  Using get-Aduser obtain LastLogon properties to validate and report users object with 90 days or more without logon.
#>

# --- CONFIGURATION ---
# Output HTML file path
$OutputPath = "C:\Scripts\AccountsInactive90daysandSensitiveInfoInDescription.html"

# --- DATA RETRIEVAL ---

Write-Host "Gathering inactive accounts (90+ days)..." -ForegroundColor Cyan
# Query 1: Users inactive for more than 90 days
# We include 'Enabled' to check account status
$InactiveUsers = Get-ADUser -Filter * -Properties LastLogonDate, Enabled |
    Where-Object {
        $_.LastLogonDate -le (Get-Date).AddDays(-90) -and $_.LastLogonDate -ne $null
    } |
    Select-Object Name, SamAccountName, LastLogonDate, Enabled |
    Sort-Object LastLogonDate

Write-Host "Gathering accounts with sensitive descriptions..." -ForegroundColor Cyan
# Query 2: Users with sensitive keywords in Description
# Regex '(?i)' makes it case-insensitive
$SensitiveUsers = Get-ADUser -Filter * -Properties Description, LastLogonDate, Enabled |
    Where-Object {
        $_.Description -match "(?i)contraseña|contraseñas|contrasena|contrasenas|credencial|credenciales|password|passwords|pwd|pdw"
    } |
    Select-Object Name, SamAccountName, Description, LastLogonDate, Enabled |
    Sort-Object Name

# --- HTML GENERATION ---

# Define CSS styles and HTML Header
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Active Directory Report - Inactive & Sensitive Accounts</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; color: #333; background-color: #f4f4f4; margin: 20px; }
        h1 { color: #2c3e50; }
        h2 { color: #2c3e50; margin-top: 40px; border-bottom: 2px solid #ddd; padding-bottom: 10px; }
        table { border-collapse: collapse; width: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.1); background-color: white; }
        th { background-color: #0078D7; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f1f1f1; }
        .status-disabled { color: #e74c3c; font-weight: bold; } /* Red */
        .status-active { color: #27ae60; font-weight: bold; }   /* Green */
        .footer { margin-top: 50px; font-size: 0.8em; color: #777; }
    </style>
</head>
<body>
    <h1>Active Directory Audit Report</h1>
"@

# --- TABLE 1: INACTIVE ACCOUNTS ---
$HtmlTable1 = @"
    <h2>Accounts inactive for 90 days or more</h2>
    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>SamAccountName</th>
                <th>Last Logon</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
"@

# Generate rows for Inactive Users
$RowsTable1 = foreach ($User in $InactiveUsers) {
    # Format Date
    $DateStr = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
    
    # Determine Status Class and Text
    $StatusClass = if ($User.Enabled) { "status-active" } else { "status-disabled" }
    $StatusText  = if ($User.Enabled) { "Habilitada" } else { "Deshabilitada" }

    "<tr>
        <td>$($User.Name)</td>
        <td>$($User.SamAccountName)</td>
        <td>$DateStr</td>
        <td class='$StatusClass'>$StatusText</td>
    </tr>"
}

$HtmlTable1 += ($RowsTable1 -join "`n") + "</tbody></table>"

# --- TABLE 2: SENSITIVE DESCRIPTIONS ---
$HtmlTable2 = @"
    <h2>Accounts with sensitive keywords in Description</h2>
    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>SamAccountName</th>
                <th>Description</th>
                <th>Last Logon</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
"@

# Generate rows for Sensitive Users
$RowsTable2 = foreach ($User in $SensitiveUsers) {
    # Format Date
    $DateStr = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "Never" }

    # Determine Status Class and Text
    $StatusClass = if ($User.Enabled) { "status-active" } else { "status-disabled" }
    $StatusText  = if ($User.Enabled) { "Habilitada" } else { "Deshabilitada" }

    "<tr>
        <td>$($User.Name)</td>
        <td>$($User.SamAccountName)</td>
        <td>$($User.Description)</td>
        <td>$DateStr</td>
        <td class='$StatusClass'>$StatusText</td>
    </tr>"
}

$HtmlTable2 += ($RowsTable2 -join "`n") + "</tbody></table>"

# --- FOOTER ---
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$HtmlFooter = @"
    <div class="footer">Generated on: $Timestamp</div>
</body>
</html>
"@

# --- EXPORT ---
Write-Host "Exporting to HTML..." -ForegroundColor Cyan

# Assemble final HTML content
$FinalHtml = $HtmlHeader + $HtmlTable1 + $HtmlTable2 + $HtmlFooter

# Write to file
try {
    $FinalHtml | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
    Write-Host "Success! Report saved to: $OutputPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to save the report: $_"
}
