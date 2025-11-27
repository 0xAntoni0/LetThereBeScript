<#
    .SYNOPSIS
    Launcher script to download and run the latest ADHealth.ps1 from GitHub.
#>

[CmdletBinding()]
Param(
    [Parameter(ValueFromRemainingArguments = $true)]
    $ScriptParams
)

# 1. Configuration
$GitHubUrl = "https://raw.githubusercontent.com/0xAntoni0/LetThereBeScript/refs/heads/main/Ms365Mailbox_stats.ps1"
$LocalFile = "$env:TEMP\Ms365Mailbox_stats.ps1"

# 2. Setup Security Protocol (Required for GitHub)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# 3. Download the latest version
Write-Host "Downloading latest version from GitHub..." -ForegroundColor Cyan
try {
    Invoke-RestMethod -Uri $GitHubUrl -OutFile $LocalFile -ErrorAction Stop
}
catch {
    Write-Error "Failed to download script. Check internet connection or URL."
    Write-Error $_
    Exit 1
}

# 4. Execute the downloaded script
# We use @args to pass any parameters used on this launcher directly to the downloaded script
Write-Host "Executing ADHealth check..." -ForegroundColor Green
& $LocalFile @args

# Optional: Clean up after execution
# Remove-Item $LocalFile -ErrorAction SilentlyContinue
