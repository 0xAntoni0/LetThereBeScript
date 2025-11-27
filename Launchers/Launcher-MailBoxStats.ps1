<#
    .SYNOPSIS
    Launcher script to download and run the latest script from GitHub securely.
    
    .DESCRIPTION
    Downloads the script content into memory first, then saves it as UTF-8 to avoid
    encoding issues (like "smart quotes" or special characters breaking).
    Then executes the local copy passing all arguments.
#>

[CmdletBinding()]
Param(
    [Parameter(ValueFromRemainingArguments = $true)]
    $ScriptParams
)

# 1. Configuration
$GitHubUrl = "https://raw.githubusercontent.com/0xAntoni0/LetThereBeScript/refs/heads/main/Ms365Mailbox_stats.ps1"
$LocalFile = "$env:TEMP\Ms365Mailbox_stats.ps1"

# 2. Setup Security Protocol (Required for GitHub TLS 1.2)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# 3. Download the latest version forcing UTF-8
Write-Host "Downloading latest version from GitHub..." -ForegroundColor Cyan
try {
    # Content downloaded in memory
    $ScriptContent = Invoke-RestMethod -Uri $GitHubUrl -UseBasicParsing -ErrorAction Stop
    
    # Force disk writting using UTF-8
    [System.IO.File]::WriteAllText($LocalFile, $ScriptContent, [System.Text.Encoding]::UTF8)
}
catch {
    Write-Error "Failed to download script from $GitHubUrl"
    Write-Error "Error details: $_"
    Exit 1
}

# 4. Execute the downloaded script
if (Test-Path $LocalFile) {
    Write-Host "Executing script..." -ForegroundColor Green
    
    # Run with @args arguments
    & $LocalFile @args
}
else {
    Write-Error "The script file could not be found at $LocalFile."
    Exit 1
}
