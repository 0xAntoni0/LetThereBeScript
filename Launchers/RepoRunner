# --- NETWORK CONFIGURATION ---
# Force TLS 1.2 usage (Required by GitHub API modern standards)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Define headers to act as a legitimate browser/client (User-Agent is mandatory)
$headers = @{ "User-Agent" = "PowerShellScript/1.0" }

# --- REPOSITORY DETAILS ---
$githubUser = "0xAntoni0"
$githubRepo = "LetThereBeScript"
$apiUrl = "https://api.github.com/repos/$githubUser/$githubRepo/contents"

Clear-Host
Write-Host "Conectando al repositorio..." -ForegroundColor Cyan

try {
    # Fetch file list using the API with the specific headers
    $allFiles = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
    
    # [CRITICAL FIX] @(...) Forces the result to be an array. 
    # This prevents errors when the repo contains only one single script.
    $scripts = @($allFiles | Where-Object { $_.name -like "*.ps1" } | Sort-Object name)

    if ($scripts.Count -eq 0) {
        Write-Warning "El repositorio existe, pero no hay archivos .ps1."
        exit
    }

    # --- MENU DISPLAY ---
    Write-Host "`n--- SCRIPTS DISPONIBLES ($($scripts.Count)) ---" -ForegroundColor Yellow
    
    # Loop through the array to generate the numbered list
    for ($i = 0; $i -lt $scripts.Count; $i++) {
        # Display index + 1 for user friendliness
        Write-Host "[$($i+1)] $($scripts[$i].name)"
    }

    Write-Host ""
    $selection = Read-Host "Introduce el número del script a ejecutar"

    # --- INPUT VALIDATION & EXECUTION ---
    # Check if the input contains only digits
    if ($selection -match "^\d+$") {
        
        # Explicitly cast string to integer for safe mathematical comparison
        $num = [int]$selection 
        
        # Verify the number is within the valid range of the array
        if ($num -gt 0 -and $num -le $scripts.Count) {
            
            # Select the script (Subtract 1 because array index starts at 0)
            $chosenScript = $scripts[$num - 1]
            
            Write-Host "`nLanzando: $($chosenScript.name)..." -ForegroundColor Green
            Write-Host "URL: $($chosenScript.download_url)" -ForegroundColor DarkGray
            
            # --- EXECUTION PAYLOAD ---
            # Download the raw content and execute immediately in memory
            Invoke-WebRequest -Uri $chosenScript.download_url -UseBasicParsing | Invoke-Expression
            
        } else {
            Write-Error "El número $num está fuera del rango (1 - $($scripts.Count))."
        }
    } else {
        Write-Error "Entrada inválida. Debes escribir un número."
    }

} catch {
    # Error handling for network or API issues
    Write-Error "Ocurrió un error:"
    Write-Error $_.Exception.Message
}
