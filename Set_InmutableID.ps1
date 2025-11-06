# Verificar si el módulo AzureAD está instalado
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "El módulo AzureAD no está instalado. Instalando..." -ForegroundColor Yellow
    try {
        Install-Module -Name AzureAD -Force -Scope CurrentUser
    } catch {
        Write-Host "Error al instalar el módulo AzureAD. Verifica tu conexión a Internet o permisos." -ForegroundColor Red
        exit
    }
}

# Importar el módulo
Import-Module AzureAD

# Conectarse a AzureAD
Write-Host "Conectando a AzureAD..." -ForegroundColor Cyan
try {
    Connect-AzureAD
} catch {
    Write-Host "Error al conectar con AzureAD. Verifica tus credenciales." -ForegroundColor Red
    exit
}

# Solicitar datos del usuario
$localSamAccountName = Read-Host "Introduce el nombre de usuario en AD local (sAMAccountName)"
$cloudUPN = Read-Host "Introduce el UPN del usuario en Microsoft 365 (ej. usuario@dominio.com)"

# Obtener el ObjectGUID del usuario local
try {
    $adUser = Get-ADUser -Identity $localSamAccountName
    $guid = [System.Convert]::ToBase64String($adUser.ObjectGUID.ToByteArray())
    Write-Host "GUID convertido a ImmutableID: $guid" -ForegroundColor Green
} catch {
    Write-Host "No se pudo obtener el usuario local. Verifica el nombre." -ForegroundColor Red
    exit
}

# Obtener el usuario en AzureAD
try {
    $cloudUser = Get-AzureADUser -ObjectId $cloudUPN
} catch {
    Write-Host "No se encontró el usuario en AzureAD. Verifica el UPN." -ForegroundColor Red
    exit
}

# Establecer el ImmutableID
try {
    Set-AzureADUser -ObjectId $cloudUser.ObjectId -ImmutableId $guid
    Write-Host "✅ ImmutableID actualizado correctamente para $cloudUPN" -ForegroundColor Green
} catch {
    Write-Host "❌ Error al actualizar el ImmutableID. Verifica que el usuario no esté ya sincronizado." -ForegroundColor Red
}
