<#
.SYNOPSIS
  Script para validar las contraseñas de los usuarios de dominio.
.DESCRIPTION
  Obtain with "Get-credential" user and password to validate if the password is correct or not.
#> 
 
 # Load the required .NET assembly
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

# Prompt for credentials
$cred = Get-Credential

# Extract username and password
$usuario = $cred.UserName
$clave = $cred.GetNetworkCredential().Password

# Extract domain from username (assumes format usuario@dominio.local)
$dominio = $usuario.Split("@")[1]

# Create domain context using extracted domain
$context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $dominio)

# Validate credentials
$resultado = $context.ValidateCredentials($usuario, $clave, [System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)

# Output result
if ($resultado) {
    Write-Host "✅ The password is valid."
} else {
    Write-Host "❌ The password is incorrect or the account has restrictions."
}
