<#.SYNOPSIS
  Script para enumerar los objetos usuarios con >90 días sin validarse.
  DESCRIPTION
  Using get-Aduser obtain LastLogon properties to validate and report users object with 90 days or more without logon.
#>
# Ruta del fichero HTML de salida
$rutaSalida = "C:\Scripts\AccountsInactive90daysandSensitiveInfoInDescription.html"

# ================================
#   INFORME 1: CUENTAS INACTIVAS
# ================================
$usuariosInactivos = Get-ADUser -Filter * -Properties LastLogonDate |
    Where-Object {
        $_.LastLogonDate -le (Get-Date).AddDays(-90) -and $_.LastLogonDate -ne $null
    } |
    Select-Object Name, SamAccountName, LastLogonDate |
    Sort-Object LastLogonDate

# ================================
#   INFORME 2: CUENTAS CON PALABRAS SENSIBLES EN DESCRIPCIÓN
# ================================
$usuariosDescripcion = Get-ADUser -Filter * -Properties Description, LastLogonDate |
    Where-Object {
        $_.Description -match "(?i)contraseña|contraseñas|contrasena|contrasenas|credencial|credenciales|password|passwords|pwd|pdw"
    } |
    Select-Object Name, SamAccountName, Description, LastLogonDate |
    Sort-Object Name

# ================================
#   CABECERA HTML + ESTILOS
# ================================
$htmlHeader = @"
<html>
<head>
    <meta charset="UTF-8">
    <title>Informe de Active Directory - Cuentas inactivas y con información sensible</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            color: black;
            background-color: white;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 40px;
        }
        th {
            background-color: #cccccc;
            color: black;
            padding: 8px;
            border: 1px solid #999999;
            text-align: left;
        }
        td {
            padding: 6px 8px;
            border: 1px solid #dddddd;
            color: black;
        }
        tr:nth-child(odd) {
            background-color: white;
        }
        tr:nth-child(even) {
            background-color: #ccffcc;
        }
        h2 {
            margin-top: 40px;
        }
    </style>
</head>
<body>
    <h1>Informe de Active Directory - Cuentas inactivas y con información sensible</h1>
"@

# ================================
#   TABLA 1: CUENTAS INACTIVAS
# ================================
$htmlInactivos = @"
    <h2>Cuentas que no han iniciado sesión en 90 días o más</h2>
    <table>
        <tr>
            <th>Nombre</th>
            <th>SamAccountName</th>
            <th>Último inicio de sesión</th>
        </tr>
"@

foreach ($u in $usuariosInactivos) {
    $fecha = if ($u.LastLogonDate) { $u.LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "" }
    $htmlInactivos += "<tr><td>$($u.Name)</td><td>$($u.SamAccountName)</td><td>$fecha</td></tr>`r`n"
}

$htmlInactivos += "</table>"

# ================================
#   TABLA 2: CUENTAS CON PALABRAS SENSIBLES
# ================================
$htmlDescripcion = @"
    <h2>Cuentas cuyo campo Descripción contiene palabras sensibles</h2>
    <table>
        <tr>
            <th>Nombre</th>
            <th>SamAccountName</th>
            <th>Descripción</th>
            <th>Último inicio de sesión</th>
        </tr>
"@

foreach ($u in $usuariosDescripcion) {
    $fecha = if ($u.LastLogonDate) { $u.LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "" }
    $htmlDescripcion += "<tr><td>$($u.Name)</td><td>$($u.SamAccountName)</td><td>$($u.Description)</td><td>$fecha</td></tr>`r`n"
}

$htmlDescripcion += "</table>"

# ================================
#   PIE DEL HTML
# ================================
$htmlFooter = @"
</body>
</html>
"@

# ================================
#   GENERAR ARCHIVO FINAL
# ================================
$htmlCompleto = $htmlHeader + $htmlInactivos + $htmlDescripcion + $htmlFooter
$htmlCompleto | Out-File -FilePath $rutaSalida -Encoding UTF8

Write-Host "Informe generado en: $rutaSalida"
