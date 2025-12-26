<#
    .SYNOPSIS
    Generates a detailed HTML Health Report for Active Directory and ADCS.

    .DESCRIPTION
    This script performs health checks on Domain Controllers including:
    - Infrastructure: Connectivity (Ping), Uptime, Time Sync, and Disk Space.
    - Services: DNS, NTDS, NetLogon status.
    - ADCS: Certificate Authority service status and Certificate expiration monitoring.
    - DCDIAG: Detailed analysis with tooltips explaining each test.
     
    The output is a categorized, localized (Spanish) HTML report for easy reading.

    .OUTPUTS
    HTML Report file (saved locally or sent via SMTP).
#>

[CmdletBinding()]
Param(
    [Parameter( Mandatory = $false)]
    [string]$DomainName,

    [Parameter( Mandatory = $false)]
    [switch]$ReportFile,

    [Parameter( Mandatory = $false)]
    [switch]$SendEmail
)

# ---------------------------------------------------------------------------
# DEFAULT BEHAVIOR
# ---------------------------------------------------------------------------
if (-not $ReportFile -and -not $SendEmail) {
    $ReportFile = $true
    Write-Verbose "No output parameters specified. Defaulting to -ReportFile."
}

#...................................
# Global Variables
#...................................

$allTestedDomainControllers = [System.Collections.Generic.List[Object]]::new()
$now = Get-Date
$date = $now.ToShortDateString()
$reportTime = $now
$reportFileNameTime = $now.ToString("yyyyMMdd_HHmmss")
$reportemailsubject = "Domain Controller Health Report"

$smtpsettings = @{
    To         = 'email@domain.com'
    From       = 'adhealth@yourdomain.com'
    Subject    = "$reportemailsubject - $date"
    SmtpServer = "mail.domain.com"
    Port       = "25"
}

#...................................
# Functions
#...................................

Function Get-AllDomains() {
    $allDomains = (Get-ADForest).Domains
    return $allDomains
}

Function Get-AllDomainControllers ($ComputerName) {
    $allDomainControllers = Get-ADDomainController -Filter * -Server $ComputerName | Sort-Object HostName
    return $allDomainControllers
}

Function Get-DomainControllerNSLookup($ComputerName) {
    try {
        $domainControllerNSLookupResult = Resolve-DnsName $ComputerName -Type A -ErrorAction Stop | Select-Object -ExpandProperty IPAddress -First 1
        $domainControllerNSLookupResult = 'Success'
    }
    catch {
        $domainControllerNSLookupResult = 'Fail'
    }
    return $domainControllerNSLookupResult
}

Function Get-DomainControllerPingStatus($ComputerName) {
    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) { return "Success" }
    return 'Fail'
}

Function Get-DomainControllerUpTime($ComputerName) {
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $W32OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
            $timespan = (Get-Date) - $W32OS.LastBootUpTime
            [int]$uptime = "{0:00}" -f $timespan.TotalHours
        }
        catch { $uptime = 'CIM Failure' }
    }
    else { $uptime = 'Fail' }
    return $uptime
}

function Get-TimeDifference($ComputerName) {
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $currentTime, $timeDifference = (& w32tm /stripchart /computer:$ComputerName /samples:1 /dataonly)[-1].Trim("s") -split ',\s*'
            $diff = [double]$timeDifference
            $diffRounded = [Math]::Round($diff, 1, [MidPointRounding]::AwayFromZero)
        }
        catch { $diffRounded = 'Fail' }
    }
    else { $diffRounded = 'Fail' }
    return $diffRounded
}

Function Get-DomainControllerServices($ComputerName) {
    $res = [PSCustomObject]@{ DNSService = $null; NTDSService = $null; NETLOGONService = $null }
    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        try {
            $services = Get-Service -ComputerName $ComputerName -Name DNS, NTDS, netlogon -ErrorAction SilentlyContinue
            if (($services | Where-Object {$_.Name -eq 'DNS'}).Status -eq 'Running') { $res.DNSService = 'Success' } else { $res.DNSService = 'Fail' }
            if (($services | Where-Object {$_.Name -eq 'NTDS'}).Status -eq 'Running') { $res.NTDSService = 'Success' } else { $res.NTDSService = 'Fail' }
            if (($services | Where-Object {$_.Name -eq 'netlogon'}).Status -eq 'Running') { $res.NETLOGONService = 'Success' } else { $res.NETLOGONService = 'Fail' }
        } catch { $res.DNSService = 'Fail'; $res.NTDSService = 'Fail'; $res.NETLOGONService = 'Fail' }
    } else { $res.DNSService = 'Fail'; $res.NTDSService = 'Fail'; $res.NETLOGONService = 'Fail' }
    return $res
}

Function Get-DomainControllerADCS($ComputerName) {
    $Result = [PSCustomObject]@{ Installed = $false; Status = "N/A"; MinCertDays = "N/A" }
    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        try {
            $RemoteCheck = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                $feature = Get-WindowsFeature -Name AD-Certificate -ErrorAction SilentlyContinue
                if ($feature.Installed) {
                    $svc = Get-Service -Name CertSvc -ErrorAction SilentlyContinue
                    $certs = Get-ChildItem Cert:\LocalMachine\My
                    $minDays = 9999
                    if ($certs) {
                        foreach ($c in $certs) {
                            $days = ($c.NotAfter - (Get-Date)).Days
                            if ($days -lt $minDays) { $minDays = $days }
                        }
                    } else { $minDays = "No Certs" }
                    return @{ IsInstalled = $true; SvcStatus = $svc.Status.ToString(); DaysLeft = $minDays }
                } else { return @{ IsInstalled = $false } }
            } -ErrorAction Stop
            if ($RemoteCheck.IsInstalled) {
                $Result.Installed = $true; $Result.Status = $RemoteCheck.SvcStatus; $Result.MinCertDays = $RemoteCheck.DaysLeft
            }
        } catch { $Result.Status = "WinRM Fail" }
    } else { $Result.Status = "Offline" }
    return $Result
}

Function Get-DomainControllerDCDiagTestResults($ComputerName) {
    $DCDiagTestResults = [PSCustomObject]@{
        ServerName = $ComputerName; Connectivity = $null; Advertising = $null; FrsEvent = $null; DFSREvent = $null;
        SysVolCheck = $null; KccEvent = $null; KnowsOfRoleHolders = $null; MachineAccount = $null; NCSecDesc = $null;
        NetLogons = $null; ObjectsReplicated = $null; Replications = $null; RidManager = $null; Services = $null;
        SystemLog = $null; VerifyReferences = $null; CheckSDRefDom = $null; CrossRefValidation = $null;
        LocatorCheck = $null; Intersite = $null; FSMOCheck = $null
    }
    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        $params = @("/s:$ComputerName", "/test:Connectivity", "/test:Advertising", "/test:FrsEvent", "/test:DFSREvent", "/test:SysVolCheck", "/test:KccEvent", "/test:KnowsOfRoleHolders", "/test:MachineAccount", "/test:NCSecDesc", "/test:NetLogons", "/test:ObjectsReplicated", "/test:Replications", "/test:RidManager", "/test:Services", "/test:SystemLog", "/test:VerifyReferences", "/test:CheckSDRefDom", "/test:CrossRefValidation", "/test:LocatorCheck", "/test:Intersite", "/test:FSMOCheck")
        try {
            $DCDiagTest = (Dcdiag.exe @params) -split ('[\r\n]')
            $TestName = $null; $TestStatus = $null
            $DCDiagTest | ForEach-Object {
                switch -Regex ($_) {
                    "Starting test:|Iniciando prueba:" { $TestName = ($_ -replace ".*Starting test:", "" -replace ".*Iniciando prueba:", "").Trim().Trim(".") }
                    "passed test|failed test|super. la prueba|no super. la prueba" { if ($_ -match "passed test" -or $_ -match "super. la prueba") { $TestStatus = "Passed" } else { $TestStatus = "Failed" } }
                }
                if ($TestName -and $TestStatus) { $DCDiagTestResults.$TestName = $TestStatus; $TestName = $null; $TestStatus = $null }
            }
        } catch { foreach ($property in $DCDiagTestResults.PSObject.Properties.Name) { if ($property -ne "ServerName") { $DCDiagTestResults.$property = "Access Fail" } } }
    } else { foreach ($property in $DCDiagTestResults.PSObject.Properties.Name) { if ($property -ne "ServerName") { $DCDiagTestResults.$property = "Failed" } } }
    return $DCDiagTestResults
}

Function Get-DomainControllerOSDriveFreeSpaceGB ($ComputerName) {
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $thisOSDriveLetter = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop).SystemDrive
            $thisOSDiskDrive = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='$thisOSDriveLetter'" -ErrorAction Stop
            $freeSpaceGB = [math]::Round($thisOSDiskDrive.FreeSpace / 1GB, 2)
        } catch { $freeSpaceGB = 'CIM Failure' }
    } else { $freeSpaceGB = 'Fail' }
    return $freeSpaceGB
}

# ---------------------------------------------------------------------------
# MAIN EXECUTION LOOP
# ---------------------------------------------------------------------------

# Define base path and date-specific folder (C:\Scripts\dd-MM-yyyy)
$basePath = "C:\Scripts"
$dateFolderName = $now.ToString("dd-MM-yyyy")
$targetFolder = Join-Path -Path $basePath -ChildPath $dateFolderName

# Create the full directory structure if it doesn't exist
if (-not (Test-Path -Path $targetFolder)) {
    Write-Host "Creating folder: $targetFolder" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
} else {
    Write-Host "Using existing folder: $targetFolder" -ForegroundColor Cyan
}

# Define filename based on input parameters
if (!($DomainName)) {
    Write-Host "No domain specified, using all domains in forest" -ForegroundColor Yellow
    $allDomains = Get-AllDomains
    $reportNameOnly = 'forest_health_report_' + (Get-ADForest).name + '_' + $reportFileNameTime + '.html'
} else {
    Write-Host "Domain name specified on cmdline" -ForegroundColor Cyan
    $allDomains = $DomainName
    $reportNameOnly = 'dc_health_report_' + $DomainName + '_' + $reportFileNameTime + '.html'
}

# Construct the full output path
$reportFileName = Join-Path -Path $targetFolder -ChildPath $reportNameOnly

foreach ($domain in $allDomains) {
    Write-Host "Testing domain" $domain -ForegroundColor Green
    $allDomainControllers = Get-AllDomainControllers $domain
    $totalDCs = $allDomainControllers.Count
    $currentDCNumber = 0

    foreach ($domainController in $allDomainControllers) {
        $currentDCNumber++
        Write-Host "Testing DC ($currentDCNumber of $totalDCs) $($domainController.HostName)" -ForegroundColor Cyan
        
        $DCDiagTestResults = Get-DomainControllerDCDiagTestResults $domainController.HostName
        $ServicesStatus = Get-DomainControllerServices $domainController.HostName
        $ADCSStatus = Get-DomainControllerADCS $domainController.HostName

        $thisDomainController = [PSCustomObject]@{
            Server = ($domainController.HostName).ToLower(); Site = $domainController.Site; "OS Version" = $domainController.OperatingSystem; "IPv4 Address" = $domainController.IPv4Address;
            "DNS" = Get-DomainControllerNSLookup $domainController.HostName; "Ping" = Get-DomainControllerPingStatus $domainController.HostName;
            "Uptime (hours)" = Get-DomainControllerUpTime $domainController.HostName; "OS Free Space (GB)" = Get-DomainControllerOSDriveFreeSpaceGB $domainController.HostName;
            "Time offset (seconds)" = Get-TimeDifference $domainController.HostName;
            "DNS Service" = $ServicesStatus.DNSService; "NTDS Service" = $ServicesStatus.NTDSService; "NetLogon Service" = $ServicesStatus.NETLOGONService;
            "ADCS Installed" = $ADCSStatus.Installed; "ADCS Service" = $ADCSStatus.Status; "ADCS Cert Days" = $ADCSStatus.MinCertDays;
            
            # Add DCDIAG results
            "DCDIAG: Connectivity" = $DCDiagTestResults.Connectivity; "DCDIAG: Advertising" = $DCDiagTestResults.Advertising
        }
        
        foreach ($prop in $DCDiagTestResults.PSObject.Properties) {
             if ($prop.Name -ne "ServerName") { $thisDomainController | Add-Member -MemberType NoteProperty -Name "DCDIAG: $($prop.Name)" -Value $prop.Value -Force }
        }
        
        $allTestedDomainControllers.Add($thisDomainController)
    }
}

# -----------------------------------------------------------------------------------------
# HTML GENERATION
# -----------------------------------------------------------------------------------------
$htmlhead = "<html><head><meta charset='UTF-8'>
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 14px; background-color: #f0f2f5; margin: 0; padding: 20px; }
    .report-container { width: 1000px; max-width: 100%; margin: 0 auto; background-color: #ffffff; padding: 30px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); border-radius: 8px; }
    h1 { font-size: 24px; color: #2C3E50; border-bottom: 2px solid #2C3E50; padding-bottom: 10px; }
    h3 { font-size: 16px; color: #555; margin-top: 5px; }
    h4 { font-size: 15px; color: #ffffff; background-color: #34495e; padding: 8px 12px; margin-top: 30px; margin-bottom: 0; border-radius: 5px 5px 0 0; }
    .section-title { color: #2980b9; margin-top: 30px; margin-bottom: 10px; font-size: 18px; border-bottom: 1px solid #eee; padding-bottom: 5px; }
    .table-responsive { width: 100%; overflow-x: auto; margin-bottom: 20px; border: 1px solid #ccc; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 100%; }
    th { background-color: #2C3E50; color: #ffffff; padding: 10px; text-align: left; font-weight: 600; white-space: nowrap; }
    
    /* TOOLTIPS STYLES */
    th[title], td[title] { cursor: help; }
    th[title]:hover { background-color: #1a252f; }

    td { padding: 8px; border-bottom: 1px solid #eee; color: #333; white-space: nowrap; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    td.pass { background-color: #d4edda; color: #155724; font-weight: bold; text-align: center; }
    td.warn { background-color: #fff3cd; color: #856404; font-weight: bold; text-align: center; }
    td.fail { background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center; }
    td.na { color: #aaa; text-align: center; font-style: italic; }
</style>
</head>
<body>
<div class='report-container'>
    <h1>Reporte de Salud: Controladores de Dominio</h1>
    <h3>Generado: $reportTime</h3>
    <h3>Bosque: $((Get-ADForest).Name)</h3>"

# ------------------- TABLE 1: INFRASTRUCTURE -------------------
$tableInfraHead = "<h2 class='section-title'>1. Infraestructura y Estado del Sistema</h2>
<div class='table-responsive'><table><tr>
<th title='Nombre del Servidor (Hostname). Es la identidad de la maquina en la red.'>Server</th>
<th title='Ubicacion logica en AD. Importante para saber si la replicacion esta optimizada por zonas geograficas.'>Site</th>
<th title='Version de Windows Server instalada. util para detectar sistemas obsoletos o sin parches.'>OS Version</th>
<th title='Direccion IP principal. Verifica que coincida con lo esperado en el DNS.'>IPv4</th>
<th title='Prueba de resolucion de nombres. ¿El servidor es capaz de resolver su propia IP a traves del DNS?' style='text-align:center'>DNS</th>
<th title='Prueba de vida (ICMP). ¿El servidor responde o esta totalmente desconectado/apagado?' style='text-align:center'>Ping</th>
<th title='Tiempo encendido sin reinicios. Si es muy alto, faltan parches; si es muy bajo, se reinicio hace poco.' style='text-align:center'>Uptime (H)</th>
<th title='Espacio libre en disco C:. Si llega a 0, el servidor dejara de funcionar y la base de datos AD se detendra.' style='text-align:center'>Free Space GB</th>
<th title='Diferencia de hora con quien ejecuta el script. Si varia mas de 5 min, la autenticacion (Kerberos) fallara.' style='text-align:center'>Time Offset</th>
</tr>"

$tableInfraRows = ""
foreach ($reportline in $allTestedDomainControllers) {
    $row = "<tr><td><b>$($reportline.Server)</b></td><td>$($reportline.Site)</td><td>$($reportline."OS Version")</td><td>$($reportline."IPv4 Address")</td>"
    
    $valDNS = if($reportline.DNS -eq 'Success'){"Correcto"}else{"Fallo"}; $clsDNS = if($valDNS -eq 'Correcto'){"pass"}else{"fail"}
    $row += "<td class='$clsDNS'>$valDNS</td>"

    $valPing = if($reportline.Ping -eq 'Success'){"Correcto"}else{"Fallo"}; $clsPing = if($valPing -eq 'Correcto'){"pass"}else{"fail"}
    $row += "<td class='$clsPing'>$valPing</td>"
    
    $h = $reportline."Uptime (hours)"
    if($h -eq "Fail" -or $h -eq "CIM Failure") { $row += "<td class='fail'>Error</td>" }
    elseif([int]$h -le 24){ $row += "<td class='warn'>$h</td>" }
    else { $row += "<td class='pass'>$h</td>" }
    
    $gb = $reportline."OS Free Space (GB)"
    if($gb -eq "Fail" -or $gb -eq "CIM Failure") { $row += "<td class='fail'>Error</td>" }
    elseif($gb -lt 10) { $row += "<td class='fail'>$gb</td>" }
    elseif($gb -lt 20) { $row += "<td class='warn'>$gb</td>" }
    else { $row += "<td class='pass'>$gb</td>" }

    $t = $reportline."Time offset (seconds)"
    if($t -eq "Fail") { $row += "<td class='fail'>Error</td>" }
    elseif($t -ge 2) { $row += "<td class='fail'>$t</td>" }
    else { $row += "<td class='pass'>$t</td>" }

    $row += "</tr>"
    $tableInfraRows += $row
}
$tableInfra = $tableInfraHead + $tableInfraRows + "</table></div>"

# ------------------- TABLE 2: SERVICES -------------------
$tableSvcHead = "<h2 class='section-title'>2. Estado de Servicios y Roles</h2>
<div class='table-responsive'><table><tr>
<th title='Nombre del Servidor'>Server</th>
<th title='Servicio DNS Server. Si se detiene, nadie podra encontrar recursos en la red.' style='text-align:center'>Svc DNS</th>
<th title='Servicio de Dominio (NTDS). Es el corazon de Active Directory; si para, no hay logins.' style='text-align:center'>Svc NTDS</th>
<th title='Servicio NetLogon. Mantiene el canal seguro y procesa las peticiones de inicio de sesion.' style='text-align:center'>Svc NetLogon</th>
<th title='Servicio de Certificados. Si se detiene, no se pueden emitir ni renovar certificados.' style='text-align:center'>ADCS Svc</th>
<th title='Cuenta atras para el certificado que caduca primero. ¡Evita que venza o los servicios seguros dejaran de funcionar!' style='text-align:center'>ADCS Cert (Days)</th>
</tr>"

$tableSvcRows = ""
foreach ($reportline in $allTestedDomainControllers) {
    $row = "<tr><td><b>$($reportline.Server)</b></td>"
    
    foreach($svc in @("DNS Service", "NTDS Service", "NetLogon Service")){
        $v = $reportline.$svc
        if($v -eq "Success") { $row += "<td class='pass'>Funcionando</td>" } else { $row += "<td class='fail'>Fallo</td>" }
    }

    if ($reportline."ADCS Installed") {
        $adcs = $reportline."ADCS Service"
        if($adcs -eq "Running") { $row += "<td class='pass'>Funcionando</td>" } 
        else { $row += "<td class='fail'>$adcs</td>" }

        $days = $reportline."ADCS Cert Days"
        if ($days -eq "No Certs" -or $days -eq "N/A") { $row += "<td class='warn'>Revisar Store</td>" }
        elseif ([int]$days -lt 30) { $row += "<td class='fail'>$days</td>" }
        elseif ([int]$days -lt 60) { $row += "<td class='warn'>$days</td>" }
        else { $row += "<td class='pass'>$days</td>" }
    } else {
        $row += "<td class='na'>No Instalado</td><td class='na'>-</td>"
    }
    $row += "</tr>"
    $tableSvcRows += $row
}
$tableSvc = $tableSvcHead + $tableSvcRows + "</table></div>"

# ------------------- TABLE 3: DCDIAG (WITH DYNAMIC TOOLTIPS) -------------------

# Dictionary for DCDIAG explanations
$DCDiagHelp = @{
    "Connectivity"       = "Comprobacion basica. Verifica que el servidor tiene DNS registrado y responde a llamadas LDAP y RPC."
    "Advertising"        = "Verifica si el servidor se esta anunciando a la red como un Controlador de Dominio valido."
    "FrsEvent"           = "Busca errores graves en el registro de eventos del servicio de replicacion de archivos (FRS)."
    "DFSREvent"          = "Busca errores en la replicacion moderna (DFSR). Si falla, la carpeta SYSVOL podria no estar sincronizada."
    "SysVolCheck"        = "Confirma que la carpeta SYSVOL (donde estan las GPOs y scripts) esta compartida y accesible."
    "KccEvent"           = "Revisa el Arquitecto de la red (KCC). Si falla, el mapa de replicacion entre servidores no se esta calculando bien."
    "KnowsOfRoleHolders" = "¿Sabe este servidor quienes son los Jefes (Maestros FSMO) del dominio?"
    "MachineAccount"     = "Comprueba que la cuenta de maquina del propio servidor es valida y segura."
    "NCSecDesc"          = "Verifica los permisos de seguridad en los objetos principales del Directorio Activo."
    "NetLogons"          = "Asegura que los permisos de inicio de sesion y replicacion son correctos."
    "ObjectsReplicated"  = "Verifica que ciertos objetos criticos y cuentas del sistema se han replicado correctamente."
    "Replications"       = "La prueba mas importante. Comprueba si la replicacion con otros servidores ha ocurrido sin errores recientemente."
    "RidManager"         = "Verifica la comunicacion con el maestro RID (necesario para crear nuevos objetos/usuarios)."
    "Services"           = "Comprueba que los servicios criticos (RPC, DNS, KDC, etc.) estan encendidos."
    "SystemLog"          = "Analiza el visor de eventos en busca de errores criticos del sistema en los ultimos 60 minutos."
    "VerifyReferences"   = "Comprueba que los objetos criticos del sistema tienen referencias correctas entre si."
    "CheckSDRefDom"      = "Verifica descriptores de seguridad internos del dominio."
    "CrossRefValidation" = "Valida que las referencias cruzadas entre dominios y sitios son coherentes."
    "LocatorCheck"       = "Verifica que este DC puede ser encontrado por clientes y otros servidores."
    "Intersite"          = "Comprueba la replicacion entre diferentes sitios geograficos (si existen)."
    "FSMOCheck"          = "Comprueba que puede contactar con los servidores que tienen roles FSMO."
}

$htmlDCDiagTable = "<h2 class='section-title'>3. Detalle de Pruebas DCDIAG</h2>"
foreach ($reportline in $allTestedDomainControllers) {
    $dcName = $reportline.Server.ToUpper()
    $htmlDCDiagTable += "<h4>CONTROLADOR: $dcName</h4><table><tr><th style='width:70%'>Prueba Ejecutada</th><th style='width:30%; text-align:center'>Estado</th></tr>"
    $properties = $reportline.PSObject.Properties | Sort-Object Name
    foreach ($property in $properties) {
        if ($property.Name -like "DCDIAG:*") {
            $testName = $property.Name -replace "DCDIAG:\s*", ""
            $testResult = $property.Value
            
            # Lookup tooltip logic
            $tip = $DCDiagHelp[$testName]
            if (-not $tip) { $tip = "Prueba estandar de diagnostico de directorio activo ($testName)." }

            if ([string]::IsNullOrWhiteSpace($testResult)) { $displayResult = "Sin Datos"; $colorClass = "fail" }
            else {
                switch ($testResult) { 
                    "Passed" { $displayResult = "Superado"; $colorClass = "pass" } 
                    "Failed" { $displayResult = "Fallo"; $colorClass = "fail" } 
                    "Access Fail" { $displayResult = "Error Acceso"; $colorClass = "fail" } 
                    default { $displayResult = $testResult; $colorClass = "warn" } 
                }
            }
            # Add tooltip title and underline style
            $htmlDCDiagTable += "<tr><td title='$tip' style='border-bottom:1px dotted #999; width:fit-content;'>$testName</td><td class='$colorClass'>$displayResult</td></tr>"
        }
    }
    $htmlDCDiagTable += "</table>"
}

$htmltail = "<p style='font-size:11px; color:#777; margin-top:20px; border-top:1px solid #eee; padding-top:10px;'>
* Pasa el raton por encima de los titulos de las tablas y los nombres de las pruebas DCDIAG para obtener ayuda.</p></div></body></html>"

$htmlreport = $htmlhead + $tableInfra + $tableSvc + $htmlDCDiagTable + $htmltail

if ($ReportFile) {
    Write-Host "Guardando informe en: $reportFileName" -ForegroundColor Cyan
    $htmlreport | Out-File $reportFileName -Encoding UTF8
    Invoke-Item $reportFileName
}

if ($SendEmail) {
    try {
        $htmlreport | Out-File $reportFileName -Encoding UTF8
        Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Attachments $reportFileName -Encoding ([System.Text.Encoding]::UTF8) -ErrorAction Stop
        Write-Host "Email sent successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to send email. Error: $_" -ForegroundColor Red
    }
}
