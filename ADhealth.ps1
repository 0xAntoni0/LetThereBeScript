<#
    .SYNOPSIS
    Get-ADHealth.ps1 - Domain Controller Health Check Script.

    .DESCRIPTION
    This script performs a list of common health checks to a specific domain, or the entire forest. The results are then compiled into a colour coded HTML report.

    .OUTPUTS
    The results are currently only output to HTML for email or as an HTML report file, or sent as an SMTP message with an HTML body.

    .PARAMETER DomainName
    Perform a health check on a specific Active Directory domain.

    .PARAMETER ReportFile
    Output the report details to a file in the current directory.

    .PARAMETER SendEmail
    Send the report via email. You have to configure the correct SMTP settings.

    .EXAMPLE
    .\Get-ADHealth.ps1
    Checks all domains and creates a report automatically (Default behavior).
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
# If no output parameter is specified, enable -ReportFile automatically.
# ---------------------------------------------------------------------------
if (-not $ReportFile -and -not $SendEmail) {
    $ReportFile = $true
    Write-Verbose "No output parameters specified. Defaulting to -ReportFile."
}

#...................................
# Global Variables
#...................................

$allTestedDomainControllers = [System.Collections.Generic.List[Object]]::new()
$allDomainControllers = [System.Collections.Generic.List[Object]]::new()
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
    #Credential = (Get-Credential)
    #UseSsl     = $true
}

#...................................
# Functions
#...................................

Function Get-AllDomains() {
    Write-Verbose "Running function Get-AllDomains"
    $allDomains = (Get-ADForest).Domains
    return $allDomains
}

Function Get-AllDomainControllers ($ComputerName) {
    Write-Verbose "Running function Get-AllDomainControllers"
    $allDomainControllers = Get-ADDomainController -Filter * -Server $ComputerName | Sort-Object HostName
    return $allDomainControllers
}

Function Get-DomainControllerNSLookup($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerNSLookup"
    try {
        $domainControllerNSLookupResult = Resolve-DnsName $ComputerName -Type A | Select-Object -ExpandProperty IPAddress
        $domainControllerNSLookupResult = 'Success'
    }
    catch {
        $domainControllerNSLookupResult = 'Fail'
    }
    return $domainControllerNSLookupResult
}

Function Get-DomainControllerPingStatus($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerPingStatus"
    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        $domainControllerPingStatus = "Success"
    }
    else {
        $domainControllerPingStatus = 'Fail'
    }
    return $domainControllerPingStatus
}

Function Get-DomainControllerUpTime($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerUpTime"
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $W32OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction SilentlyContinue
            $timespan = (Get-Date) - $W32OS.LastBootUpTime
            [int]$uptime = "{0:00}" -f $timespan.TotalHours
        }
        catch {
            $uptime = 'CIM Failure'
        }
    }
    else {
        $uptime = 'Fail'
    }
    return $uptime
}

function Get-TimeDifference($ComputerName) {
    Write-Verbose "Running function Get-TimeDifference"
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $currentTime, $timeDifference = (& w32tm /stripchart /computer:$ComputerName /samples:1 /dataonly)[-1].Trim("s") -split ',\s*'
            $diff = [double]$timeDifference
            $diffRounded = [Math]::Round($diff, 1, [MidPointRounding]::AwayFromZero)
        }
        catch {
            $diffRounded = 'Fail'
        }
    }
    else {
        $diffRounded = 'Fail'
    }
    return $diffRounded
}

Function Get-DomainControllerServices($ComputerName) {
    Write-Verbose "Running function DomainControllerServices"
    $thisDomainControllerServicesTestResult = [PSCustomObject]@{
        DNSService      = $null
        NTDSService     = $null
        NETLOGONService = $null
    }

    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        if ((Get-Service -ComputerName $ComputerName -Name DNS -ErrorAction SilentlyContinue).Status -eq 'Running') {
            $thisDomainControllerServicesTestResult.DNSService = 'Success'
        }
        else {
            $thisDomainControllerServicesTestResult.DNSService = 'Fail'
        }
        if ((Get-Service -ComputerName $ComputerName -Name NTDS -ErrorAction SilentlyContinue).Status -eq 'Running') {
            $thisDomainControllerServicesTestResult.NTDSService = 'Success'
        }
        else {
            $thisDomainControllerServicesTestResult.NTDSService = 'Fail'
        }
        if ((Get-Service -ComputerName $ComputerName -Name netlogon -ErrorAction SilentlyContinue).Status -eq 'Running') {
            $thisDomainControllerServicesTestResult.NETLOGONService = 'Success'
        }
        else {
            $thisDomainControllerServicesTestResult.NETLOGONService = 'Fail'
        }
    }
    else {
        $thisDomainControllerServicesTestResult.DNSService = 'Fail'
        $thisDomainControllerServicesTestResult.NTDSService = 'Fail'
        $thisDomainControllerServicesTestResult.NETLOGONService = 'Fail'
    }
    return $thisDomainControllerServicesTestResult
}

Function Get-DomainControllerDCDiagTestResults($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerDCDiagTestResults"

    $DCDiagTestResults = [PSCustomObject]@{
        ServerName             = $ComputerName
        Connectivity           = $null
        Advertising            = $null
        FrsEvent               = $null
        DFSREvent              = $null
        SysVolCheck            = $null
        KccEvent               = $null
        KnowsOfRoleHolders     = $null
        MachineAccount         = $null
        NCSecDesc              = $null
        NetLogons              = $null
        ObjectsReplicated      = $null
        Replications           = $null
        RidManager             = $null
        Services               = $null
        SystemLog              = $null
        VerifyReferences       = $null
        CheckSDRefDom          = $null
        CrossRefValidation     = $null
        LocatorCheck           = $null
        Intersite              = $null
        FSMOCheck              = $null
    }

    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        $params = @(
            "/s:$ComputerName", "/test:Connectivity", "/test:Advertising", "/test:FrsEvent", "/test:DFSREvent",
            "/test:SysVolCheck", "/test:KccEvent", "/test:KnowsOfRoleHolders", "/test:MachineAccount", "/test:NCSecDesc",
            "/test:NetLogons", "/test:ObjectsReplicated", "/test:Replications", "/test:RidManager", "/test:Services",
            "/test:SystemLog", "/test:VerifyReferences", "/test:CheckSDRefDom", "/test:CrossRefValidation",
            "/test:LocatorCheck", "/test:Intersite", "/test:FSMOCheck"
        )

        try {
            $DCDiagTest = (Dcdiag.exe @params) -split ('[\r\n]')
            $TestName = $null
            $TestStatus = $null
    
            $DCDiagTest | ForEach-Object {
                switch -Regex ($_) {
                    # Detect test start (English and Spanish)
                    "Starting test:|Iniciando prueba:" {
                        $TestName = ($_ -replace ".*Starting test:", "" -replace ".*Iniciando prueba:", "").Trim().Trim(".")
                    }
                    # Detect result (English and Spanish)
                    "passed test|failed test|super. la prueba|no super. la prueba" {
                        if ($_ -match "passed test" -or $_ -match "super. la prueba") {
                            $TestStatus = "Passed"
                        } else {
                            $TestStatus = "Failed"
                        }
                    }
                }
                if ($TestName -and $TestStatus) {
                    $DCDiagTestResults.$TestName = $TestStatus
                    $TestName = $null
                    $TestStatus = $null
                }
            }
        }
        catch {
             # If execution fails, mark all DCDiag properties as Access Fail
             foreach ($property in $DCDiagTestResults.PSObject.Properties.Name) {
                if ($property -ne "ServerName") { $DCDiagTestResults.$property = "Access Fail" }
            }
        }
    }
    else {
        # If Ping fails, mark all as Failed
        foreach ($property in $DCDiagTestResults.PSObject.Properties.Name) {
            if ($property -ne "ServerName") {
                $DCDiagTestResults.$property = "Failed"
            }
        }
    }
    return $DCDiagTestResults
}

Function Get-DomainControllerOSDriveFreeSpace ($ComputerName) {
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $thisOSDriveLetter = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop).SystemDrive
            $thisOSDiskDrive = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='$thisOSDriveLetter'" -ErrorAction Stop
            $thisOSPercentFree = [math]::Round($thisOSDiskDrive.FreeSpace / $thisOSDiskDrive.Size * 100)
        }
        catch { $thisOSPercentFree = 'CIM Failure' }
    } else { $thisOSPercentFree = "Fail" }
    return $thisOSPercentFree
}

Function Get-DomainControllerOSDriveFreeSpaceGB ($ComputerName) {
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $thisOSDriveLetter = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop).SystemDrive
            $thisOSDiskDrive = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='$thisOSDriveLetter'" -ErrorAction Stop
            $freeSpaceGB = [math]::Round($thisOSDiskDrive.FreeSpace / 1GB, 2)
        }
        catch { $freeSpaceGB = 'CIM Failure' }
    } else { $freeSpaceGB = 'Fail' }
    return $freeSpaceGB
}

Function New-ServerHealthHTMLTableCell() {
    param( $lineitem )
    $htmltablecell = $null
    switch ($($reportline."$lineitem")) {
        "Success" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
        "Passed" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
        "Pass" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
        "Warn" { $htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>" }
        "Fail" { $htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>" }
        "Failed" { $htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>" }
        "Could not test server uptime." { $htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>" }
        default { $htmltablecell = "<td>$($reportline."$lineitem")</td>" }
    }
    return $htmltablecell
}

if (!($DomainName)) {
    Write-Host "No domain specified, using all domains in forest" -ForegroundColor Yellow
    $allDomains = Get-AllDomains
    $reportFileName = 'forest_health_report_' + (Get-ADForest).name + '_' + $reportFileNameTime + '.html'
}
else {
    Write-Host "Domain name specified on cmdline" -ForegroundColor Cyan
    $allDomains = $DomainName
    $reportFileName = 'dc_health_report_' + $DomainName + '_' + $reportFileNameTime + '.html'
}

foreach ($domain in $allDomains) {
    Write-Host "Testing domain" $domain -ForegroundColor Green
    $allDomainControllers = Get-AllDomainControllers $domain
    $allDomainControllers = @($allDomainControllers)
    $totalDCs = $allDomainControllers.Count
    $currentDCNumber = 0

    foreach ($domainController in $allDomainControllers) {
        $currentDCNumber++
        $stopWatch = [system.diagnostics.stopwatch]::StartNew()
        Write-Host "Testing domain controller ($currentDCNumber of $totalDCs) $($domainController.HostName)" -ForegroundColor Cyan
        $DCDiagTestResults = Get-DomainControllerDCDiagTestResults $domainController.HostName

        $thisDomainController = [PSCustomObject]@{
            Server = ($domainController.HostName).ToLower(); Site = $domainController.Site; "OS Version" = $domainController.OperatingSystem;
            "IPv4 Address" = $domainController.IPv4Address; "Operation Master Roles" = $domainController.OperationMasterRoles;
            "DNS" = Get-DomainControllerNSLookup $domainController.HostName; "Ping" = Get-DomainControllerPingStatus $domainController.HostName;
            "Uptime (hours)" = Get-DomainControllerUpTime $domainController.HostName; "OS Free Space (%)" = Get-DomainControllerOSDriveFreeSpace $domainController.HostName;
            "OS Free Space (GB)" = Get-DomainControllerOSDriveFreeSpaceGB $domainController.HostName; "Time offset (seconds)" = Get-TimeDifference $domainController.HostName;
            "DNS Service" = (Get-DomainControllerServices $domainController.HostName).DNSService; "NTDS Service" = (Get-DomainControllerServices $domainController.HostName).NTDSService;
            "NetLogon Service" = (Get-DomainControllerServices $domainController.HostName).NETLOGONService; "DCDIAG: Connectivity" = $DCDiagTestResults.Connectivity;
            "DCDIAG: Advertising" = $DCDiagTestResults.Advertising; "DCDIAG: FrsEvent" = $DCDiagTestResults.FrsEvent;
            "DCDIAG: DFSREvent" = $DCDiagTestResults.DFSREvent; "DCDIAG: SysVolCheck" = $DCDiagTestResults.SysVolCheck;
            "DCDIAG: KccEvent" = $DCDiagTestResults.KccEvent; "DCDIAG: FSMO KnowsOfRoleHolders" = $DCDiagTestResults.KnowsOfRoleHolders;
            "DCDIAG: MachineAccount" = $DCDiagTestResults.MachineAccount; "DCDIAG: NCSecDesc" = $DCDiagTestResults.NCSecDesc;
            "DCDIAG: NetLogons" = $DCDiagTestResults.NetLogons; "DCDIAG: ObjectsReplicated" = $DCDiagTestResults.ObjectsReplicated;
            "DCDIAG: Replications" = $DCDiagTestResults.Replications; "DCDIAG: RidManager" = $DCDiagTestResults.RidManager;
            "DCDIAG: Services" = $DCDiagTestResults.Services; "DCDIAG: SystemLog" = $DCDiagTestResults.SystemLog;
            "DCDIAG: VerifyReferences" = $DCDiagTestResults.VerifyReferences; "DCDIAG: CheckSDRefDom" = $DCDiagTestResults.CheckSDRefDom;
            "DCDIAG: CrossRefValidation" = $DCDiagTestResults.CrossRefValidation; "DCDIAG: LocatorCheck" = $DCDiagTestResults.LocatorCheck;
            "DCDIAG: Intersite" = $DCDiagTestResults.Intersite; "DCDIAG: FSMO Check" = $DCDiagTestResults.FSMOCheck;
            "Processing Time (seconds)" = $stopWatch.Elapsed.Seconds
        }
        $allTestedDomainControllers.Add($thisDomainController)
    }
}

# -----------------------------------------------------------------------------------------
# CSS STYLES (Fixed layout with Scroll)
# -----------------------------------------------------------------------------------------
$htmlhead = "<html><head><meta charset='UTF-8'>
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 14px; background-color: #f0f2f5; margin: 0; padding: 20px; }
    
    /* Main Container */
    .report-container {
        width: 1000px;
        max-width: 100%;
        margin: 0 auto;
        background-color: #ffffff;
        padding: 30px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        border-radius: 8px;
    }

    h1 { font-size: 24px; color: #2C3E50; margin-bottom: 5px; border-bottom: 2px solid #2C3E50; padding-bottom: 10px; }
    h3 { font-size: 16px; color: #555; margin-top: 5px; }
    h4 { font-size: 15px; color: #ffffff; background-color: #34495e; padding: 8px 12px; margin-top: 30px; margin-bottom: 0; border-radius: 5px 5px 0 0; }
    
    /* Responsive Table Wrapper */
    .table-responsive {
        width: 100%;
        overflow-x: auto;
        margin-bottom: 20px;
        border: 1px solid #ccc;
    }

    table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 100%; }
    
    /* Summary Table specific styles */
    .summary-table th, .summary-table td {
        font-size: 12px;
        padding: 8px 6px; 
        white-space: nowrap; 
    }
    .summary-table td.wrap-text {
        white-space: normal; 
        min-width: 150px;
    }

    th { background-color: #2C3E50; color: #ffffff; padding: 12px; text-align: left; font-weight: 600; border-bottom: 2px solid #ddd; }
    td { padding: 10px; border-bottom: 1px solid #eee; color: #333; }
    
    tr:nth-child(even) { background-color: #f9f9f9; }
    tr:hover { background-color: #f1f1f1; }

    td.pass { background-color: #d4edda; color: #155724; font-weight: bold; text-align: center; border-radius: 4px; }
    td.warn { background-color: #fff3cd; color: #856404; font-weight: bold; text-align: center; border-radius: 4px; }
    td.fail { background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center; border-radius: 4px; }
</style>
</head>
<body>
<div class='report-container'>
    <h1>Reporte de Salud: Controladores de Dominio</h1>
    <h3>Generado: $reportTime</h3>"

# Main Table wrapped in DIV
$htmltableheader = "<h3>Resumen General del Bosque: $((Get-ADForest).Name)</h3>
<div class='table-responsive'>
<table class='summary-table'>
    <tr>
    <th>Server</th><th>Site</th><th>OS Version</th><th>IP</th><th style='min-width:140px'>Roles FSMO</th>
    <th style='text-align:center'>DNS</th><th style='text-align:center'>Ping</th>
    <th style='text-align:center'>Uptime (H)</th><th style='text-align:center'>Free Space %</th>
    <th style='text-align:center'>Free Space GB</th><th style='text-align:center'>Time Offset</th>
    <th style='text-align:center'>Svc DNS</th><th style='text-align:center'>Svc NTDS</th><th style='text-align:center'>Svc NetLogon</th>
    </tr>"

# Generate independent DCDIAG tables
$htmlDCDiagTable = "<h2 style='color:#2C3E50; margin-top:40px; border-bottom:1px solid #ccc; padding-bottom:5px;'>Detalle de Pruebas DCDIAG</h2>"

foreach ($reportline in $allTestedDomainControllers) {
    $dcName = $reportline.Server.ToUpper()
    $htmlDCDiagTable += "<h4>CONTROLADOR: $dcName</h4>"
    $htmlDCDiagTable += "<table>"
    $htmlDCDiagTable += "<tr><th style='width:70%'>Prueba Ejecutada</th><th style='width:30%; text-align:center'>Estado</th></tr>"

    $properties = $reportline.PSObject.Properties | Sort-Object Name
    foreach ($property in $properties) {
        if ($property.Name -like "DCDIAG:*") {
            $testName = $property.Name -replace "DCDIAG:\s*", ""
            $testResult = $property.Value
            if ([string]::IsNullOrWhiteSpace($testResult)) {
                $displayResult = "Sin Datos / Error"
                $colorClass = "fail"
            }
            else {
                $displayResult = $testResult
                $colorClass = switch ($testResult) {
                    "Passed" { "pass" }
                    "Failed" { "fail" }
                    "Access Fail" { "fail" }
                    default { "warn" }
                }
            }
            $htmlDCDiagTable += "<tr><td>$testName</td><td class='$colorClass'>$displayResult</td></tr>"
        }
    }
    $htmlDCDiagTable += "</table>"
}

$serverhealthhtmltable = $serverhealthhtmltable + $htmltableheader

foreach ($reportline in $allTestedDomainControllers) {
    if (Test-Path variable:fsmoRoleHTML) { Remove-Variable fsmoRoleHTML }
    if (($reportline."Operation Master Roles").Count -gt 0) {
        $fsmoRoleHTML = ($reportline."Operation Master Roles" | ForEach-Object { "$_`r`n" }) -join '<br>'
    } else { $fsmoRoleHTML = 'None<br>' }

    $htmltablerow = "<tr>"
    $htmltablerow += "<td><b>$($reportline.Server)</b></td>"
    $htmltablerow += "<td>$($reportline.Site)</td>"
    $htmltablerow += "<td>$($reportline."OS Version")</td>"
    $htmltablerow += "<td>$($reportline."IPv4 Address")</td>"
    $htmltablerow += "<td class='wrap-text' style='font-size:11px'>$fsmoRoleHTML</td>" # class wrap-text for FSMO
    $htmltablerow += (New-ServerHealthHTMLTableCell "DNS" )
    $htmltablerow += (New-ServerHealthHTMLTableCell "Ping")

    if ($($reportline."Uptime (hours)") -eq "CIM Failure") { $htmltablerow += "<td class='warn'>Error</td>" }
    elseif ($($reportline."Uptime (hours)") -eq "Fail") { $htmltablerow += "<td class='fail'>Fail</td>" }
    else {
        $hours = [int]$($reportline."Uptime (hours)")
        if ($hours -le 24) { $htmltablerow += "<td class='warn'>$hours</td>" }
        else { $htmltablerow += "<td class='pass'>$hours</td>" }
    }

    $osSpace = $reportline."OS Free Space (%)"
    if ($osSpace -eq "CIM Failure") { $htmltablerow += "<td class='warn'>Error</td>" }
    elseif ($osSpace -eq "Fail") { $htmltablerow += "<td class='fail'>$osSpace</td>" }
    elseif ($osSpace -le 5) { $htmltablerow += "<td class='fail'>$osSpace</td>" }
    elseif ($osSpace -le 30) { $htmltablerow += "<td class='warn'>$osSpace</td>" }
    else { $htmltablerow += "<td class='pass'>$osSpace</td>" }

    $osSpaceGB = $reportline."OS Free Space (GB)"
    if ($osSpaceGB -eq "CIM Failure") { $htmltablerow += "<td class='warn'>Error</td>" }
    elseif ($osSpaceGB -eq "Fail") { $htmltablerow += "<td class='fail'>$osSpaceGB</td>" }
    elseif ($osSpaceGB -lt 5) { $htmltablerow += "<td class='fail'>$osSpaceGB</td>" }
    elseif ($osSpaceGB -lt 10) { $htmltablerow += "<td class='warn'>$osSpaceGB</td>" }
    else { $htmltablerow += "<td class='pass'>$osSpaceGB</td>" }

    $time = $reportline."Time offset (seconds)"
    if ($time -ge 1) { $htmltablerow += "<td class='fail'>$time</td>" }
    else { $htmltablerow += "<td class='pass'>$time</td>" }

    $htmltablerow += (New-ServerHealthHTMLTableCell "DNS Service")
    $htmltablerow += (New-ServerHealthHTMLTableCell "NTDS Service")
    $htmltablerow += (New-ServerHealthHTMLTableCell "NetLogon Service")
    [array]$serverhealthhtmltable += $htmltablerow
}

$serverhealthhtmltable += "</table></div>" # Close responsive DIV
$htmltail = "<p style='font-size:11px; color:#777; margin-top:20px; border-top:1px solid #eee; padding-top:10px;'>
* DNS test is performed using Resolve-DnsName (Windows 2012+). Auto-generated report.</p>
</div>
</body></html>"

$htmlreport = $htmlhead + $serverhealthhtmltable + $htmlDCDiagTable + $htmltail

if ($ReportFile) {
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
