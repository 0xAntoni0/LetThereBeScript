# Output path for the CSV file
$CSVpath = "C:\Scripts\LastLogonReport.csv"

# Get all domain controllers
$DCs = Get-ADDomainController -Filter *

# Get all enabled users in Active Directory with extra properties
$USERS = Get-ADUser -Filter {Enabled -eq $true} -Properties SamAccountName, DisplayName, EmailAddress, DistinguishedName

# Create a list to store the results
$LASTLOGONLIST = @()

foreach ($user in $USERS) {
    $ultimoLogon = 0

    foreach ($dc in $DCs) {
        # Query each domain controller for the user's LastLogon value
        $logon = (Get-ADUser $user.SamAccountName -Server $dc.HostName -Properties LastLogon).LastLogon
        if ($logon -gt $ultimoLogon) {
            $ultimoLogon = $logon
        }
    }

    # Convert FileTime to readable DateTime format
    if ($ultimoLogon -ne 0) {
        $fechaLogon = [DateTime]::FromFileTime($ultimoLogon)
    } else {
        $fechaLogon = $null
    }

    # Add the result to the list
    $LASTLOGONLIST += [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        SamAccountName    = $user.SamAccountName
        EmailAddress      = $user.EmailAddress
        DistinguishedName = $user.DistinguishedName
        LastLogonDate     = $fechaLogon
    }
}

# Sort alphabetically by DisplayName and export to CSV
$LASTLOGONLIST | Sort-Object DisplayName | Export-Csv -Path $CSVpath -NoTypeInformation -Encoding UTF8