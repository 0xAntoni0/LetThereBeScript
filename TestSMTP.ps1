$Tab = [char]9
$Salto = "`r`n"
$EmailTo = "ehernandez@atlantistecnologia.com"
$EmailFrom = "authreservas@sierraygonzalez.com"
$Subject = "Prueba Script Correo"
$Body= "TEST VA"
$SMTPServer = "smtp.office365.com"
$Username= "authreservas@sierraygonzalez.com"
$Password= ""
$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
$smtpclient.EnableSsl = $true
$Smtpclient.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)

$SMTPClient.Send($SMTPMessage)
