# send_mail.ps1 — Version simplifiée avec ton mail perso en attendant
# ou avec un service tiers gratuit comme SendGrid

$smtpServer = "smtp.office365.com"
$smtpPort   = 587
$from       = "autoreport@ningen-group.com"
$to         = "pw39f@ningen-group.com"
$password   = ConvertTo-SecureString "Cctsnlrvqwnrsrsh" -AsPlainText -Force
$cred       = New-Object System.Management.Automation.PSCredential($from, $password)

Send-MailMessage `
    -From $from `
    -To $to `
    -Subject "Sanity Check - Resultat: $env:BUILD_RESULT" `
    -Body "Build: $env:BUILD_NAME | Resultat: $env:BUILD_RESULT" `
    -BodyAsHtml `
    -Attachments "C:\Autoreports\SanityCheck\reports\sanity_check_report.html" `
    -SmtpServer $smtpServer `
    -Port $smtpPort `
    -UseSsl `
    -Credential $cred
