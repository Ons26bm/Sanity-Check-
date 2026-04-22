def content = '''$Outlook = New-Object -ComObject Outlook.Application
$Mail    = $Outlook.CreateItem(0)

$Mail.To       = "pw39f@ningen-group.com"
$Mail.Subject  = "Sanity Check - Resultat: $env:BUILD_RESULT"
$Mail.HTMLBody = "<p>Le pipeline est termine.</p><p><b>Build:</b> $env:BUILD_NAME</p><p><b>Resultat:</b> $env:BUILD_RESULT</p>"

$attachmentPath = "C:\\Autoreports\\SanityCheck\\reports\\sanity_check_report.html"
if (Test-Path $attachmentPath) {
    $Mail.Attachments.Add($attachmentPath) | Out-Null
    Write-Host "Rapport attache"
} else {
    Write-Host "Rapport introuvable : $attachmentPath"
}

$Mail.Send()
Write-Host "Mail envoye"'''

new File("C:\\Autoreports\\SanityCheck\\reports\\send_mail.ps1").text = content
println "Fichier mis a jour ✅"
println new File("C:\\Autoreports\\SanityCheck\\reports\\send_mail.ps1").text
