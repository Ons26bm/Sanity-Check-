$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)

$Mail.To = "pw39f@ningen-group.com"
$Mail.Subject = "Sanity Check - Résultat: $env:BUILD_RESULT"
$Mail.HTMLBody = @"
<p>Le pipeline est terminé.</p>
<p><b>Build:</b> $env:BUILD_NAME</p>
<p><b>Résultat:</b> $env:BUILD_RESULT</p>
"@

# Attacher le rapport
$attachmentPath = "C:\Autoreports\SanityCheck\reports\sanity_check_report.html"
if (Test-Path $attachmentPath) {
    $Mail.Attachments.Add($attachmentPath)
}

$Mail.Send()
