# Standalone SMTP auth test - isolates Gmail credential issue
$gmailAddress = "cliff@mycreditcleared.com"
$gmailAppPassword = "xlecxiglejalakfc"

try {
    $smtp = New-Object System.Net.Mail.SmtpClient('smtp.gmail.com', 587)
    $smtp.EnableSsl = $true
    $smtp.UseDefaultCredentials = $false
    $smtp.Credentials = New-Object System.Net.NetworkCredential($gmailAddress, $gmailAppPassword)

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = $gmailAddress
    $mail.To.Add($gmailAddress)
    $mail.Subject = "SMTP Test"
    $mail.Body = "If you receive this, SMTP auth is working."

    $smtp.Send($mail)
    Write-Host "SUCCESS: Email sent." -ForegroundColor Green
}
catch {
    Write-Host "FAILED:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Inner exception:" -ForegroundColor Yellow
        Write-Host $_.Exception.InnerException.Message -ForegroundColor Yellow
    }
}
