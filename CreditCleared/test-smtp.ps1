# Standalone SMTP auth test - isolates Gmail credential issue
# Reads credentials from config.json (gitignored) rather than hardcoding them.
$ScriptRoot = $PSScriptRoot
. (Join-Path $ScriptRoot 'common.ps1')
$Config = Get-CreditClearedConfig -Path (Join-Path $ScriptRoot 'config.json')
$gmailAddress = $Config.gmail_address
$gmailAppPassword = $Config.gmail_app_password

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
