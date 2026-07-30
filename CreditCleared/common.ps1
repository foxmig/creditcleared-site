# Shared helpers for the Credit Cleared automation modules (main.ps1, analyze.ps1,
# deliver.ps1). Dot-source this file: . (Join-Path $PSScriptRoot 'common.ps1')

function Get-CreditClearedConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config file not found at '$Path'. Copy config.example.json to config.json and fill in your values."
    }

    $config = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json

    $required = @(
        'anthropic_api_key', 'claude_model', 'gmail_address', 'gmail_app_password',
        'your_name', 'your_phone', 'your_email', 'formspree_secret', 'webhook_port',
        'delivery_delay_hours', 'delivery_time_hour', 'delivery_time_minute',
        'job_storage_path', 'log_path', 'template_path'
    )
    $missing = $required | Where-Object { $config.PSObject.Properties.Name -notcontains $_ }
    if ($missing) {
        throw "config.json is missing required field(s): $($missing -join ', ')"
    }

    $baseDir = Split-Path -Parent $Path
    foreach ($field in @('job_storage_path', 'log_path', 'template_path')) {
        $raw = $config.$field
        if (-not [System.IO.Path]::IsPathRooted($raw)) {
            $config.$field = (Join-Path $baseDir $raw)
        }
    }

    return $config
}

function Write-JobLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][string]$LogPath
    )

    if (-not (Test-Path -LiteralPath $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }

    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -LiteralPath (Join-Path $LogPath "$JobId.log") -Value $line
    Write-Verbose $line
}

function New-JobId {
    [CmdletBinding()]
    param(
        [string]$FirstName,
        [string]$LastName
    )

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $slug = ("$FirstName $LastName").Trim().ToLowerInvariant() -replace '[^a-z0-9]+', '-'
    $slug = $slug.Trim('-')
    if ([string]::IsNullOrWhiteSpace($slug)) { $slug = 'client' }

    return "$timestamp-$slug"
}

function Expand-EmailTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][hashtable]$Tokens
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        throw "Email template not found at '$TemplatePath'."
    }

    $body = Get-Content -LiteralPath $TemplatePath -Raw
    foreach ($key in $Tokens.Keys) {
        $body = $body.Replace("{{$key}}", [string]$Tokens[$key])
    }
    return $body
}

function Send-CreditClearedEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$To,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyHtml,
        [string[]]$Attachments = @(),
        [string]$ReplyTo
    )

    $smtp = New-Object System.Net.Mail.SmtpClient('smtp.gmail.com', 587)
    $smtp.EnableSsl = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential($Config.gmail_address, $Config.gmail_app_password)

    $mail = New-Object System.Net.Mail.MailMessage
    try {
        $mail.From = New-Object System.Net.Mail.MailAddress($Config.gmail_address, "$($Config.your_name) at $($Config.business_dba)")
        $mail.To.Add($To)
        if ($ReplyTo) { $mail.ReplyToList.Add($ReplyTo) }
        $mail.Subject = $Subject
        $mail.Body = $BodyHtml
        $mail.IsBodyHtml = $true

        foreach ($path in $Attachments) {
            if (Test-Path -LiteralPath $path) {
                $mail.Attachments.Add((New-Object System.Net.Mail.Attachment($path)))
            } else {
                Write-Warning "Attachment not found, skipping: $path"
            }
        }

        $smtp.Send($mail)
    } finally {
        foreach ($att in $mail.Attachments) { $att.Dispose() }
        $mail.Dispose()
        $smtp.Dispose()
    }
}

function Send-OperatorAlert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Message,
        [string[]]$Attachments = @()
    )

    $encoded = [System.Net.WebUtility]::HtmlEncode($Message) -replace "`n", '<br>'
    $bodyHtml = "<p>$encoded</p>"
    try {
        Send-CreditClearedEmail -Config $Config -To $Config.your_email -Subject "[Credit Cleared Alert] $Subject" -BodyHtml $bodyHtml -Attachments $Attachments
    } catch {
        Write-Warning "Failed to send operator alert email: $_"
    }
}

function Get-JobFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][ValidateSet('queued', 'processing', 'completed', 'archive', 'failed', 'test', 'funding')][string]$Stage,
        [Parameter(Mandatory)][string]$JobId,
        [switch]$Create
    )

    $path = Join-Path (Join-Path $Config.job_storage_path $Stage) $JobId
    if ($Create -and -not (Test-Path -LiteralPath $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
    return $path
}

function Move-JobFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][ValidateSet('queued', 'processing', 'completed', 'archive', 'failed', 'test', 'funding')][string]$FromStage,
        [Parameter(Mandatory)][ValidateSet('queued', 'processing', 'completed', 'archive', 'failed', 'test', 'funding')][string]$ToStage
    )

    $source = Get-JobFolder -Config $Config -Stage $FromStage -JobId $JobId
    $destRoot = Join-Path $Config.job_storage_path $ToStage
    if (-not (Test-Path -LiteralPath $destRoot)) {
        New-Item -ItemType Directory -Path $destRoot -Force | Out-Null
    }
    $dest = Join-Path $destRoot $JobId
    Move-Item -LiteralPath $source -Destination $dest -Force
    return $dest
}

function Test-FormspreeSignature {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][byte[]]$Body,
        [Parameter(Mandatory)][string]$Secret,
        [string]$SignatureHeaderValue
    )

    # ASSUMPTION -- unverified against a real Formspree payload: this checks for
    # an HMAC-SHA256 signature ("sha256=<hex>") the way Stripe/GitHub-style
    # webhooks do, using formspree_secret as the HMAC key. Formspree may
    # instead send the raw shared secret as a plain header value. CONFIRM the
    # actual scheme using Formspree's "Send test" button (spec Section 9B) --
    # inspect the incoming request headers logged below and adjust this
    # function to match before relying on it for real traffic.
    if ([string]::IsNullOrEmpty($SignatureHeaderValue)) { return $false }

    $hmac = New-Object System.Security.Cryptography.HMACSHA256([System.Text.Encoding]::UTF8.GetBytes($Secret))
    try {
        $hash = $hmac.ComputeHash($Body)
    } finally {
        $hmac.Dispose()
    }
    $computed = 'sha256=' + (($hash | ForEach-Object { $_.ToString('x2') }) -join '')

    # Also accept a bare shared-secret header as a fallback, in case Formspree
    # sends the plain secret rather than an HMAC signature.
    if ($SignatureHeaderValue.Trim() -eq $Secret) { return $true }

    if ($computed.Length -ne $SignatureHeaderValue.Trim().Length) { return $false }
    $diff = 0
    for ($i = 0; $i -lt $computed.Length; $i++) {
        $diff = $diff -bor ([byte][char]$computed[$i] -bxor [byte][char]$SignatureHeaderValue.Trim()[$i])
    }
    return ($diff -eq 0)
}

function Get-CreditReportBytes {
    [CmdletBinding()]
    param(
        $CreditReportField
    )

    # ASSUMPTION -- unverified: Formspree's exact webhook JSON shape for an
    # uploaded file is not confirmed. This handles the plausible shapes
    # (a direct URL, an object with a "url" property, a base64 string, or an
    # array wrapping any of those). CONFIRM against a real "Send test"
    # payload (spec Section 9B) and simplify once the actual shape is known.
    if ($null -eq $CreditReportField) { return $null }

    if ($CreditReportField -is [System.Array]) {
        if ($CreditReportField.Count -eq 0) { return $null }
        return Get-CreditReportBytes -CreditReportField $CreditReportField[0]
    }

    if ($CreditReportField -is [string]) {
        if ($CreditReportField -match '^https?://') {
            # WebClient.DownloadData reliably returns raw bytes for binary
            # content (a PDF) regardless of PowerShell version, unlike
            # Invoke-WebRequest's .Content, whose type varies by version.
            $webClient = New-Object System.Net.WebClient
            try {
                return $webClient.DownloadData($CreditReportField)
            } finally {
                $webClient.Dispose()
            }
        }
        try {
            return [Convert]::FromBase64String($CreditReportField)
        } catch {
            return $null
        }
    }

    if ($CreditReportField.PSObject.Properties.Name -contains 'url') {
        return Get-CreditReportBytes -CreditReportField $CreditReportField.url
    }

    return $null
}
