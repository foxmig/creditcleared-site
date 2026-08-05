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
        'job_storage_path', 'log_path', 'template_path',
        'shadow_bureau_reminder_1hr_delay_hours', 'shadow_bureau_reminder_24hr_delay_hours',
        'shadow_bureau_manual_flag_hours', 'shadow_bureau_reminder_template_path'
    )
    $missing = $required | Where-Object { $config.PSObject.Properties.Name -notcontains $_ }
    if ($missing) {
        throw "config.json is missing required field(s): $($missing -join ', ')"
    }

    $baseDir = Split-Path -Parent $Path
    foreach ($field in @('job_storage_path', 'log_path', 'template_path', 'shadow_bureau_reminder_template_path')) {
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

function Expand-Template {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][hashtable]$Tokens
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        throw "Template not found at '$TemplatePath'."
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
        [Parameter(Mandatory)][ValidateSet('queued', 'processing', 'completed', 'archive', 'failed', 'test', 'funding', 'hold')][string]$Stage,
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
        [Parameter(Mandatory)][ValidateSet('queued', 'processing', 'completed', 'archive', 'failed', 'test', 'funding', 'hold')][string]$FromStage,
        [Parameter(Mandatory)][ValidateSet('queued', 'processing', 'completed', 'archive', 'failed', 'test', 'funding', 'hold')][string]$ToStage
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

# The 5 shadow bureaus clients are asked to freeze before analysis proceeds.
# Keep this list in one place -- main.ps1 (upload gate + confirmation
# handler) and shadow-bureau-followup.ps1 both need the exact same set.
$Script:ShadowBureaus = [ordered]@{
    lexisnexis  = 'LexisNexis'
    sagestream  = 'SageStream'
    innovis     = 'Innovis'
    ars         = 'ARS'
    chexsystems = 'ChexSystems'
}

function Get-ShadowBureauChecklist {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Submission
    )

    # Formspree (like any standard HTML form) only includes a checkbox field
    # in the payload when it was checked -- an unchecked box is simply
    # absent, not present-with-value-false. Presence of a truthy value is
    # "checked".
    $checklist = [ordered]@{}
    foreach ($key in $Script:ShadowBureaus.Keys) {
        $fieldName = "freeze_$key"
        $hasField = $Submission.PSObject.Properties.Name -contains $fieldName
        $checklist[$key] = [bool]($hasField -and $Submission.$fieldName)
    }
    return $checklist
}

function Test-FormspreeSignature {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][byte[]]$Body,
        [Parameter(Mandatory)][string]$Secret,
        [string]$SignatureHeaderValue
    )

    # Confirmed against a real webhook delivery. Formspree sends a
    # "Formspree-Signature" header (note: no "X-" prefix) shaped like:
    #   t=<unix-timestamp>,v1=<hex-hmac>
    # The signed message is "{timestamp}.{raw_body}" (raw bytes, not
    # re-serialized), HMAC-SHA256'd with the webhook secret and hex-encoded.
    # The timestamp is included so replay of an old captured request can be
    # rejected -- we enforce a 5 minute tolerance on it below.
    if ([string]::IsNullOrEmpty($SignatureHeaderValue)) { return $false }

    $parts = @{}
    foreach ($segment in $SignatureHeaderValue.Split(',')) {
        $kv = $segment.Split('=', 2)
        if ($kv.Count -eq 2) { $parts[$kv[0].Trim()] = $kv[1].Trim() }
    }

    $timestamp = $parts['t']
    $signature = $parts['v1']
    if ([string]::IsNullOrEmpty($timestamp) -or [string]::IsNullOrEmpty($signature)) { return $false }

    $timestampSeconds = 0L
    if (-not [long]::TryParse($timestamp, [ref]$timestampSeconds)) { return $false }
    $requestTime = [DateTimeOffset]::FromUnixTimeSeconds($timestampSeconds)
    $skew = [Math]::Abs(([DateTimeOffset]::UtcNow - $requestTime).TotalSeconds)
    if ($skew -gt 300) { return $false }

    $bodyText = [System.Text.Encoding]::UTF8.GetString($Body)
    $signedPayloadBytes = [System.Text.Encoding]::UTF8.GetBytes("$timestamp.$bodyText")

    # Use [Type]::new(...) rather than New-Object here -- New-Object's
    # -ArgumentList enumerates an array argument into one constructor
    # parameter per element (e.g. a 64-byte key becomes 64 arguments), which
    # never matches any HMACSHA256 overload. ::new() passes the byte[] key
    # as a single argument, as intended.
    $hmac = [System.Security.Cryptography.HMACSHA256]::new([System.Text.Encoding]::UTF8.GetBytes($Secret))
    try {
        $hash = $hmac.ComputeHash($signedPayloadBytes)
    } finally {
        $hmac.Dispose()
    }
    $computed = ($hash | ForEach-Object { $_.ToString('x2') }) -join ''

    if ($computed.Length -ne $signature.Length) { return $false }
    $diff = 0
    for ($i = 0; $i -lt $computed.Length; $i++) {
        $diff = $diff -bor ([byte][char]$computed[$i] -bxor [byte][char]$signature[$i])
    }
    return ($diff -eq 0)
}

function Invoke-ClaudeMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$SystemPrompt,
        [Parameter(Mandatory)][array]$Messages,
        [int]$MaxTokens = 4096,
        [double]$Temperature = 0.3
    )

    $bodyObj = @{
        model       = $Config.claude_model
        max_tokens  = $MaxTokens
        temperature = $Temperature
        system      = $SystemPrompt
        messages    = $Messages
    }
    $bodyJson = $bodyObj | ConvertTo-Json -Depth 12

    $headers = @{
        'x-api-key'         = $Config.anthropic_api_key
        'anthropic-version' = '2023-06-01'
    }

    # Section 8: "Claude API call fails (timeout or error) -> Retry once after
    # 60 seconds. If still failing, move job to failed\ and email alert."
    $maxAttempts = 2
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            return Invoke-RestMethod -Uri 'https://api.anthropic.com/v1/messages' -Method Post `
                -Headers $headers -ContentType 'application/json' -Body $bodyJson -TimeoutSec 300
        } catch {
            if ($attempt -ge $maxAttempts) { throw }
            Write-Warning "Claude API call failed (attempt $attempt of $maxAttempts): $_. Retrying in 60 seconds..."
            Start-Sleep -Seconds 60
        }
    }
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
            #
            # Windows PowerShell 5.1 (.NET Framework) defaults ServicePointManager
            # to older SSL/TLS protocols and will silently fail the handshake
            # against a TLS 1.2-only host. Force TLS 1.2 before connecting.
            if (-not ([System.Net.ServicePointManager]::SecurityProtocol -band [System.Net.SecurityProtocolType]::Tls12)) {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
            }
            $webClient = New-Object System.Net.WebClient
            try {
                return $webClient.DownloadData($CreditReportField)
            } catch {
                # Don't let a transient download failure (expired signed URL,
                # TLS/network hiccup, etc.) crash the whole webhook request --
                # log the real reason and let the caller treat it as "no
                # usable attachment", which alerts the operator instead.
                Write-Warning "Failed to download credit report from '$CreditReportField': $_"
                return $null
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
