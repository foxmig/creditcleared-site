<#
.SYNOPSIS
    Credit Cleared webhook listener and job orchestrator.
.DESCRIPTION
    Listens for Formspree webhook POSTs, validates the request, extracts client
    data and the credit report PDF, creates a job folder, sends an immediate
    confirmation email, and hands the job to analyze.ps1 as a background
    process. See docs/CreditCleared-ClaudeCode-Spec.docx Section 4.
.PARAMETER TestMode
    Routes jobs to jobs\test\ instead of jobs\queued\, and tags job.json with
    test_mode = true so downstream modules (analyze.ps1, deliver.ps1) use the
    shortened delay and deliver to the operator instead of the client.
#>
[CmdletBinding()]
param(
    [switch]$TestMode
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = $PSScriptRoot
. (Join-Path $ScriptRoot 'common.ps1')

function Write-CrashLog {
    # Last-resort logger for failures that happen before config is loaded, or
    # that would otherwise have nowhere to go and fail completely silently
    # (e.g. an exception escaping the main loop). Never throws itself.
    param([string]$Message, $Config)

    $logDir = if ($Config -and $Config.log_path) { $Config.log_path } else { Join-Path $ScriptRoot 'logs' }
    try {
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -LiteralPath (Join-Path $logDir 'crash.log') -Value $line
    } catch {
        # If even the crash log can't be written, at least surface it on the console.
        Write-Error "Write-CrashLog failed: $_"
    }
}

try {
    $Config = Get-CreditClearedConfig -Path (Join-Path $ScriptRoot 'config.json')
} catch {
    Write-CrashLog -Message "FATAL -- failed to load config.json: $($_.Exception.GetType().FullName): $_`n$($_.ScriptStackTrace)"
    throw
}
$Port = [int]$Config.webhook_port
$Stage = if ($TestMode) { 'test' } else { 'queued' }

function Send-HttpResponse {
    param($Context, [int]$StatusCode = 200, [string]$Body = 'OK')
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Context.Response.StatusCode = $StatusCode
    $Context.Response.ContentType = 'text/plain'
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Context.Response.OutputStream.Close()
}

function Start-AnalysisJob {
    param([string]$JobId, [switch]$TestMode)

    $analyzeScript = Join-Path $ScriptRoot 'analyze.ps1'
    if (-not (Test-Path -LiteralPath $analyzeScript)) {
        Write-Warning "analyze.ps1 not found yet -- job '$JobId' was created but will not be analyzed until analyze.ps1 is built."
        return
    }

    $arguments = @('-NoProfile', '-File', $analyzeScript, '-JobId', $JobId)
    if ($TestMode) { $arguments += '-TestMode' }
    Start-Process -FilePath 'pwsh' -ArgumentList $arguments -WindowStyle Hidden
}

function Invoke-WebhookRequest {
    param($Context, $Config, [switch]$TestMode, [string]$Stage)

    $request = $Context.Request
    $remoteIp = $request.RemoteEndPoint.Address.ToString()

    if ($request.HttpMethod -ne 'POST') {
        Send-HttpResponse -Context $Context -StatusCode 405 -Body 'Method Not Allowed'
        return
    }

    # Read the raw body bytes BEFORE any parsing -- signature verification
    # must run against the exact bytes Formspree sent, not a re-serialized copy.
    $memoryStream = New-Object System.IO.MemoryStream
    $request.InputStream.CopyTo($memoryStream)
    $bodyBytes = $memoryStream.ToArray()
    $memoryStream.Dispose()

    $signatureHeader = $request.Headers[$Config.formspree_signature_header]
    $signatureOk = Test-FormspreeSignature -Body $bodyBytes -Secret $Config.formspree_secret -SignatureHeaderValue $signatureHeader

    if (-not $signatureOk) {
        # If this rejects your very first real "Send test" from Formspree, the
        # assumed header name/scheme in Test-FormspreeSignature (common.ps1) is
        # probably wrong for your account -- the headers dumped below show what
        # Formspree actually sent, so you can fix formspree_signature_header in
        # config.json or adjust the verification scheme accordingly.
        $headerDump = ($request.Headers.AllKeys | ForEach-Object { "$_=$($request.Headers[$_])" }) -join ' | '
        Write-Warning "Rejected webhook request from $remoteIp -- signature verification failed. Headers: $headerDump"
        Send-HttpResponse -Context $Context -StatusCode 401 -Body 'Invalid signature'
        return
    }

    $bodyText = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
    try {
        $payload = $bodyText | ConvertFrom-Json
    } catch {
        Write-Warning "Rejected webhook request from $remoteIp -- body is not valid JSON: $_"
        Send-HttpResponse -Context $Context -StatusCode 400 -Body 'Invalid JSON'
        return
    }

    # Formspree wraps the actual form fields under a "submission" object --
    # {"form": "...", "submission": {first_name, email, credit_report, ...}}.
    # Fall back to the top level too, in case a future Formspree payload (or
    # the "Send test" button) ever sends the fields unwrapped.
    $submission = if ($payload.PSObject.Properties.Name -contains 'submission') { $payload.submission } else { $payload }

    $firstName = $submission.first_name
    $lastName = $submission.last_name
    $email = if ($submission._replyto) { $submission._replyto } else { $submission.email }
    $jobId = New-JobId -FirstName $firstName -LastName $lastName

    if ([string]::IsNullOrWhiteSpace($email)) {
        Write-Warning "Webhook payload for job '$jobId' has no client email -- cannot proceed."
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Webhook from $remoteIp rejected -- no usable client email (_replyto/email)."
        Send-OperatorAlert -Config $Config -Subject "$firstName $lastName -- webhook missing client email" `
            -Message "A Formspree submission came in without a usable email address (_replyto/email). Check manually.`n`nRaw payload:`n$bodyText"
        Send-HttpResponse -Context $Context -StatusCode 200 -Body 'Received (alert sent -- missing email)'
        return
    }

    $reportBytes = Get-CreditReportBytes -CreditReportField $submission.credit_report
    if (-not $reportBytes -or $reportBytes.Length -eq 0) {
        # Section 8: "Webhook received but no PDF attachment" -- alert operator, do not proceed.
        Write-Warning "Job '$jobId' has no credit report attachment -- alerting operator, not creating job."
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Webhook from $remoteIp rejected -- no usable credit_report attachment for $email."
        Send-OperatorAlert -Config $Config -Subject "$firstName $lastName uploaded without PDF -- check manually" `
            -Message "Formspree submission from $email came in without a usable credit_report attachment.`n`nJob would have been: $jobId`n`nRaw payload:`n$bodyText"
        Send-HttpResponse -Context $Context -StatusCode 200 -Body 'Received (alert sent -- missing PDF)'
        return
    }

    $jobFolder = Get-JobFolder -Config $Config -Stage $Stage -JobId $jobId -Create
    $reportPath = Join-Path $jobFolder 'report.pdf'
    [System.IO.File]::WriteAllBytes($reportPath, $reportBytes)

    $job = [ordered]@{
        job_id        = $jobId
        status        = $Stage
        test_mode     = [bool]$TestMode
        created_at    = (Get-Date).ToString('o')
        client        = [ordered]@{
            first_name = $firstName
            last_name  = $lastName
            email      = $email
            phone      = $submission.phone
            goal       = $submission.goal
            deadline   = $submission.deadline
            notes      = $submission.notes
        }
        files         = @{ report_pdf = 'report.pdf' }
        confirmation  = @{ sent_at = $null }
        analysis      = @{ consumer_completed_at = $null; consumer_response_file = $null }
        delivery      = @{ scheduled_at = $null; delivered_at = $null }
        funding_addon = @{ purchased = $false; completed_at = $null }
    }
    $jobJsonPath = Join-Path $jobFolder 'job.json'
    ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8

    Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Webhook received from $remoteIp. Client: $firstName $lastName <$email>. Report saved ($($reportBytes.Length) bytes)."

    # Confirmation email: in test mode, send to the operator (not the client),
    # per spec Section 10 test mode requirements.
    $confirmationTemplate = Join-Path $Config.template_path 'confirmation-email.html'
    $confirmationTo = if ($TestMode) { $Config.your_email } else { $email }
    $subject = 'Got it — your analysis has started'
    if ($TestMode) { $subject = "[TEST] $subject" }

    try {
        $bodyHtml = Expand-Template -TemplatePath $confirmationTemplate -Tokens @{
            CLIENT_FIRST_NAME = $firstName
            YOUR_NAME         = $Config.your_name
            YOUR_PHONE        = $Config.your_phone
            BUSINESS_DBA      = $Config.business_dba
            BUSINESS_NAME     = $Config.business_name
        }
        Send-CreditClearedEmail -Config $Config -To $confirmationTo -Subject $subject -BodyHtml $bodyHtml -ReplyTo $Config.gmail_address
        $job.confirmation.sent_at = (Get-Date).ToString('o')
        ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Confirmation email sent to $confirmationTo."
    } catch {
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "FAILED to send confirmation email: $_"
        Send-OperatorAlert -Config $Config -Subject "$firstName $lastName -- confirmation email failed to send" `
            -Message "Job $jobId was created but the confirmation email failed to send:`n$_"
    }

    Start-AnalysisJob -JobId $jobId -TestMode:$TestMode
    Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Handed off to analyze.ps1."

    Send-HttpResponse -Context $Context -StatusCode 200 -Body "Received. Job: $jobId"
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$Port/webhook/")

try {
    $listener.Start()
} catch {
    Write-CrashLog -Config $Config -Message "FATAL -- failed to start listener on port $Port`: $($_.Exception.GetType().FullName): $_`n$($_.ScriptStackTrace)"
    Write-Error "Failed to start listener on port $Port. On Windows this usually means the URL isn't reserved for your account or the port is in use. Try running as Administrator, or reserve the URL once with: netsh http add urlacl url=http://+:$Port/webhook/ user=Everyone`n$_"
    throw
}

Write-Host "Credit Cleared webhook listener ready on port $Port$(if ($TestMode) { ' [TEST MODE]' })."

try {
    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
        } catch {
            Write-CrashLog -Config $Config -Message "FATAL -- listener.GetContext() failed, listener loop is exiting: $($_.Exception.GetType().FullName): $_`n$($_.ScriptStackTrace)"
            throw
        }
        try {
            Invoke-WebhookRequest -Context $context -Config $Config -TestMode:$TestMode -Stage $Stage
        } catch {
            $errDetail = "Unhandled error processing webhook request: $($_.Exception.GetType().FullName): $_`n$($_.ScriptStackTrace)"
            Write-Warning $errDetail
            Write-CrashLog -Config $Config -Message $errDetail
            try {
                Send-HttpResponse -Context $context -StatusCode 500 -Body 'Internal error'
            } catch {}
        }
    }
} catch {
    # Anything that escapes the per-request handling above (e.g. GetContext()
    # failing) used to fall through this try with only a `finally` and no
    # `catch` -- the process would exit with the listener loop simply ending,
    # nothing written anywhere, nothing printed. Log it before it disappears.
    Write-CrashLog -Config $Config -Message "FATAL -- main.ps1 is exiting due to an unhandled error: $($_.Exception.GetType().FullName): $_`n$($_.ScriptStackTrace)"
    Write-Error "main.ps1 is exiting due to an unhandled error: $_"
    throw
} finally {
    $listener.Stop()
    $listener.Close()
}
