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

function Find-RecentDuplicateJob {
    # Guards against Formspree retries (or any other double-POST) creating a
    # second job for the same submission. Not a perfect dedupe key -- Formspree's
    # webhook payload doesn't include a stable submission ID in what we've seen --
    # so this matches on client email + exact report byte length within a short
    # window. Good enough to catch retries (seconds apart) without false-positiving
    # on a client who legitimately re-uploads later.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Email,
        [Parameter(Mandatory)][int]$ReportByteLength,
        [int]$WindowSeconds = 120
    )

    $normalizedEmail = $Email.Trim().ToLowerInvariant()
    $cutoff = (Get-Date).AddSeconds(-$WindowSeconds)

    foreach ($stage in @('queued', 'hold', 'test')) {
        $stageRoot = Join-Path $Config.job_storage_path $stage
        if (-not (Test-Path -LiteralPath $stageRoot)) { continue }

        $matches = Get-ChildItem -LiteralPath $stageRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $jsonPath = Join-Path $_.FullName 'job.json'
            $pdfPath = Join-Path $_.FullName 'report.pdf'
            if (-not (Test-Path -LiteralPath $jsonPath) -or -not (Test-Path -LiteralPath $pdfPath)) { return }

            try {
                $candidate = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
            } catch { return }

            $createdAt = [datetime]$candidate.created_at
            if ($createdAt -lt $cutoff) { return }

            $candidateEmail = [string]$candidate.client.email
            if (-not $candidateEmail -or ($candidateEmail.Trim().ToLowerInvariant() -ne $normalizedEmail)) { return }

            $candidatePdfLength = (Get-Item -LiteralPath $pdfPath).Length
            if ($candidatePdfLength -ne $ReportByteLength) { return }

            [pscustomobject]@{ JobId = $candidate.job_id; CreatedAt = $createdAt }
        }

        if ($matches) {
            return ($matches | Sort-Object CreatedAt -Descending | Select-Object -First 1)
        }
    }

    return $null
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

# Peek at the form ID so we know which secret to verify against.
# This peek doesn't grant trust -- Test-FormspreeSignature below still
# does the real verification against the matched secret.
$bodyTextPeek = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
$formId = $null
try {
    $peekJson = $bodyTextPeek | ConvertFrom-Json
    $formId = $peekJson.form
} catch {
    $formId = $null
}

if (-not $formId -or -not $Config.formspree_secrets.PSObject.Properties[$formId]) {
    Write-Warning "Rejected webhook request from $remoteIp -- unrecognized form ID '$formId'"
    Send-HttpResponse -Context $Context -StatusCode 401 -Body 'Invalid signature or form.'
    return
}

$formSecret = $Config.formspree_secrets.$formId
$signatureOk = Test-FormspreeSignature -Body $bodyBytes -Secret $formSecret -SignatureHeaderValue $signatureHeader

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

    # This endpoint receives two different Formspree forms: the credit report
    # upload (upload-page.html) and the shadow bureau freeze confirmation
    # (confirm-freeze.html). The confirmation form always posts "client_email"
    # and never a credit_report file; the upload form is the opposite. Route
    # accordingly before assuming this is an upload.
    $hasClientEmail = ($submission.PSObject.Properties.Name -contains 'client_email') -and (-not [string]::IsNullOrWhiteSpace([string]$submission.client_email))
    $hasCreditReportField = ($submission.PSObject.Properties.Name -contains 'credit_report') -and $submission.credit_report
    if ($hasClientEmail -and -not $hasCreditReportField) {
        Invoke-ShadowBureauConfirmation -Context $Context -Config $Config -TestMode:$TestMode -Submission $submission -RemoteIp $remoteIp -BodyText $bodyText
        return
    }

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
    $duplicateJob = Find-RecentDuplicateJob -Config $Config -Email $email -ReportByteLength $reportBytes.Length
    if ($duplicateJob) {
        Write-Warning "Webhook from $remoteIp for '$email' looks like a duplicate of job '$($duplicateJob.JobId)' (matching email + report size within window) -- likely a Formspree retry. Skipping."
        Write-JobLog -JobId $duplicateJob.JobId -LogPath $Config.log_path -Message "Duplicate webhook POST received from $remoteIp and skipped (same email + report size)."
        Send-HttpResponse -Context $Context -StatusCode 200 -Body "Duplicate ignored. Existing job: $($duplicateJob.JobId)"
        return
    }

    # Section: Shadow Bureau Freeze Confirmation. Formspree omits a checkbox
    # field entirely from the payload when it's unchecked -- see
    # Get-ShadowBureauChecklist (common.ps1) -- so "checked" is presence of a
    # truthy freeze_<bureau> field.
    $shadowBureauChecklist = Get-ShadowBureauChecklist -Submission $submission
    $shadowBureauIncomplete = $shadowBureauChecklist.Values -contains $false
    foreach ($key in $shadowBureauChecklist.Keys) {
        $status = if ($shadowBureauChecklist[$key]) { 'checked' } else { 'unchecked' }
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Shadow bureau freeze checkbox -- $($Script:ShadowBureaus[$key]): $status."
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
        shadow_bureau = [ordered]@{
            incomplete   = $shadowBureauIncomplete
            checklist    = $shadowBureauChecklist
            confirmed_at = $null
        }
        analysis      = @{ consumer_completed_at = $null; consumer_response_file = $null }
        delivery      = @{ scheduled_at = $null; delivered_at = $null }
        funding_addon = @{ purchased = $false; completed_at = $null }
    }
    $jobJsonPath = Join-Path $jobFolder 'job.json'
    ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8

    Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Webhook received from $remoteIp. Client: $firstName $lastName <$email>. Report saved ($($reportBytes.Length) bytes)."

    # Ack Formspree immediately -- everything after this point (SMTP send,
    # shadow bureau branching) is slow-ish and must not block the response,
    # or Formspree's retry logic fires and creates duplicate jobs.
    Send-HttpResponse -Context $Context -StatusCode 200 -Body "Received. Job: $jobId"

    # Confirmation email: in test mode, send to the operator (not the client),
    # per spec Section 10 test mode requirements.

# Which template/subject we use depends on whether the shadow bureau
    # freeze gate (computed earlier) is still incomplete -- an incomplete job
    # needs the freeze-instructions variant with a confirm link, not the
    # "roadmap is in progress" copy.
    if ($shadowBureauIncomplete) {
        $confirmationTemplate = Join-Path $Config.template_path 'confirmation-email-hold.html'
        $subject = 'Got it — one thing before we start'
    } else {
        $confirmationTemplate = Join-Path $Config.template_path 'confirmation-email.html'
        $subject = 'Got it — your analysis has started'
    }
    $confirmationTo = if ($TestMode) { $Config.your_email } else { $email }
    if ($TestMode) { $subject = "[TEST] $subject" }

    # Same confirm-link construction shadow-bureau-followup.ps1 uses for the
    # 1hr/24hr reminders -- built here too so the very first email (not just
    # the later reminders) can link straight to confirm-freeze.html.
    $confirmUri = New-Object System.Uri([Uri]$Config.upload_page_url, 'confirm-freeze.html')
    $confirmUrl = "$confirmUri" + '?email=' + [Uri]::EscapeDataString([string]$email)

    try {
        $bodyHtml = Expand-Template -TemplatePath $confirmationTemplate -Tokens @{
            CLIENT_FIRST_NAME = $firstName
            CONFIRM_URL       = $confirmUrl
            YOUR_NAME         = $Config.your_name
            BUSINESS_DBA      = $Config.business_dba
            BUSINESS_NAME     = $Config.business_name
        }
        $educationalSummaryPath = Join-Path $Config.template_path 'Credit Cleared - Educational Summary.docx'
        Send-CreditClearedEmail -Config $Config -To $confirmationTo -Subject $subject -BodyHtml $bodyHtml -ReplyTo $Config.gmail_address -Attachments @($educationalSummaryPath)
        $job.confirmation.sent_at = (Get-Date).ToString('o')
        ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Confirmation email sent to $confirmationTo."
    } catch {
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "FAILED to send confirmation email: $_"
        Send-OperatorAlert -Config $Config -Subject "$firstName $lastName -- confirmation email failed to send" `
            -Message "Job $jobId was created but the confirmation email failed to send:`n$_"
    }

    # Shadow bureau freeze confirmation is a hard gate, not a courtesy flag --
    # an unfrozen bureau can cause a challenge to be overlooked as "already
    # verified", so an incomplete job does not enter the analysis queue.
    # TestMode bypasses this (like it bypasses the delivery delay) so the
    # operator can exercise the pipeline end-to-end without a second, real
    # confirmation submission from confirm-freeze.html.
    if ($shadowBureauIncomplete -and -not $TestMode) {
        $jobFolder = Move-JobFolder -Config $Config -JobId $jobId -FromStage $Stage -ToStage 'hold'
        $jobJsonPath = Join-Path $jobFolder 'job.json'
        $job.status = 'hold'
        ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8

        Start-ShadowBureauReminders -Config $Config -JobId $jobId
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Shadow bureau freeze confirmation incomplete -- job held out of the analysis queue. Reminders scheduled at +$($Config.shadow_bureau_reminder_1hr_delay_hours)h and +$($Config.shadow_bureau_reminder_24hr_delay_hours)h; manual follow-up flag at +$($Config.shadow_bureau_manual_flag_hours)h."
    } else {
        if ($shadowBureauIncomplete) {
            Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message 'Shadow bureau freeze confirmation incomplete, but TestMode is on -- gate bypassed, proceeding to analysis immediately.'
        }
        Start-AnalysisJob -JobId $jobId -TestMode:$TestMode
        Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Handed off to analyze.ps1."
    }

}

function Find-HeldShadowBureauJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Email
    )

    $holdRoot = Join-Path $Config.job_storage_path 'hold'
    if (-not (Test-Path -LiteralPath $holdRoot)) { return $null }

    $normalizedEmail = $Email.Trim().ToLowerInvariant()
    $candidates = Get-ChildItem -LiteralPath $holdRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $candidateJsonPath = Join-Path $_.FullName 'job.json'
        if (-not (Test-Path -LiteralPath $candidateJsonPath)) { return }
        try {
            $candidateJob = Get-Content -LiteralPath $candidateJsonPath -Raw | ConvertFrom-Json
        } catch {
            return
        }
        $candidateEmail = [string]$candidateJob.client.email
        if ($candidateJob.shadow_bureau.incomplete -and $candidateEmail -and ($candidateEmail.Trim().ToLowerInvariant() -eq $normalizedEmail)) {
            [pscustomobject]@{ Job = $candidateJob; Folder = $_.FullName; JobJsonPath = $candidateJsonPath }
        }
    }

    # Handles multiple submissions from the same address: job_id is
    # timestamp-prefixed (New-JobId, common.ps1), so sorting the id string
    # descending picks the most recently uploaded matching job.
    return $candidates | Sort-Object { $_.Job.job_id } -Descending | Select-Object -First 1
}

function Get-ShadowBureauFollowupTaskName {
    param([string]$JobId, [string]$Suffix)
    return "CreditCleared-ShadowBureau-$Suffix-$JobId"
}

function Start-ShadowBureauReminders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$JobId
    )

    $followupScript = Join-Path $ScriptRoot 'shadow-bureau-followup.ps1'
    $schedule = @(
        @{ Suffix = 'Reminder1hr';  Hours = [double]$Config.shadow_bureau_reminder_1hr_delay_hours }
        @{ Suffix = 'Reminder24hr'; Hours = [double]$Config.shadow_bureau_reminder_24hr_delay_hours }
        @{ Suffix = 'ManualFlag';   Hours = [double]$Config.shadow_bureau_manual_flag_hours }
    )

    foreach ($item in $schedule) {
        $taskName = Get-ShadowBureauFollowupTaskName -JobId $JobId -Suffix $item.Suffix
        $fireAt = (Get-Date).AddHours($item.Hours)
        $argumentList = "-NoProfile -File `"$followupScript`" -JobId `"$JobId`" -Kind $($item.Suffix)"
        $action = New-ScheduledTaskAction -Execute 'pwsh' -Argument $argumentList
        $trigger = New-ScheduledTaskTrigger -Once -At $fireAt
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
            -Description "Credit Cleared shadow bureau $($item.Suffix) for job $JobId" -Force | Out-Null
    }
}

function Stop-ShadowBureauReminders {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$JobId)

    foreach ($suffix in @('Reminder1hr', 'Reminder24hr', 'ManualFlag')) {
        $taskName = Get-ShadowBureauFollowupTaskName -JobId $JobId -Suffix $suffix
        try {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
        } catch {
            # Not registered -- never scheduled (e.g. TestMode), or already
            # fired/removed. Nothing to do.
        }
    }
}

function Invoke-ShadowBureauConfirmation {
    [CmdletBinding()]
    param($Context, $Config, [switch]$TestMode, $Submission, [string]$RemoteIp, [string]$BodyText)

    $email = [string]$Submission.client_email

    $checklist = Get-ShadowBureauChecklist -Submission $Submission
    foreach ($key in $checklist.Keys) {
        $status = if ($checklist[$key]) { 'checked' } else { 'unchecked' }
        Write-Host "Shadow bureau confirmation from $RemoteIp for '$email' -- $($Script:ShadowBureaus[$key]): $status."
    }
    $allChecked = -not ($checklist.Values -contains $false)

    if ([string]::IsNullOrWhiteSpace($email)) {
        Write-Warning "Shadow bureau confirmation from $RemoteIp has no client_email -- cannot proceed."
        Send-OperatorAlert -Config $Config -Subject 'Shadow bureau confirmation missing client_email -- check manually' `
            -Message "A shadow bureau confirmation submission came in without a usable client_email.`n`nRaw payload:`n$BodyText"
        Send-HttpResponse -Context $Context -StatusCode 200 -Body 'Received (alert sent -- missing client_email)'
        return
    }

    if (-not $allChecked) {
        # confirm-freeze.html disables submit until all 5 boxes are checked,
        # so this should be unreachable via the real page -- treat it as
        # tampered or malformed input and do NOT release any job.
        Write-Warning "Shadow bureau confirmation from $RemoteIp for '$email' arrived with not all 5 boxes checked -- not releasing any job."
        Send-OperatorAlert -Config $Config -Subject "Shadow bureau confirmation incomplete -- check manually ($email)" `
            -Message "A shadow bureau confirmation submission for $email arrived without all 5 boxes checked. This should not be possible via confirm-freeze.html -- check for tampering or a bug.`n`nRaw payload:`n$BodyText"
        Send-HttpResponse -Context $Context -StatusCode 200 -Body 'Received (alert sent -- confirmation incomplete)'
        return
    }

    $match = Find-HeldShadowBureauJob -Config $Config -Email $email
    if (-not $match) {
        Write-Warning "Shadow bureau confirmation from $RemoteIp for '$email' has no matching held job."
        Send-OperatorAlert -Config $Config -Subject "Shadow bureau confirmation for $email -- no matching held job found" `
            -Message "A shadow bureau confirmation came in for $email but no on-hold job for that email was found (already released, wrong email, or a race). Check manually.`n`nRaw payload:`n$BodyText"
        Send-HttpResponse -Context $Context -StatusCode 200 -Body 'Received (alert sent -- no matching job)'
        return
    }

    $job = $match.Job
    $jobId = $job.job_id

    $job.shadow_bureau.incomplete = $false
    $job.shadow_bureau.checklist = $checklist
    # Server time, not the client-supplied confirmed_at field on the
    # submission -- client clocks aren't trustworthy for the authoritative record.
    $job.shadow_bureau.confirmed_at = (Get-Date).ToString('o')

    # Release into the analysis queue timestamped from confirmation, not the
    # original upload: Move-JobFolder puts it back in 'queued', and
    # analyze.ps1's delivery-time math runs relative to whenever it actually
    # executes (i.e. now), so no separate timestamp plumbing is needed.
    $destFolder = Move-JobFolder -Config $Config -JobId $jobId -FromStage 'hold' -ToStage 'queued'
    $jobJsonPath = Join-Path $destFolder 'job.json'
    $job.status = 'queued'
    ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8

    Stop-ShadowBureauReminders -JobId $jobId
    Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message "Shadow bureau freeze confirmed by $email from $RemoteIp. Released into analysis queue and scheduled reminders/manual-flag cancelled."

    Start-AnalysisJob -JobId $jobId -TestMode:$TestMode
    Write-JobLog -JobId $jobId -LogPath $Config.log_path -Message 'Handed off to analyze.ps1 after shadow bureau confirmation.'

    Send-HttpResponse -Context $Context -StatusCode 200 -Body "Confirmed. Job released: $jobId"
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
