<#
.SYNOPSIS
    Credit Cleared delivery module.
.DESCRIPTION
    Triggered by the scheduled task analyze.ps1 registers. Sends the five
    roadmap documents to the client, notifies the operator, and archives the
    job. See spec Section 7.
.PARAMETER JobId
    The job to deliver.
.PARAMETER TestMode
    Job lives in jobs\test\; delivery goes to the operator instead of the
    client, and the job is not moved to jobs\archive\ afterward.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$JobId,
    [switch]$TestMode
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = $PSScriptRoot
. (Join-Path $ScriptRoot 'common.ps1')
$Config = Get-CreditClearedConfig -Path (Join-Path $ScriptRoot 'config.json')

$stage = if ($TestMode) { 'test' } else { 'completed' }
$jobFolder = Get-JobFolder -Config $Config -Stage $stage -JobId $JobId
if (-not (Test-Path -LiteralPath $jobFolder)) {
    throw "Job folder not found for '$JobId' in stage '$stage' ($jobFolder)."
}

$jobJsonPath = Join-Path $jobFolder 'job.json'
$job = Get-Content -LiteralPath $jobJsonPath -Raw | ConvertFrom-Json

function Save-Job {
    ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8
}

Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message 'deliver.ps1 started.'

$docsDir = Join-Path $jobFolder 'docs'
$attachments = @(
    'Credit Cleared Roadmap.docx',
    'Call Scripts.docx',
    'Dispute Letters.docx',
    'Goodwill Letters.docx',
    'Debt Validation Letters.docx'
) | ForEach-Object { Join-Path $docsDir $_ }

$missingDocs = $attachments | Where-Object { -not (Test-Path -LiteralPath $_) }
if ($missingDocs.Count -gt 0) {
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "FAILED: missing document(s): $($missingDocs -join '; ')"
    Send-OperatorAlert -Config $Config `
        -Subject "$($job.client.first_name) $($job.client.last_name) -- delivery blocked, documents missing ($JobId)" `
        -Message "deliver.ps1 ran for job $JobId but the following document(s) were not found:`n$($missingDocs -join "`n")`n`nJob folder: $jobFolder"
    exit 1
}

$deliveryTo = if ($TestMode) { $Config.your_email } else { $job.client.email }
$subject = 'Your Credit Cleared Roadmap is here'
if ($TestMode) { $subject = "[TEST] $subject" }

$bodyHtml = Expand-Template -TemplatePath (Join-Path $Config.template_path 'delivery-email.html') -Tokens @{
    CLIENT_FIRST_NAME = [string]$job.client.first_name
    YOUR_NAME         = $Config.your_name
    YOUR_PHONE        = $Config.your_phone
    BUSINESS_DBA      = $Config.business_dba
    BUSINESS_NAME     = $Config.business_name
}

# Section 8: "Delivery email fails to send -> Retry after 30 minutes. If
# still failing after 3 attempts, email alert with all documents attached
# so you can forward manually."
$maxAttempts = 3
$delivered = $false
for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    try {
        Send-CreditClearedEmail -Config $Config -To $deliveryTo -Subject $subject -BodyHtml $bodyHtml `
            -Attachments $attachments -ReplyTo $Config.gmail_address
        $delivered = $true
        break
    } catch {
        Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Delivery attempt $attempt of $maxAttempts failed: $_"
        if ($attempt -lt $maxAttempts) {
            Start-Sleep -Seconds 1800
        }
    }
}

if (-not $delivered) {
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "FAILED: delivery did not succeed after $maxAttempts attempts."
    Send-OperatorAlert -Config $Config `
        -Subject "$($job.client.first_name) $($job.client.last_name) -- delivery failed after $maxAttempts attempts ($JobId)" `
        -Message "deliver.ps1 could not send the delivery email for job $JobId after $maxAttempts attempts. All documents are attached to this alert so you can forward them manually.`n`nJob folder: $jobFolder" `
        -Attachments $attachments
    exit 1
}

$job.delivery.delivered_at = (Get-Date).ToString('o')
$job.status = if ($TestMode) { 'delivered-test' } else { 'delivered' }
Save-Job
Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Delivered to $deliveryTo."

Send-OperatorAlert -Config $Config `
    -Subject "$($job.client.first_name) $($job.client.last_name) roadmap delivered successfully" `
    -Message "Job $JobId delivered to $deliveryTo at $((Get-Date).ToString('u'))."

if (-not $TestMode) {
    $archiveFolder = Move-JobFolder -Config $Config -JobId $JobId -FromStage 'completed' -ToStage 'archive'
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Job archived: $archiveFolder."
}
