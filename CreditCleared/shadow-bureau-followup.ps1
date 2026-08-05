<#
.SYNOPSIS
    Credit Cleared shadow bureau follow-up module.
.DESCRIPTION
    Invoked by the scheduled tasks main.ps1 registers (Start-ShadowBureauReminders)
    when a job is held out of the analysis queue for an incomplete shadow
    bureau freeze confirmation. Three kinds of firing:
      Reminder1hr / Reminder24hr -- emails the client a reminder with a link
        back to confirm-freeze.html.
      ManualFlag -- alerts the operator that a job is still unconfirmed after
        the configured window, for manual follow-up.
    If the job has already been confirmed (or the job folder is otherwise
    gone from jobs\hold\) by the time this fires, it's a no-op -- main.ps1
    cancels these scheduled tasks on confirmation, but this script re-checks
    job state itself as a defense against races between a task firing and
    the cancellation landing.
.PARAMETER JobId
    The job to check, as created by main.ps1.
.PARAMETER Kind
    Which follow-up this firing is: Reminder1hr, Reminder24hr, or ManualFlag.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$JobId,
    [Parameter(Mandatory)][ValidateSet('Reminder1hr', 'Reminder24hr', 'ManualFlag')][string]$Kind
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = $PSScriptRoot
. (Join-Path $ScriptRoot 'common.ps1')
$Config = Get-CreditClearedConfig -Path (Join-Path $ScriptRoot 'config.json')

$jobFolder = Get-JobFolder -Config $Config -Stage 'hold' -JobId $JobId
if (-not (Test-Path -LiteralPath $jobFolder)) {
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "shadow-bureau-followup.ps1 ($Kind) skipped -- job is no longer on hold (already confirmed or otherwise moved)."
    return
}

$jobJsonPath = Join-Path $jobFolder 'job.json'
$job = Get-Content -LiteralPath $jobJsonPath -Raw | ConvertFrom-Json

if (-not $job.shadow_bureau.incomplete) {
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "shadow-bureau-followup.ps1 ($Kind) skipped -- job is no longer marked incomplete."
    return
}

if ($Kind -eq 'ManualFlag') {
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Shadow bureau confirmation still incomplete after $($Config.shadow_bureau_manual_flag_hours) hours -- flagging operator for manual follow-up."
    Send-OperatorAlert -Config $Config `
        -Subject "$($job.client.first_name) $($job.client.last_name) -- shadow bureau freeze still unconfirmed after $($Config.shadow_bureau_manual_flag_hours)h ($JobId)" `
        -Message "Job $JobId (client: $($job.client.email)) has not confirmed their shadow bureau freeze $($Config.shadow_bureau_manual_flag_hours) hours after upload and remains held out of the analysis queue.`n`nJob folder: $jobFolder`n`nFollow up manually."
    return
}

$confirmUri = New-Object System.Uri([Uri]$Config.upload_page_url, 'confirm-freeze.html')
$confirmUrl = "$confirmUri" + '?email=' + [Uri]::EscapeDataString([string]$job.client.email)

$subject = if ($Kind -eq 'Reminder1hr') {
    'Quick reminder — confirm your shadow bureau freezes'
} else {
    'Still need this — your roadmap is on hold until you confirm'
}

try {
    $bodyHtml = Expand-Template -TemplatePath $Config.shadow_bureau_reminder_template_path -Tokens @{
        CLIENT_FIRST_NAME = [string]$job.client.first_name
        CONFIRM_URL       = $confirmUrl
        YOUR_NAME         = $Config.your_name
        YOUR_PHONE        = $Config.your_phone
        BUSINESS_DBA      = $Config.business_dba
        BUSINESS_NAME     = $Config.business_name
    }
    Send-CreditClearedEmail -Config $Config -To $job.client.email -Subject $subject -BodyHtml $bodyHtml -ReplyTo $Config.gmail_address
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Shadow bureau $Kind reminder sent to $($job.client.email)."
} catch {
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "FAILED to send shadow bureau $Kind reminder: $_"
    Send-OperatorAlert -Config $Config `
        -Subject "$($job.client.first_name) $($job.client.last_name) -- shadow bureau $Kind reminder failed to send ($JobId)" `
        -Message "shadow-bureau-followup.ps1 ($Kind) could not send the reminder email for job $JobId to $($job.client.email):`n$_"
}
