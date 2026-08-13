<#
.SYNOPSIS
    Credit Cleared analysis module.
.DESCRIPTION
    Reads a queued job, sends the credit report to Claude for the Consumer
    Analysis (spec Section 5 / CC-Consumer-Analysis-Prompt.docx), generates the
    five roadmap documents via generate-docs.py, then schedules delivery.

    Per the upsell-only decision, this module runs ONLY the $297 Consumer
    Analysis call. The Business Funding Eligibility Analysis ($197 add-on) is
    a separate, on-demand script run against jobs\archive\[job-id]\ when a
    client purchases it -- it is not part of this automatic pipeline.
.PARAMETER JobId
    The job to analyze, as created by main.ps1.
.PARAMETER TestMode
    Job lives in jobs\test\ instead of jobs\queued\/processing\/completed\,
    and delivery is scheduled 2 minutes out instead of the normal delay.
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

$sourceStage = if ($TestMode) { 'test' } else { 'queued' }
$jobFolder = Get-JobFolder -Config $Config -Stage $sourceStage -JobId $JobId
if (-not (Test-Path -LiteralPath $jobFolder)) {
    throw "Job folder not found for '$JobId' in stage '$sourceStage' ($jobFolder)."
}

$jobJsonPath = Join-Path $jobFolder 'job.json'
$job = Get-Content -LiteralPath $jobJsonPath -Raw | ConvertFrom-Json

Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message 'analyze.ps1 started.'

function Save-Job {
    ($job | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jobJsonPath -Encoding UTF8
}

function Move-ToFailed {
    param(
        [string]$Reason,
        [string]$CurrentStage,
        [string[]]$Attachments = @()
    )

    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "FAILED: $Reason"
    if ($TestMode) {
        $job.status = 'failed'
        Save-Job
        $destFolder = $jobFolder
    } else {
        $destFolder = Move-JobFolder -Config $Config -JobId $JobId -FromStage $CurrentStage -ToStage 'failed'
        $script:jobJsonPath = Join-Path $destFolder 'job.json'
        $job.status = 'failed'
        Save-Job
    }
    Send-OperatorAlert -Config $Config `
        -Subject "$($job.client.first_name) $($job.client.last_name) -- analysis failed ($JobId)" `
        -Message "Job $JobId failed during analyze.ps1:`n$Reason`n`nJob folder: $destFolder" `
        -Attachments $Attachments
}

# Move queued -> processing. Test-mode jobs stay in jobs\test\ throughout
# (spec Section 10: "All documents saved to jobs\test\ folder").
if (-not $TestMode) {
    $jobFolder = Move-JobFolder -Config $Config -JobId $JobId -FromStage 'queued' -ToStage 'processing'
    $jobJsonPath = Join-Path $jobFolder 'job.json'
}
$job.status = if ($TestMode) { 'test' } else { 'processing' }
Save-Job

$currentStage = if ($TestMode) { 'test' } else { 'processing' }

# ---- Consumer Analysis (Claude API) ----------------------------------------
try {
    $reportPath = Join-Path $jobFolder $job.files.report_pdf
    $reportBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($reportPath))

    $systemPrompt = Get-Content -LiteralPath (Join-Path $Config.template_path 'consumer-analysis-system-prompt.txt') -Raw

    $userPrompt = Expand-Template -TemplatePath (Join-Path $Config.template_path 'consumer-analysis-user-prompt.txt') -Tokens @{
        CLIENT_FIRST_NAME = [string]$job.client.first_name
        CLIENT_LAST_NAME  = [string]$job.client.last_name
        GOAL_FROM_FORM    = $(if ($job.client.goal) { $job.client.goal } else { 'Not specified' })
        DEADLINE_FROM_FORM = $(if ($job.client.deadline) { $job.client.deadline } else { 'Not specified' })
        NOTES_FROM_FORM   = $(if ($job.client.notes) { $job.client.notes } else { 'None provided' })
    }

    $messages = @(
        @{
            role    = 'user'
            content = @(
                @{ type = 'document'; source = @{ type = 'base64'; media_type = 'application/pdf'; data = $reportBase64 } }
                @{ type = 'text'; text = $userPrompt }
            )
        }
    )

    $maxTokens = [int]$Config.consumer_analysis_max_tokens
    $temperature = [double]$Config.temperature

    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Calling Claude API ($($Config.claude_model), max_tokens=$maxTokens)."
    $response = Invoke-ClaudeMessage -Config $Config -SystemPrompt $systemPrompt -Messages $messages -MaxTokens $maxTokens -Temperature $temperature
    $analysisText = ($response.content | Where-Object { $_.type -eq 'text' } | Select-Object -First 1).text

    # Section 8: "API returns incomplete response -> check for all required
    # sections; if any missing, retry with a follow-up prompt requesting them."
    $requiredSections = @(
        'SECTION 1: ACCOUNT SUMMARY', 'SECTION 2: PRIORITY ACTION PLAN', 'SECTION 3: CREDITOR CONTACT DIRECTORY',
        'SECTION 4: WORD-FOR-WORD CALL SCRIPTS', 'SECTION 5: DISPUTE LETTERS', 'SECTION 6: GOODWILL LETTERS',
        'SECTION 7: DEBT VALIDATION LETTERS', 'SECTION 8: INQUIRY CROSS-REFERENCE ANALYSIS',
        'SECTION 9: MONTH-BY-MONTH MILESTONE MAP', 'SECTION 10: PERSONAL CREDIT PROFILE BUILDING'
    )
    $missing = $requiredSections | Where-Object { $analysisText -notmatch [regex]::Escape($_) }

    if ($missing.Count -gt 0) {
        Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Response missing sections: $($missing -join '; '). Requesting follow-up."
        $messages += @{ role = 'assistant'; content = $analysisText }
        $messages += @{
            role    = 'user'
            content = "Your previous response was missing the following required sections: $($missing -join ', '). Provide ONLY those missing sections now, using the exact section headers shown in my original instructions."
        }
        $followUp = Invoke-ClaudeMessage -Config $Config -SystemPrompt $systemPrompt -Messages $messages -MaxTokens $maxTokens -Temperature $temperature
        $followUpText = ($followUp.content | Where-Object { $_.type -eq 'text' } | Select-Object -First 1).text
        $analysisText = "$analysisText`n`n$followUpText"

        $stillMissing = $requiredSections | Where-Object { $analysisText -notmatch [regex]::Escape($_) }
        if ($stillMissing.Count -gt 0) {
            Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Still missing after follow-up: $($stillMissing -join '; '). Proceeding -- flagged for operator review."
            Send-OperatorAlert -Config $Config `
                -Subject "$($job.client.first_name) $($job.client.last_name) -- roadmap missing sections ($JobId)" `
                -Message "After a follow-up request, the analysis for job $JobId is still missing: $($stillMissing -join ', ')`n`nProceeding with document generation -- please review the output manually."
        }
    }

    $responseFileName = 'consumer-analysis-response.json'
    @{
        raw_text     = $analysisText
        model        = $Config.claude_model
        completed_at = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $jobFolder $responseFileName) -Encoding UTF8

    $job.analysis.consumer_completed_at = (Get-Date).ToString('o')
    $job.analysis.consumer_response_file = $responseFileName
    Save-Job
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Consumer analysis complete ($($analysisText.Length) chars)."
} catch {
    Move-ToFailed -Reason "Claude API analysis failed: $_" -CurrentStage $currentStage
    return
}

# ---- Document generation ----------------------------------------------------
try {
    $genScript = Join-Path $ScriptRoot 'generate-docs.py'
    if (-not (Test-Path -LiteralPath $genScript)) {
        Write-Warning "generate-docs.py not found yet -- analysis for job '$JobId' is complete but documents were not generated. This step will need to be re-run once generate-docs.py exists."
        Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message 'generate-docs.py not found -- skipping document generation for now.'
    } else {
        Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message 'Calling generate-docs.py.'
        $pyOutput = & python $genScript --job-folder $jobFolder --job-id $JobId 2>&1
        $pyOutput | ForEach-Object { Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "[generate-docs.py] $_" }
        if ($LASTEXITCODE -ne 0) {
            throw "generate-docs.py exited with code $LASTEXITCODE"
        }
        Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message 'Document generation complete.'
    }
} catch {
    # Section 8: "Document generation fails -> log, move to failed\, email
    # operator with the Claude API JSON output so they can generate manually."
    $responsePath = Join-Path $jobFolder 'consumer-analysis-response.json'
    Move-ToFailed -Reason "Document generation failed: $_" -CurrentStage $currentStage -Attachments @($responsePath)
    return
}

# ---- Move to completed and schedule delivery --------------------------------
if (-not $TestMode) {
    $jobFolder = Move-JobFolder -Config $Config -JobId $JobId -FromStage 'processing' -ToStage 'completed'
    $jobJsonPath = Join-Path $jobFolder 'job.json'
}
$job.status = if ($TestMode) { 'test' } else { 'completed' }

# Delivery timing (spec Section 7). NOTE: the spec's own worked examples
# ("Monday 2pm -> Tuesday 9am", "Monday 10pm -> Tuesday 9am", "Friday 8pm ->
# Saturday 9am") are inconsistent with its prose ("completion + 24 hours,
# then next 9am after that" would put the Monday-2pm case at Wednesday 9am,
# not Tuesday). All three examples match a simpler rule instead: always the
# NEXT calendar day at the configured time. That's what's implemented below
# -- it keeps delivery comfortably under the "within 48 hours" promise made
# in the confirmation email and the funding upsell copy, which the literal
# prose reading could occasionally violate.
try {
    if ($TestMode) {
        $deliveryTime = (Get-Date).AddMinutes(2)
    } else {
        $nextDay = (Get-Date).Date.AddDays(1)
        $deliveryTime = Get-Date -Year $nextDay.Year -Month $nextDay.Month -Day $nextDay.Day `
            -Hour ([int]$Config.delivery_time_hour) -Minute ([int]$Config.delivery_time_minute) -Second 0
    }

    $deliverScript = Join-Path $ScriptRoot 'deliver.ps1'
    $taskName = "CreditCleared-Deliver-$JobId"
    $command = "pwsh -NoProfile -File `"$deliverScript`" -JobId `"$JobId`""
    if ($TestMode) { $command += ' -TestMode' }

    Register-CreditClearedAtJob -TaskName $taskName -Command $command -FireAt $deliveryTime -JobStoragePath $Config.job_storage_path

    $job.delivery.scheduled_at = $deliveryTime.ToString('o')
    Save-Job
    Write-JobLog -JobId $JobId -LogPath $Config.log_path -Message "Delivery scheduled for $($deliveryTime.ToString('u')) via task '$taskName'."
} catch {
    Move-ToFailed -Reason "Failed to schedule delivery: $_" -CurrentStage 'completed'
    return
}
