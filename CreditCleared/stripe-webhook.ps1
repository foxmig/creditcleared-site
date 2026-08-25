<#
.SYNOPSIS
    Credit Cleared Stripe checkout webhook handler.
.DESCRIPTION
    Defines Invoke-StripeWebhookRequest and its helpers. Unlike analyze.ps1/
    deliver.ps1 (standalone scripts run as their own process), this file is a
    function library dot-sourced into main.ps1 -- Caddy proxies every path
    under webhook.mycreditcleared.com to the same upstream (127.0.0.1:8080),
    and only one process can bind that port, so main.ps1's single
    HttpListener registers a second prefix ("/stripe-webhook/") and
    dispatches to Invoke-StripeWebhookRequest by path instead of running a
    second listener. This means the functions here depend on main.ps1
    already being loaded when they're called -- specifically Send-HttpResponse
    (defined in main.ps1) alongside the usual common.ps1 helpers.
#>

function Test-StripeEventProcessed {
    # Marker-file dedupe under job_storage_path, the same general pattern
    # Register-CreditClearedAtJob (common.ps1) uses for .at-jobs state --
    # there's no per-event "job" folder for a Stripe sale the way there is
    # for a credit report upload, so a flat marker directory keyed on the
    # checkout session ID stands in for Find-RecentDuplicateJob's job.json
    # scan (main.ps1), which doesn't apply here.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$SessionId
    )

    $markerPath = Join-Path (Join-Path $Config.job_storage_path '.stripe-processed') "$SessionId.processed"
    return Test-Path -LiteralPath $markerPath
}

function Set-StripeEventProcessed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$SessionId
    )

    $markerDir = Join-Path $Config.job_storage_path '.stripe-processed'
    if (-not (Test-Path -LiteralPath $markerDir)) {
        New-Item -ItemType Directory -Path $markerDir -Force | Out-Null
    }
    Set-Content -LiteralPath (Join-Path $markerDir "$SessionId.processed") -Value (Get-Date).ToString('o')
}

function Get-StripeCheckoutTier {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Session)

    # ASSUMPTION -- unverified: a plain checkout.session.completed webhook
    # payload does not include line_items unless the session is retrieved
    # again with an expand, which this endpoint deliberately does not do
    # (spec: no extra API call). If line_items ever does show up with a
    # price ID here, prefer matching on it once real DIY/DWY Stripe price
    # IDs are known and added to config -- for now amount_total (always
    # present on the session object, no expansion needed) is the only
    # reliable signal, same spirit as Get-CreditReportBytes's payload-shape
    # assumption in common.ps1.
    if ($Session.line_items -and $Session.line_items.data -and $Session.line_items.data.Count -gt 0) {
        Write-Warning "Stripe session $($Session.id) included line_items but no price-ID-to-tier mapping is configured yet -- falling back to amount_total."
    }

    switch ([long]$Session.amount_total) {
        29700   { return 'DIY' }
        199700  { return 'DWY' }
        default { return $null }
    }
}

function Get-StripeTierLabel {
    param([Parameter(Mandatory)][string]$Tier)
    if ($Tier -eq 'DWY') { 'Done-With-You Intensive ($1,997)' } else { 'DIY Roadmap ($297)' }
}

function Send-StripeSaleAlert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Tier,
        [Parameter(Mandatory)][string]$CustomerEmail
    )

    $tierLabel = Get-StripeTierLabel -Tier $Tier
    $bodyHtml = "<p><strong>New sale:</strong> $tierLabel</p><p><strong>Customer:</strong> $CustomerEmail</p>"
    Send-CreditClearedEmail -Config $Config -To $Config.your_email -Subject "New sale -- $tierLabel" -BodyHtml $bodyHtml
}

function Invoke-StripeWebhookRequest {
    [CmdletBinding()]
    param($Context, $Config)

    $request = $Context.Request
    $remoteIp = $request.RemoteEndPoint.Address.ToString()

    if ($request.HttpMethod -ne 'POST') {
        Send-HttpResponse -Context $Context -StatusCode 405 -Body 'Method Not Allowed'
        return
    }

    # Read raw body bytes BEFORE any parsing -- signature verification must
    # run against the exact bytes Stripe sent, not a re-serialized copy.
    $memoryStream = New-Object System.IO.MemoryStream
    $request.InputStream.CopyTo($memoryStream)
    $bodyBytes = $memoryStream.ToArray()
    $memoryStream.Dispose()

    $signatureHeader = $request.Headers['Stripe-Signature']
    $signatureOk = Test-StripeSignature -Body $bodyBytes -Secret $Config.stripe_webhook_secret -SignatureHeaderValue $signatureHeader

    if (-not $signatureOk) {
        Write-Warning "Rejected Stripe webhook request from $remoteIp -- signature verification failed."
        Write-JobLog -JobId 'stripe' -LogPath $Config.log_path -Message "Rejected Stripe webhook from $remoteIp -- signature verification failed."
        Send-HttpResponse -Context $Context -StatusCode 400 -Body 'Invalid signature'
        return
    }

    $bodyText = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
    try {
        $event = $bodyText | ConvertFrom-Json
    } catch {
        Write-Warning "Rejected Stripe webhook from $remoteIp -- body is not valid JSON: $_"
        Write-JobLog -JobId 'stripe' -LogPath $Config.log_path -Message "Rejected Stripe webhook from $remoteIp -- body is not valid JSON: $_"
        Send-HttpResponse -Context $Context -StatusCode 400 -Body 'Invalid JSON'
        return
    }

    if ($event.type -ne 'checkout.session.completed') {
        # Ack anything else so Stripe doesn't retry -- this endpoint only
        # acts on checkout.session.completed today.
        Send-HttpResponse -Context $Context -StatusCode 200 -Body "Ignored event type: $($event.type)"
        return
    }

    $session = $event.data.object
    $sessionId = [string]$session.id
    $customerEmail = [string]$session.customer_details.email
    $customerName = [string]$session.customer_details.name
    $firstName = if ($customerName) { ($customerName.Trim() -split '\s+')[0] } else { 'there' }

    if (Test-StripeEventProcessed -Config $Config -SessionId $sessionId) {
        Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message "Duplicate Stripe webhook for session $sessionId (already processed) -- Stripe retry, skipped."
        Send-HttpResponse -Context $Context -StatusCode 200 -Body "Already processed: $sessionId"
        return
    }

    $tier = Get-StripeCheckoutTier -Session $session

    if (-not $tier) {
        Write-Warning "Stripe session $sessionId has unrecognized amount_total $($session.amount_total) -- cannot determine tier."
        Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message "Unrecognized amount_total $($session.amount_total) for session $sessionId -- alerting operator, not sending welcome email."
        Send-OperatorAlert -Config $Config -Subject 'Stripe sale with unrecognized amount -- check manually' `
            -Message "Checkout session $sessionId completed with amount_total $($session.amount_total), which doesn't match a known tier (29700 = DIY, 199700 = DWY). Customer: $customerEmail. Check manually."
        Set-StripeEventProcessed -Config $Config -SessionId $sessionId
        Send-HttpResponse -Context $Context -StatusCode 200 -Body "Received (unrecognized amount, alert sent): $sessionId"
        return
    }

    if ([string]::IsNullOrWhiteSpace($customerEmail)) {
        Write-Warning "Stripe session $sessionId ($tier) has no customer_details.email -- cannot send welcome email."
        Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message "Session $sessionId ($tier) has no customer email -- alerting operator."
        Send-OperatorAlert -Config $Config -Subject "$tier sale with no customer email -- check manually" `
            -Message "Checkout session $sessionId ($tier) completed but customer_details.email was missing. Check manually in the Stripe dashboard."
        Set-StripeEventProcessed -Config $Config -SessionId $sessionId
        Send-HttpResponse -Context $Context -StatusCode 200 -Body "Received (no customer email, alert sent): $sessionId"
        return
    }

    Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message "Stripe checkout completed. Tier: $tier. Customer: $customerEmail."
    Set-StripeEventProcessed -Config $Config -SessionId $sessionId

    # Ack Stripe immediately -- everything after this (SendGrid calls) is
    # slow-ish and must not block the response. The marker file above (not
    # response timing) is what actually guards against double-processing a
    # Stripe retry, same division of concerns as the Formspree webhook in
    # main.ps1: ack fast, dedupe via durable state, not via response speed.
    Send-HttpResponse -Context $Context -StatusCode 200 -Body "Received. Session: $sessionId"

    $calendlySection = if ($tier -eq 'DWY') {
        "<p><strong>Your 1-on-1 strategy call.</strong> As a Done-With-You Intensive client, you get a dedicated strategy call with me to map out your exact plan. Book a time that works for you: <a href=`"$($Config.calendly_link)`">$($Config.calendly_link)</a></p>"
    } else {
        ''
    }

    # upload_page_url is the site root -- same Uri-combine pattern main.ps1
    # uses to build the confirm-freeze link, not a raw pass-through.
    $uploadPageUri = New-Object System.Uri([Uri]$Config.upload_page_url, 'upload-page.html')

    try {
        $welcomeTemplate = Join-Path $Config.template_path 'purchase-welcome-email.html'
        $bodyHtml = Expand-Template -TemplatePath $welcomeTemplate -Tokens @{
            CLIENT_FIRST_NAME   = $firstName
            MYFREESCORENOW_LINK = $Config.myfreescorenow_link
            UPLOAD_PAGE_URL     = "$uploadPageUri"
            CALENDLY_SECTION    = $calendlySection
            YOUR_NAME           = $Config.your_name
            BUSINESS_DBA        = $Config.business_dba
            BUSINESS_NAME       = $Config.business_name
        }
        Send-CreditClearedEmail -Config $Config -To $customerEmail -Subject "Welcome to Credit Cleared -- here's what happens next" -BodyHtml $bodyHtml -ReplyTo $Config.gmail_address
        Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message "Welcome email sent to $customerEmail ($tier)."
    } catch {
        Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message "FAILED to send welcome email: $_"
        Send-OperatorAlert -Config $Config -Subject "$tier welcome email failed to send -- check manually" `
            -Message "Stripe session $sessionId ($tier, $customerEmail) completed but the welcome email failed to send:`n$_"
    }

    try {
        Send-StripeSaleAlert -Config $Config -Tier $tier -CustomerEmail $customerEmail
        Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message 'Internal sale alert sent.'
    } catch {
        Write-JobLog -JobId $sessionId -LogPath $Config.log_path -Message "FAILED to send internal sale alert: $_"
    }
}
