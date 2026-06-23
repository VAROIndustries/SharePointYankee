<#
.SYNOPSIS
    Sends email via the SendGrid v3 REST API.

.DESCRIPTION
    Uses the SendGrid v3 /mail/send REST endpoint to send transactional email.
    Supports plain text and HTML bodies, multiple recipients, and optional file
    attachments. The API key can be supplied as a parameter or read automatically
    from the SENDGRID_API_KEY environment variable, making it safe for use in
    CI/CD pipelines and Azure Automation without hardcoding credentials.

    This script intentionally avoids Send-MailMessage (deprecated in PowerShell 7)
    and SMTP relay in favor of the authenticated REST API, which is more reliable
    and does not require firewall rules for port 587.

.PARAMETER ApiKey
    SendGrid API key with at least the "Mail Send" permission. If omitted, the
    function falls back to the SENDGRID_API_KEY environment variable.

.PARAMETER From
    Sender email address. Must be a verified sender or domain in your SendGrid account.

.PARAMETER FromName
    Optional display name for the sender. Defaults to the From address.

.PARAMETER To
    One or more recipient email addresses. Accepts a string array.

.PARAMETER Subject
    Email subject line.

.PARAMETER Body
    Email body content. Use -Html to render as HTML; otherwise sent as plain text.

.PARAMETER Html
    Switch. When present, the Body parameter is sent as text/html. When absent,
    body is sent as text/plain.

.PARAMETER AttachmentPath
    Optional. Full path to a file to attach. The file is Base64-encoded and sent
    as a single attachment. For multiple attachments call the function multiple
    times or extend the attachments array in the body hashtable.

.EXAMPLE
    Send-SendGridEmail -From "alerts@contoso.com" -To "admin@contoso.com" `
        -Subject "Daily Report" -Body "The report is attached." `
        -AttachmentPath "C:\Reports\daily.csv"

    Sends a plain-text email with a CSV attachment using the API key from the
    SENDGRID_API_KEY environment variable.

.EXAMPLE
    $key = (Get-AzKeyVaultSecret -VaultName "corp-kv" -Name "SendGridKey").SecretValue |
        ConvertFrom-SecureString -AsPlainText
    Send-SendGridEmail -ApiKey $key -From "no-reply@contoso.com" `
        -To @("alice@contoso.com","bob@contoso.com") `
        -Subject "Welcome" -Body "<h1>Hello!</h1>" -Html

    Retrieves the API key from Azure Key Vault and sends an HTML email to two recipients.

.NOTES
    Author  : Geoff Varosky
    Version : 1.0.0
    Requires: PowerShell 5.1 or 7+; no external modules needed.
    API Ref : https://docs.sendgrid.com/api-reference/mail-send/mail-send
    GitHub  : https://github.com/VAROIndustries/SharePointYankee
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter()]
    [string]$ApiKey = $env:SENDGRID_API_KEY,

    [Parameter(Mandatory)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$From,

    [Parameter()]
    [string]$FromName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$To,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Subject,

    [Parameter(Mandatory)]
    [string]$Body,

    [Parameter()]
    [switch]$Html,

    [Parameter()]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$AttachmentPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Validate API key
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw 'No SendGrid API key provided. Supply -ApiKey or set the SENDGRID_API_KEY environment variable.'
}
#endregion

#region Build recipient list
$toAddresses = foreach ($address in $To) {
    @{ email = $address.Trim() }
}
#endregion

#region Build personalization and message body
$personalization = @{
    to = @($toAddresses)
}

$contentType = if ($Html) { 'text/html' } else { 'text/plain' }

$senderObject = @{ email = $From }
if ($PSBoundParameters.ContainsKey('FromName') -and -not [string]::IsNullOrWhiteSpace($FromName)) {
    $senderObject['name'] = $FromName
}

$messageBody = @{
    personalizations = @($personalization)
    from             = $senderObject
    subject          = $Subject
    content          = @(
        @{
            type  = $contentType
            value = $Body
        }
    )
}
#endregion

#region Handle attachment
if ($PSBoundParameters.ContainsKey('AttachmentPath')) {
    Write-Verbose "Reading attachment: $AttachmentPath"
    $fileBytes    = [System.IO.File]::ReadAllBytes($AttachmentPath)
    $encoded      = [Convert]::ToBase64String($fileBytes)
    $fileName     = [System.IO.Path]::GetFileName($AttachmentPath)

    # Attempt to determine MIME type from extension; fall back to octet-stream
    $mimeMap = @{
        '.pdf'  = 'application/pdf'
        '.csv'  = 'text/csv'
        '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        '.xls'  = 'application/vnd.ms-excel'
        '.docx' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        '.txt'  = 'text/plain'
        '.html' = 'text/html'
        '.htm'  = 'text/html'
        '.zip'  = 'application/zip'
        '.json' = 'application/json'
        '.png'  = 'image/png'
        '.jpg'  = 'image/jpeg'
        '.jpeg' = 'image/jpeg'
        '.gif'  = 'image/gif'
    }
    $ext      = [System.IO.Path]::GetExtension($AttachmentPath).ToLower()
    $mimeType = if ($mimeMap.ContainsKey($ext)) { $mimeMap[$ext] } else { 'application/octet-stream' }

    $messageBody['attachments'] = @(
        @{
            content     = $encoded
            filename    = $fileName
            type        = $mimeType
            disposition = 'attachment'
        }
    )
    Write-Verbose "Attachment MIME type resolved as: $mimeType"
}
#endregion

#region Send the request
$jsonBody = $messageBody | ConvertTo-Json -Depth 10 -Compress
$headers  = @{
    'Authorization' = "Bearer $ApiKey"
    'Content-Type'  = 'application/json'
}

$recipientDisplay = $To -join ', '

if ($PSCmdlet.ShouldProcess($recipientDisplay, "Send email via SendGrid: '$Subject'")) {
    try {
        Write-Verbose "POST https://api.sendgrid.com/v3/mail/send"
        Write-Verbose "To: $recipientDisplay | Subject: $Subject | ContentType: $contentType"

        $response = Invoke-RestMethod `
            -Uri     'https://api.sendgrid.com/v3/mail/send' `
            -Method  Post `
            -Headers $headers `
            -Body    $jsonBody

        # SendGrid returns HTTP 202 Accepted with an empty body on success.
        # Invoke-RestMethod does not throw for 2xx responses so we reach here on success.
        Write-Verbose 'Email accepted by SendGrid (202).'
        Write-Output "Email successfully submitted to SendGrid for delivery to: $recipientDisplay"
    }
    catch {
        # Unpack SendGrid's error body which Invoke-RestMethod wraps in the exception
        $statusCode   = $_.Exception.Response.StatusCode.value__
        $errorMessage = $_.ErrorDetails.Message

        if ($errorMessage) {
            try {
                $parsed = $errorMessage | ConvertFrom-Json
                $detail = ($parsed.errors | ForEach-Object { $_.message }) -join '; '
                throw "SendGrid API error $statusCode`: $detail"
            }
            catch [System.Management.Automation.RuntimeException] {
                throw   # Re-throw our formatted error
            }
            catch {
                throw "SendGrid API error $statusCode`: $errorMessage"
            }
        }
        else {
            throw "SendGrid API call failed: $_"
        }
    }
}
#endregion
