<#
.SYNOPSIS
    Retrieves HTTP response headers from one or more URLs.

.DESCRIPTION
    Get-HttpResponseHeaders sends an HTTP GET request to the target URL and returns
    each response header as a structured object with Key and Value properties.

    Use the -SecurityOnly switch to filter the output to headers that are relevant
    to web-application security hardening, such as Content-Security-Policy, HSTS,
    X-Frame-Options, and others from the OWASP recommended set.

    The original approach (System.Net.WebRequest) has been replaced with
    Invoke-WebRequest so that modern TLS negotiation, redirect following, and
    proxy awareness all work correctly out of the box.

.PARAMETER Url
    The URL to query.  Accepts pipeline input.

.PARAMETER TimeoutSeconds
    Maximum seconds to wait for a response.  Default is 30.

.PARAMETER SecurityOnly
    When specified, only security-relevant response headers are returned.
    The security header list includes:
        Content-Security-Policy, Strict-Transport-Security, X-Frame-Options,
        X-Content-Type-Options, X-XSS-Protection, Referrer-Policy,
        Permissions-Policy, Cross-Origin-Opener-Policy,
        Cross-Origin-Embedder-Policy, Cross-Origin-Resource-Policy,
        Cache-Control (when present on sensitive endpoints).

.EXAMPLE
    Get-HttpResponseHeaders -Url "https://www.contoso.com"

    Returns all response headers for the given URL.

.EXAMPLE
    "https://contoso.com","https://fabrikam.com" | Get-HttpResponseHeaders -SecurityOnly

    Returns only security-relevant headers for each URL piped in.

.EXAMPLE
    Get-HttpResponseHeaders "https://contoso.com" | Format-Table -AutoSize

    Formats headers in a table for quick review.

.NOTES
    Author  : Geoff Varosky
    Version : 1.0.0
    GitHub  : https://github.com/VAROIndustries/SharePointYankee

    Original concept: Get-HTTPHeaders.ps1 (System.Net.WebRequest approach).
    This version modernises the implementation using Invoke-WebRequest.

    Requires PowerShell 5.1 or later.  Compatible with PowerShell 7+.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(
        Mandatory,
        Position = 0,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        HelpMessage = 'URL to retrieve headers from (e.g. https://contoso.com)'
    )]
    [ValidateNotNullOrEmpty()]
    [string[]] $Url,

    [Parameter()]
    [ValidateRange(1, 300)]
    [int] $TimeoutSeconds = 30,

    [Parameter()]
    [switch] $SecurityOnly
)

begin {
    $ErrorActionPreference = 'Stop'

    # OWASP-aligned list of security-relevant response headers.
    # Checked as case-insensitive prefix/exact matches below.
    $securityHeaders = @(
        'Content-Security-Policy',
        'Strict-Transport-Security',
        'X-Frame-Options',
        'X-Content-Type-Options',
        'X-XSS-Protection',
        'Referrer-Policy',
        'Permissions-Policy',
        'Feature-Policy',                   # legacy predecessor to Permissions-Policy
        'Cross-Origin-Opener-Policy',
        'Cross-Origin-Embedder-Policy',
        'Cross-Origin-Resource-Policy',
        'Access-Control-Allow-Origin',
        'Cache-Control',
        'Expect-CT',                         # Certificate Transparency enforcement
        'Public-Key-Pins',                   # HPKP (deprecated but still seen in wild)
        'X-Permitted-Cross-Domain-Policies'
    )
}

process {
    foreach ($u in $Url) {
        Write-Verbose "Retrieving headers from: $u"

        try {
            $prevProgress    = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            $response = Invoke-WebRequest -Uri $u `
                                          -Method GET `
                                          -UseBasicParsing `
                                          -TimeoutSec $TimeoutSeconds `
                                          -ErrorAction Stop

            $ProgressPreference = $prevProgress

            Write-Verbose "HTTP $([int]$response.StatusCode) — $($response.Headers.Count) header(s) returned."

            foreach ($key in $response.Headers.Keys) {

                # When -SecurityOnly is active, skip headers not in our list
                if ($SecurityOnly) {
                    $matched = $false
                    foreach ($sh in $securityHeaders) {
                        if ($key -ieq $sh) {
                            $matched = $true
                            break
                        }
                    }
                    if (-not $matched) { continue }
                }

                # Headers can return as arrays (multiple values for same key)
                $rawValue = $response.Headers[$key]
                $displayValue = if ($rawValue -is [array]) {
                    $rawValue -join '; '
                }
                else {
                    [string]$rawValue
                }

                [PSCustomObject]@{
                    Url         = $u
                    Key         = $key
                    Value       = $displayValue
                    IsSecurity  = ($securityHeaders -icontains $key)
                }
            }

            # When -SecurityOnly is active, report any MISSING recommended headers
            # so the caller can immediately see gaps without a separate diff step.
            if ($SecurityOnly) {
                $returnedKeys = @($response.Headers.Keys | ForEach-Object { $_.ToLower() })
                foreach ($sh in $securityHeaders) {
                    if ($returnedKeys -notcontains $sh.ToLower()) {
                        [PSCustomObject]@{
                            Url        = $u
                            Key        = $sh
                            Value      = '** HEADER NOT PRESENT **'
                            IsSecurity = $true
                        }
                    }
                }
            }
        }
        catch [System.Net.WebException] {
            # A WebException may still have a usable response (e.g. HTTP 301, 403)
            if ($_.Exception.Response) {
                $httpResponse = [System.Net.HttpWebResponse]$_.Exception.Response
                Write-Warning "HTTP $([int]$httpResponse.StatusCode) ($($httpResponse.StatusDescription)) for '$u'. Headers may be incomplete."

                foreach ($key in $httpResponse.Headers.AllKeys) {
                    if ($SecurityOnly -and $securityHeaders -notcontains $key) { continue }

                    [PSCustomObject]@{
                        Url        = $u
                        Key        = $key
                        Value      = $httpResponse.Headers[$key]
                        IsSecurity = ($securityHeaders -icontains $key)
                    }
                }
            }
            else {
                Write-Error "Network error retrieving headers from '$u': $($_.Exception.Message)"
            }
        }
        catch {
            Write-Error "Unexpected error for '$u': $($_.Exception.Message)"
        }
    }
}

end {
    # Nothing to dispose — Invoke-WebRequest manages its own connections
}
