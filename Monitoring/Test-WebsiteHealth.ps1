<#
.SYNOPSIS
    Tests the health of one or more websites by checking HTTP status, response time,
    and SSL certificate expiry.

.DESCRIPTION
    Test-WebsiteHealth sends an HTTP GET request to each target URL and returns a
    structured result object containing the HTTP status code, response time in
    milliseconds, SSL certificate expiry date, and an overall IsHealthy flag.

    URLs can be supplied directly via the -Url parameter, piped in from the pipeline,
    or read from a CSV file using -InputFile.  The CSV must have a column named "Url".

    Use the -Detailed switch to include the full HTTP response-header collection on
    each output object.

.PARAMETER Url
    One or more URLs to test.  Accepts pipeline input.

.PARAMETER InputFile
    Path to a CSV file whose rows contain a "Url" column.

.PARAMETER TimeoutSeconds
    Maximum seconds to wait for a response.  Default is 30.

.PARAMETER Detailed
    When specified, the output objects include a Headers property containing all
    HTTP response headers as a hashtable.

.EXAMPLE
    Test-WebsiteHealth -Url "https://www.contoso.com"

    Tests a single URL and returns a health result object.

.EXAMPLE
    "https://contoso.com","https://fabrikam.com" | Test-WebsiteHealth

    Pipes multiple URLs through the pipeline.

.EXAMPLE
    Test-WebsiteHealth -InputFile "C:\sites.csv" -Detailed

    Reads URLs from a CSV and returns full header detail on each result.

.EXAMPLE
    Test-WebsiteHealth -Url "https://contoso.com" | Where-Object { -not $_.IsHealthy }

    Returns only unhealthy results.

.NOTES
    Author  : Geoff Varosky
    Version : 1.0.0
    GitHub  : https://github.com/VAROIndustries/SharePointYankee

    SSL expiry checking requires the target to use HTTPS.  HTTP targets will have
    SslExpiryDate set to $null and SslDaysRemaining set to $null.

    Requires PowerShell 5.1 or later.  Compatible with PowerShell 7+.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'ByUrl')]
param(
    [Parameter(
        ParameterSetName = 'ByUrl',
        Mandatory,
        Position = 0,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        HelpMessage = 'One or more URLs to test (e.g. https://contoso.com)'
    )]
    [ValidateNotNullOrEmpty()]
    [string[]] $Url,

    [Parameter(
        ParameterSetName = 'ByFile',
        Mandatory,
        HelpMessage = 'Path to a CSV file with a Url column'
    )]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $InputFile,

    [Parameter()]
    [ValidateRange(1, 300)]
    [int] $TimeoutSeconds = 30,

    [Parameter()]
    [switch] $Detailed
)

begin {
    $ErrorActionPreference = 'Stop'

    # Accumulate URLs when reading from a file so we can process them in process{}
    $urlQueue = [System.Collections.Generic.List[string]]::new()

    if ($PSCmdlet.ParameterSetName -eq 'ByFile') {
        Write-Verbose "Reading URLs from: $InputFile"
        try {
            $csv = Import-Csv -Path $InputFile
            if (-not ($csv | Get-Member -Name 'Url' -ErrorAction SilentlyContinue)) {
                throw "CSV file '$InputFile' must contain a column named 'Url'."
            }
            foreach ($row in $csv) {
                if ($row.Url) { $urlQueue.Add($row.Url.Trim()) }
            }
        }
        catch {
            Write-Error "Failed to read input file: $_"
            return
        }
        Write-Verbose "$($urlQueue.Count) URL(s) loaded from CSV."
    }

    # Helper: retrieve SSL certificate expiry for an HTTPS endpoint
    function Get-SslExpiry {
        [CmdletBinding()]
        param([string] $TargetUrl)

        try {
            $uri = [System.Uri]::new($TargetUrl)
            if ($uri.Scheme -ne 'https') { return $null }

            $tcpClient = [System.Net.Sockets.TcpClient]::new()
            $tcpClient.Connect($uri.Host, $uri.Port)

            $sslStream = [System.Net.Security.SslStream]::new(
                $tcpClient.GetStream(),
                $false,
                { $true }   # accept all certs — we only care about expiry date
            )
            $sslStream.AuthenticateAsClient($uri.Host)
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                $sslStream.RemoteCertificate
            )
            $expiry = $cert.NotAfter
            $sslStream.Dispose()
            $tcpClient.Dispose()
            return $expiry
        }
        catch {
            Write-Verbose "SSL check failed for '$TargetUrl': $_"
            return $null
        }
    }

    # Helper: test a single URL and return a result object
    function Test-SingleUrl {
        [CmdletBinding()]
        param([string] $TargetUrl)

        Write-Verbose "Testing: $TargetUrl"

        $result = [PSCustomObject]@{
            Url              = $TargetUrl
            StatusCode       = $null
            StatusDescription = $null
            ResponseTimeMs   = $null
            SslExpiryDate    = $null
            SslDaysRemaining = $null
            IsHealthy        = $false
            ErrorMessage     = $null
            Headers          = $null
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $iwrParams = @{
                Uri             = $TargetUrl
                Method          = 'GET'
                UseBasicParsing = $true
                TimeoutSec      = $TimeoutSeconds
                ErrorAction     = 'Stop'
            }

            # Suppress progress bar for Invoke-WebRequest — it skews timing
            $prevProgress = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            $response = Invoke-WebRequest @iwrParams

            $ProgressPreference = $prevProgress
            $stopwatch.Stop()

            $result.StatusCode        = [int]$response.StatusCode
            $result.StatusDescription = $response.StatusDescription
            $result.ResponseTimeMs    = $stopwatch.ElapsedMilliseconds
            $result.IsHealthy         = ($response.StatusCode -ge 200 -and $response.StatusCode -lt 400)

            if ($Detailed) {
                $headerTable = @{}
                foreach ($key in $response.Headers.Keys) {
                    $headerTable[$key] = $response.Headers[$key]
                }
                $result.Headers = $headerTable
            }
        }
        catch [System.Net.WebException] {
            $stopwatch.Stop()
            $result.ResponseTimeMs = $stopwatch.ElapsedMilliseconds

            # A WebException may still carry a response (e.g. HTTP 404, 500)
            if ($_.Exception.Response) {
                $httpResponse = [System.Net.HttpWebResponse]$_.Exception.Response
                $result.StatusCode        = [int]$httpResponse.StatusCode
                $result.StatusDescription = $httpResponse.StatusDescription
                $result.IsHealthy         = $false
            }
            $result.ErrorMessage = $_.Exception.Message
            Write-Verbose "WebException for '$TargetUrl': $($_.Exception.Message)"
        }
        catch {
            $stopwatch.Stop()
            $result.ResponseTimeMs = $stopwatch.ElapsedMilliseconds
            $result.ErrorMessage   = $_.Exception.Message
            $result.IsHealthy      = $false
            Write-Verbose "Error testing '$TargetUrl': $($_.Exception.Message)"
        }

        # SSL check is independent of the HTTP call — run it even if HTTP failed
        $result.SslExpiryDate = Get-SslExpiry -TargetUrl $TargetUrl
        if ($result.SslExpiryDate) {
            $result.SslDaysRemaining = ([int]($result.SslExpiryDate - (Get-Date)).TotalDays)
        }

        return $result
    }
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'ByUrl') {
        foreach ($u in $Url) {
            Test-SingleUrl -TargetUrl $u.Trim()
        }
    }
    else {
        # ByFile: urlQueue was populated in begin{}
        $total   = $urlQueue.Count
        $current = 0

        foreach ($u in $urlQueue) {
            $current++
            $pct = [int](($current / $total) * 100)
            Write-Progress -Activity 'Testing website health' `
                           -Status "[$current/$total] $u" `
                           -PercentComplete $pct

            Test-SingleUrl -TargetUrl $u
        }

        Write-Progress -Activity 'Testing website health' -Completed
    }
}

end {
    # Nothing to clean up — connections are disposed inside helpers
}
