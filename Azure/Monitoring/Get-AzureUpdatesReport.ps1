<#
.SYNOPSIS
    Fetches Azure Updates from the public RSS feed and produces a Markdown, HTML,
    or console report for the specified service categories.

.DESCRIPTION
    Queries the Azure Updates RSS feed (https://azurecomcdn.azureedge.net/en-us/updates/feed/)
    for one or more Azure service product IDs and filters results to a rolling window
    defined by -DaysBack. Outputs Markdown by default — suitable for documentation,
    wikis, and automated reporting pipelines. HTML and console output modes are also
    supported.

    The script includes a built-in hashtable of common Azure service product IDs.
    Pass friendly service names (e.g., "App Service", "Key Vault") and the script
    resolves them automatically. You may also pass raw product ID slugs directly.

    No authentication is required — the Azure Updates RSS feed is public.

.PARAMETER Services
    One or more Azure service names or product ID slugs to include in the report.
    Friendly names are resolved from the built-in lookup table. Run the script
    with -ListServices to see all available built-in mappings.
    Accepts pipeline input.

.PARAMETER DaysBack
    Number of days back from today to include in the report. Defaults to 30.

.PARAMETER OutputFormat
    Report format. Valid values: Markdown (default), HTML, Console.

.PARAMETER OutputPath
    Optional. Full path to save the report file. If omitted the report is written
    to the pipeline (stdout). Extension is appended automatically if not provided
    (.md for Markdown, .html for HTML).

.PARAMETER ListServices
    Switch. When specified, displays the built-in service name to product ID mapping
    and exits without querying the feed. Use this to discover valid -Services values.

.EXAMPLE
    Get-AzureUpdatesReport -Services "App Service","Key Vault","Azure SQL" -DaysBack 30

    Outputs a Markdown report for three services covering the last 30 days.

.EXAMPLE
    Get-AzureUpdatesReport -Services "Azure Functions","Storage" -OutputFormat HTML -OutputPath "C:\Reports\azure-updates.html"

    Generates an HTML report and saves it to disk.

.EXAMPLE
    Get-AzureUpdatesReport -ListServices

    Lists all built-in service mappings and exits.

.EXAMPLE
    "App Service","Azure SQL" | Get-AzureUpdatesReport -DaysBack 7 -OutputFormat Console

    Accepts service names from the pipeline and prints a console-formatted report.

.NOTES
    Author      : Geoff Varosky
    Version     : 1.0.0
    Requires    : PowerShell 5.1 or 7+; no external modules required.
    Feed URL    : https://azure.microsoft.com/en-us/updates/
    RSS API     : https://azurecomcdn.azureedge.net/en-us/updates/feed/?product=<slug>
    GitHub      : https://github.com/VAROIndustries/SharePointYankee
#>
#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Report')]
param (
    [Parameter(ParameterSetName = 'Report', Mandatory, ValueFromPipeline, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Services,

    [Parameter(ParameterSetName = 'Report')]
    [ValidateRange(1, 365)]
    [int]$DaysBack = 30,

    [Parameter(ParameterSetName = 'Report')]
    [ValidateSet('Markdown', 'HTML', 'Console')]
    [string]$OutputFormat = 'Markdown',

    [Parameter(ParameterSetName = 'Report')]
    [string]$OutputPath,

    [Parameter(ParameterSetName = 'List', Mandatory)]
    [switch]$ListServices
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    #region Built-in service product ID lookup table
    # Slugs are the ?product= query parameter values accepted by the Azure Updates RSS feed.
    # Source: https://azure.microsoft.com/en-us/updates/ (filter panel URL parameters)
    $ServiceMap = [ordered]@{
        'App Service'                = 'app-service'
        'App Service Plan'           = 'app-service-plan'
        'Application Gateway'        = 'application-gateway'
        'Application Insights'       = 'application-insights'
        'Azure Active Directory'     = 'azure-active-directory'
        'Azure API Management'       = 'api-management'
        'Azure Arc'                  = 'azure-arc'
        'Azure Automation'           = 'automation'
        'Azure Backup'               = 'backup'
        'Azure Bastion'              = 'azure-bastion'
        'Azure Container Apps'       = 'container-apps'
        'Azure Container Instances'  = 'container-instances'
        'Azure Container Registry'   = 'container-registry'
        'Azure Cosmos DB'            = 'cosmos-db'
        'Azure Data Factory'         = 'data-factory'
        'Azure Database for MySQL'   = 'mysql'
        'Azure Database for PostgreSQL' = 'postgresql'
        'Azure Defender'             = 'azure-defender'
        'Azure DevOps'               = 'azure-devops'
        'Azure DNS'                  = 'azure-dns'
        'Azure Firewall'             = 'azure-firewall'
        'Azure Functions'            = 'azure-functions'
        'Azure Kubernetes Service'   = 'kubernetes-service'
        'Azure Load Balancer'        = 'load-balancer'
        'Azure Logic Apps'           = 'logic-apps'
        'Azure Machine Learning'     = 'machine-learning-service'
        'Azure Monitor'              = 'monitor'
        'Azure OpenAI'               = 'azure-openai'
        'Azure Policy'               = 'azure-policy'
        'Azure Redis Cache'          = 'cache'
        'Azure Service Bus'          = 'service-bus-namespace'
        'Azure Site Recovery'        = 'site-recovery'
        'Azure SQL'                  = 'azure-sql'
        'Azure SQL Managed Instance' = 'azure-sql-managed-instance'
        'Azure Static Web Apps'      = 'static-web-apps'
        'Azure Storage'              = 'storage-account'
        'Azure Synapse Analytics'    = 'synapse-analytics'
        'Azure Virtual Desktop'      = 'virtual-desktop'
        'Azure VPN Gateway'          = 'vpn-gateway'
        'Event Grid'                 = 'event-grid'
        'Event Hubs'                 = 'event-hubs'
        'ExpressRoute'               = 'expressroute'
        'Key Vault'                  = 'key-vault'
        'Log Analytics'              = 'log-analytics-workspace'
        'Microsoft Defender for Cloud' = 'azure-defender'
        'Microsoft Entra'            = 'azure-active-directory'
        'Microsoft Purview'          = 'azure-purview'
        'Notification Hubs'          = 'notification-hubs'
        'SendGrid'                   = 'sendgrid'
        'Service Fabric'             = 'service-fabric'
        'Virtual Machine Scale Sets' = 'virtual-machine-scale-sets'
        'Virtual Machines'           = 'virtual-machines'
        'Virtual Network'            = 'virtual-network'
    }
    #endregion

    # Accumulate pipeline input across multiple process{} calls
    $allServices = [System.Collections.Generic.List[string]]::new()

    # Stores resolved (displayName, slug) pairs after pipeline is complete
    $resolvedServices = [System.Collections.Generic.List[hashtable]]::new()

    $cutoffDate = (Get-Date).AddDays(-$DaysBack).Date
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'List') { return }
    foreach ($svc in $Services) {
        $allServices.Add($svc)
    }
}

end {
    #region List mode
    if ($PSCmdlet.ParameterSetName -eq 'List') {
        Write-Host "`nBuilt-in Azure service mappings:`n" -ForegroundColor Cyan
        $ServiceMap.GetEnumerator() | ForEach-Object {
            Write-Host ("  {0,-42} -> {1}" -f $_.Key, $_.Value)
        }
        return
    }
    #endregion

    #region Resolve service names to product slugs
    foreach ($svc in $allServices) {
        $trimmed = $svc.Trim()
        if ($ServiceMap.Contains($trimmed)) {
            # Friendly name matched
            $resolvedServices.Add(@{ Name = $trimmed; Slug = $ServiceMap[$trimmed] })
        }
        elseif ($ServiceMap.Values -contains $trimmed) {
            # Raw slug passed — find the matching friendly name
            $friendlyName = ($ServiceMap.GetEnumerator() | Where-Object Value -eq $trimmed).Key
            $resolvedServices.Add(@{ Name = $friendlyName; Slug = $trimmed })
        }
        else {
            # Unknown service — treat input as a raw slug and use it as-is
            Write-Warning "Service '$trimmed' not found in built-in lookup table. Treating as a raw product slug."
            $resolvedServices.Add(@{ Name = $trimmed; Slug = $trimmed })
        }
    }

    if ($resolvedServices.Count -eq 0) {
        throw 'No services were resolved. Provide at least one valid service name or slug.'
    }
    #endregion

    #region Fetch RSS feed for each service
    $reportDate   = Get-Date -Format 'yyyy-MM-dd'
    $feedBaseUri  = 'https://azurecomcdn.azureedge.net/en-us/updates/feed/?product='
    $allItems     = [System.Collections.Generic.List[pscustomobject]]::new()

    $i = 0
    foreach ($svc in $resolvedServices) {
        $i++
        Write-Progress -Activity 'Fetching Azure Updates' `
            -Status "Querying: $($svc.Name)" `
            -PercentComplete (($i / $resolvedServices.Count) * 100)

        $uri = "$feedBaseUri$($svc.Slug)"
        Write-Verbose "GET $uri"

        try {
            $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -ErrorAction Stop
            [xml]$feed = $response.Content

            $filtered = $feed.rss.channel.item | Where-Object {
                $pubDate = [datetime]::Parse($_.pubDate)
                $pubDate.Date -ge $cutoffDate
            }

            foreach ($item in $filtered) {
                $allItems.Add([pscustomobject]@{
                    ServiceName = $svc.Name
                    Title       = $item.title
                    Link        = $item.link
                    PubDate     = [datetime]::Parse($item.pubDate)
                    Description = ($item.description -replace '<[^>]+>', '').Trim()
                })
            }
        }
        catch {
            Write-Warning "Failed to retrieve feed for '$($svc.Name)' (slug: $($svc.Slug)): $_"
        }
    }

    Write-Progress -Activity 'Fetching Azure Updates' -Completed
    #endregion

    #region Build output
    $output = switch ($OutputFormat) {

        'Markdown' {
            $sb = [System.Text.StringBuilder]::new()
            [void]$sb.AppendLine("# Azure Updates Report")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("**Generated:** $reportDate  ")
            [void]$sb.AppendLine("**Period:** Last $DaysBack days (since $($cutoffDate.ToString('yyyy-MM-dd')))  ")
            [void]$sb.AppendLine("**Services:** $($resolvedServices.Name -join ', ')  ")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("---")
            [void]$sb.AppendLine("")

            foreach ($svc in $resolvedServices) {
                $svcItems = $allItems | Where-Object ServiceName -eq $svc.Name | Sort-Object PubDate -Descending
                [void]$sb.AppendLine("## $($svc.Name)")
                [void]$sb.AppendLine("")

                if (-not $svcItems) {
                    [void]$sb.AppendLine("_No updates in the last $DaysBack days._")
                    [void]$sb.AppendLine("")
                }
                else {
                    foreach ($item in $svcItems) {
                        $dateStr = $item.PubDate.ToString('yyyy-MM-dd')
                        [void]$sb.AppendLine("### [$($item.Title)]($($item.Link))")
                        [void]$sb.AppendLine("")
                        [void]$sb.AppendLine("**Published:** $dateStr  ")
                        [void]$sb.AppendLine("")
                        if ($item.Description) {
                            [void]$sb.AppendLine($item.Description)
                            [void]$sb.AppendLine("")
                        }
                    }
                }
                [void]$sb.AppendLine("---")
                [void]$sb.AppendLine("")
            }

            $sb.ToString()
        }

        'HTML' {
            $sb = [System.Text.StringBuilder]::new()
            [void]$sb.AppendLine('<!DOCTYPE html>')
            [void]$sb.AppendLine('<html lang="en"><head><meta charset="UTF-8">')
            [void]$sb.AppendLine('<title>Azure Updates Report</title>')
            [void]$sb.AppendLine('<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:900px;margin:40px auto;color:#222}h1{color:#0078d4}h2{color:#005a9e;border-bottom:2px solid #0078d4;padding-bottom:4px}h3 a{color:#0078d4}p{line-height:1.6}.meta{color:#555;font-size:.9em}.no-updates{color:#888;font-style:italic}</style>')
            [void]$sb.AppendLine('</head><body>')
            [void]$sb.AppendLine('<h1>Azure Updates Report</h1>')
            [void]$sb.AppendLine("<p class='meta'><strong>Generated:</strong> $reportDate &nbsp;|&nbsp; <strong>Period:</strong> Last $DaysBack days (since $($cutoffDate.ToString('yyyy-MM-dd'))) &nbsp;|&nbsp; <strong>Services:</strong> $($resolvedServices.Name -join ', ')</p>")
            [void]$sb.AppendLine('<hr/>')

            foreach ($svc in $resolvedServices) {
                $svcItems = $allItems | Where-Object ServiceName -eq $svc.Name | Sort-Object PubDate -Descending
                [void]$sb.AppendLine("<h2>$($svc.Name)</h2>")

                if (-not $svcItems) {
                    [void]$sb.AppendLine("<p class='no-updates'>No updates in the last $DaysBack days.</p>")
                }
                else {
                    foreach ($item in $svcItems) {
                        $dateStr = $item.PubDate.ToString('yyyy-MM-dd')
                        [void]$sb.AppendLine("<h3><a href='$([System.Web.HttpUtility]::HtmlEncode($item.Link))'>$([System.Web.HttpUtility]::HtmlEncode($item.Title))</a></h3>")
                        [void]$sb.AppendLine("<p class='meta'><strong>Published:</strong> $dateStr</p>")
                        if ($item.Description) {
                            [void]$sb.AppendLine("<p>$([System.Web.HttpUtility]::HtmlEncode($item.Description))</p>")
                        }
                    }
                }
            }

            [void]$sb.AppendLine('</body></html>')
            $sb.ToString()
        }

        'Console' {
            $header = "Azure Updates Report | Last $DaysBack days | Generated: $reportDate"
            Write-Host "`n$header" -ForegroundColor Cyan
            Write-Host ('=' * $header.Length) -ForegroundColor Cyan

            foreach ($svc in $resolvedServices) {
                $svcItems = $allItems | Where-Object ServiceName -eq $svc.Name | Sort-Object PubDate -Descending
                Write-Host "`n$($svc.Name)" -ForegroundColor Yellow
                Write-Host ('-' * 60) -ForegroundColor DarkGray

                if (-not $svcItems) {
                    Write-Host "  No updates in the last $DaysBack days." -ForegroundColor DarkGray
                }
                else {
                    foreach ($item in $svcItems) {
                        $dateStr = $item.PubDate.ToString('yyyy-MM-dd')
                        Write-Host "  [$dateStr] $($item.Title)" -ForegroundColor White
                        Write-Host "           $($item.Link)" -ForegroundColor DarkCyan
                        Write-Host ""
                    }
                }
            }
            # Console mode writes directly; return empty string so -OutputPath still works if specified
            ''
        }
    }
    #endregion

    #region Write to file or pipeline
    if ($OutputFormat -ne 'Console' -or $OutputPath) {
        if ($PSBoundParameters.ContainsKey('OutputPath') -and $OutputPath) {
            # Append correct extension if the caller omitted it
            if ($OutputFormat -eq 'Markdown' -and $OutputPath -notmatch '\.(md|markdown)$') {
                $OutputPath = "$OutputPath.md"
            }
            elseif ($OutputFormat -eq 'HTML' -and $OutputPath -notmatch '\.html?$') {
                $OutputPath = "$OutputPath.html"
            }

            $output | Set-Content -Path $OutputPath -Encoding UTF8
            Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
        }
        else {
            Write-Output $output
        }
    }
    #endregion
}
