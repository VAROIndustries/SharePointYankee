<#
.SYNOPSIS
    Creates a self-signed certificate in the local certificate store and exports
    the public key (.cer) and, optionally, the private key (.pfx).

.DESCRIPTION
    A generic wrapper around New-SelfSignedCertificate, Export-Certificate and
    Export-PfxCertificate. It does these things:
      * Requires an export folder that already exists.
      * Refuses to export into a folder inside a Git working tree, unless you
        pass -AllowRepositoryPath.
      * Refuses to overwrite existing .cer/.pfx files.
      * Asks for the PFX password twice (unless -PfxPassword is supplied) and
        rejects an empty or mismatched password.
      * Can remove the certificate from the store after export
        (-RemoveFromStore).

    Windows only: the PKI module and Cert: drive are not available on
    Linux or macOS.

.PARAMETER Name
    Base name. Used for the subject (CN=<Name>) and for the export file names
    (<Name>.cer / <Name>.pfx). Letters, digits, '.', '_' and '-' only.

.PARAMETER ExportFolder
    An existing folder for the exported files. Keep it outside source control.

.PARAMETER Subject
    Full subject. Defaults to 'CN=<Name>'.

.PARAMETER FriendlyName
    Friendly name shown in certmgr.msc. Defaults to <Name>.

.PARAMETER DnsName
    Optional DNS names for the Subject Alternative Name extension.

.PARAMETER Type
    Certificate type passed to New-SelfSignedCertificate.
    The default, SSLServerAuthentication, is also the cmdlet's own default.

.PARAMETER StoreLocation
    CurrentUser (default) or LocalMachine. LocalMachine needs an elevated session.
    New-SelfSignedCertificate only supports the 'My' (Personal) store.

.PARAMETER KeyLength
    RSA key length: 2048, 3072 or 4096. The default is 2048.

.PARAMETER HashAlgorithm
    SHA256 (default), SHA384 or SHA512.

.PARAMETER KeySpec
    Signature (default), KeyExchange or None. Signature/KeyExchange make the
    key a legacy CSP key; None makes it a CNG key.

.PARAMETER ValidityMonths
    How long the certificate is valid, in months (1-60). The default is 6.

.PARAMETER PfxPassword
    SecureString password for the PFX. If you leave it out, the script prompts
    for it twice.

.PARAMETER PfxEncryption
    AES256_SHA256 (default) or TripleDES_SHA1. Use TripleDES_SHA1 only when the
    target system cannot import AES-protected PFX files
    (versions older than Windows 10 1709 or Windows Server 2019).

.PARAMETER NoPfx
    Exports only the public .cer file. The private key stays in the store and
    is marked NonExportable.

.PARAMETER RemoveFromStore
    Deletes the certificate (and its private key) from the store after a
    successful export. Cannot be used with -NoPfx.

.PARAMETER AllowRepositoryPath
    Lets you export into a folder inside a Git working tree.

.EXAMPLE
    .\New-LocalCertificate.ps1 -Name 'Contoso-Flows-dev-HTTP' -ExportFolder 'D:\SecureExports' `
        -FriendlyName 'Contoso Flows dev HTTP authentication'

.EXAMPLE
    .\New-LocalCertificate.ps1 -Name 'Contoso-App-prod' -ExportFolder 'D:\SecureExports' `
        -ValidityMonths 12 -KeyLength 4096 -RemoveFromStore

.EXAMPLE
    .\New-LocalCertificate.ps1 -Name 'Contoso-PublicOnly' -ExportFolder 'D:\SecureExports' -NoPfx

.NOTES
    Author  : Geoff Varosky
    Module  : PKI (built into Windows)
    Version : 1.0.0
    Requires: Windows PowerShell 5.1 or PowerShell 7+ on Windows
    GitHub  : https://github.com/VAROIndustries/SharePointYankee
    References:
      https://learn.microsoft.com/powershell/module/pki/new-selfsignedcertificate
      https://learn.microsoft.com/powershell/module/pki/export-certificate
      https://learn.microsoft.com/powershell/module/pki/export-pfxcertificate
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [ValidateLength(1, 64)]
    [string] $Name,

    [Parameter(Mandatory)]
    [string] $ExportFolder,

    [string] $Subject,

    [string] $FriendlyName,

    [string[]] $DnsName,

    [ValidateSet('SSLServerAuthentication', 'CodeSigningCert', 'DocumentEncryptionCert', 'Custom')]
    [string] $Type = 'SSLServerAuthentication',

    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string] $StoreLocation = 'CurrentUser',

    [ValidateSet(2048, 3072, 4096)]
    [int] $KeyLength = 2048,

    [ValidateSet('SHA256', 'SHA384', 'SHA512')]
    [string] $HashAlgorithm = 'SHA256',

    [ValidateSet('Signature', 'KeyExchange', 'None')]
    [string] $KeySpec = 'Signature',

    [ValidateRange(1, 60)]
    [int] $ValidityMonths = 6,

    [securestring] $PfxPassword,

    [ValidateSet('AES256_SHA256', 'TripleDES_SHA1')]
    [string] $PfxEncryption = 'AES256_SHA256',

    [switch] $NoPfx,

    [switch] $RemoveFromStore,

    [switch] $AllowRepositoryPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

#region Pre-flight checks
if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) {
    throw 'This script requires Windows (PKI module and Cert: drive).'
}

if ($NoPfx -and $RemoveFromStore) {
    throw '-RemoveFromStore with -NoPfx would destroy the only copy of the private key.'
}

if ($StoreLocation -eq 'LocalMachine') {
    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'StoreLocation LocalMachine requires an elevated PowerShell session.'
    }
}

if (-not (Test-Path -LiteralPath $ExportFolder -PathType Container)) {
    throw "The export folder must already exist: $ExportFolder"
}
$ExportFolder = (Resolve-Path -LiteralPath $ExportFolder).ProviderPath

if (-not $AllowRepositoryPath) {
    $probe = Get-Item -LiteralPath $ExportFolder
    while ($probe) {
        if (Test-Path -LiteralPath (Join-Path $probe.FullName '.git')) {
            throw "The export folder is inside a Git working tree ($($probe.FullName)). Choose a folder outside the repository, or pass -AllowRepositoryPath."
        }
        $probe = $probe.Parent
    }
}

$cerPath = Join-Path $ExportFolder "$Name.cer"
$pfxPath = Join-Path $ExportFolder "$Name.pfx"

$existing = @($cerPath, $pfxPath) | Where-Object { Test-Path -LiteralPath $_ }
if ($existing) {
    throw "Export file(s) already exist. Stop and review before proceeding: $($existing -join ', ')"
}

if (-not $Subject) { $Subject = "CN=$Name" }
if (-not $FriendlyName) { $FriendlyName = $Name }
#endregion

#region PFX password
if (-not $NoPfx -and -not $PfxPassword) {
    $PfxPassword = Read-Host 'Strong PFX password (store in your approved password manager)' -AsSecureString
    $confirm = Read-Host 'Confirm PFX password' -AsSecureString

    $p1 = [Net.NetworkCredential]::new('', $PfxPassword).Password
    $p2 = [Net.NetworkCredential]::new('', $confirm).Password
    try {
        if ([string]::IsNullOrEmpty($p1)) { throw 'The PFX password cannot be empty.' }
        if ($p1 -cne $p2) { throw 'The passwords do not match.' }
    }
    finally {
        $p1 = $null; $p2 = $null
    }
}
elseif (-not $NoPfx -and $PfxPassword.Length -eq 0) {
    throw 'The PFX password cannot be empty.'
}
#endregion

#region Create certificate
$certParams = @{
    Subject           = $Subject
    FriendlyName      = $FriendlyName
    CertStoreLocation = "Cert:\$StoreLocation\My"
    Type              = $Type
    KeyAlgorithm      = 'RSA'
    KeyLength         = $KeyLength
    HashAlgorithm     = $HashAlgorithm
    KeySpec           = $KeySpec
    KeyExportPolicy   = $(if ($NoPfx) { 'NonExportable' } else { 'Exportable' })
    NotAfter          = (Get-Date).AddMonths($ValidityMonths)
}
if ($DnsName) { $certParams.DnsName = $DnsName }

if (-not $PSCmdlet.ShouldProcess("$Subject in Cert:\$StoreLocation\My", 'Create self-signed certificate')) {
    return
}

$cert = New-SelfSignedCertificate @certParams
#endregion

#region Export
try {
    Export-Certificate -Cert $cert -FilePath $cerPath -Type CERT | Out-Null

    if (-not $NoPfx) {
        $pfxParams = @{
            Cert                  = $cert
            FilePath              = $pfxPath
            Password              = $PfxPassword
            CryptoAlgorithmOption = $PfxEncryption
            ChainOption           = 'EndEntityCertOnly'
        }
        Export-PfxCertificate @pfxParams | Out-Null
    }
}
catch {
    # If the export fails, remove the half-finished certificate and files
    # so a re-run does not leave orphaned keys behind.
    Write-Warning "Export failed; removing certificate $($cert.Thumbprint) from the store."
    Remove-Item -LiteralPath "Cert:\$StoreLocation\My\$($cert.Thumbprint)" -DeleteKey -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $cerPath, $pfxPath -ErrorAction SilentlyContinue
    throw
}

if ($RemoveFromStore) {
    Remove-Item -LiteralPath "Cert:\$StoreLocation\My\$($cert.Thumbprint)" -DeleteKey
}
#endregion

[pscustomobject]@{
    Subject      = $cert.Subject
    Thumbprint   = $cert.Thumbprint
    NotAfter     = $cert.NotAfter
    Store        = if ($RemoveFromStore) { '(removed)' } else { "Cert:\$StoreLocation\My" }
    CerPath      = $cerPath
    PfxPath      = if ($NoPfx) { $null } else { $pfxPath }
}
