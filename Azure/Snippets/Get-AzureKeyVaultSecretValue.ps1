# Get-AzureKeyVaultSecretValue.ps1
# Retrieve the plain-text value of an Azure Key Vault secret using the Az module.
#
# DEPRECATION NOTICE
# ------------------
# The original script used Get-AzureKeyVaultSecret (AzureRM module) and the
# .SecretValueText property, which was removed in the Az module. The modern equivalent
# is Get-AzKeyVaultSecret (Az.KeyVault) with the -AsPlainText switch (Az.KeyVault 3.3+).
# For older Az.KeyVault versions, convert the SecureString manually — see the fallback
# block below.
#
# Usage:
#   .\Get-AzureKeyVaultSecretValue.ps1 -VaultName "myvault" -SecretName "mysecret"
#
# Blog Post: https://sharepointyankee.com/getting-the-plain-text-value-of-an-azure-key-vault-secret-with-powershell
# Author   : Geoff Varosky
# Website  : https://sharepointyankee.com
# Updated  : 2024 — migrated from AzureRM (Get-AzureKeyVaultSecret) to Az module (Get-AzKeyVaultSecret)

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Az.KeyVault'; ModuleVersion = '3.3.0' }

[CmdletBinding()]
param (
    [Parameter(Mandatory, HelpMessage = 'Name of the Key Vault (not the full URI).')]
    [ValidateNotNullOrEmpty()]
    [string] $VaultName,

    [Parameter(Mandatory, HelpMessage = 'Name of the secret to retrieve.')]
    [ValidateNotNullOrEmpty()]
    [string] $SecretName,

    [Parameter(HelpMessage = 'Return a SecureString instead of plain text. Recommended when passing the value directly to another cmdlet that accepts SecureString.')]
    [switch] $AsSecureString
)

# Ensure the caller is authenticated before attempting Key Vault access.
# In interactive sessions, run Connect-AzAccount first.
# In Azure Automation runbooks, use Connect-AzAccount -Identity (see Connect-AzureRunAs.ps1).
if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
    throw "No active Azure session found. Run Connect-AzAccount (interactive) or Connect-AzAccount -Identity (Automation) before calling this script."
}

try {
    if ($AsSecureString) {
        # Return the SecureString directly — the secret value never touches plain-text memory
        $secret = Get-AzKeyVaultSecret -VaultName $VaultName -Name $SecretName -ErrorAction Stop
        if (-not $secret) {
            throw "Secret '$SecretName' was not found in vault '$VaultName'."
        }
        return $secret.SecretValue  # SecureString
    } else {
        # -AsPlainText requires Az.KeyVault 3.3.0 or later.
        # If you are on an older version, replace this with the SecureString conversion block below.
        $plainText = Get-AzKeyVaultSecret -VaultName $VaultName -Name $SecretName `
                        -AsPlainText -ErrorAction Stop
        if (-not $plainText) {
            throw "Secret '$SecretName' was not found in vault '$VaultName'."
        }
        return $plainText
    }
} catch {
    Write-Error "Failed to retrieve secret '$SecretName' from vault '$VaultName': $($_.Exception.Message)"
    throw
}

<#
# FALLBACK — Az.KeyVault older than 3.3.0 (no -AsPlainText parameter)
# Replace the try block above with this if you cannot upgrade the module.

try {
    $secret = Get-AzKeyVaultSecret -VaultName $VaultName -Name $SecretName -ErrorAction Stop
    if (-not $secret) {
        throw "Secret '$SecretName' was not found in vault '$VaultName'."
    }
    # Convert SecureString to plain text using the .NET DPAPI bridge
    $bstr      = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret.SecretValue)
    $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)   # Zero memory immediately
    return $plainText
} catch {
    Write-Error "Failed to retrieve secret '$SecretName' from vault '$VaultName': $($_.Exception.Message)"
    throw
}
#>
