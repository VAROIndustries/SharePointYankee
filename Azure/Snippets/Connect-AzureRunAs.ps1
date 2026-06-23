# Connect-AzureRunAs.ps1
# Authenticate to Azure from within an Azure Automation runbook using a Managed Identity.
#
# DEPRECATION NOTICE
# ------------------
# The original "Run As Account" (service principal + certificate) pattern was deprecated
# by Microsoft in September 2023 and removed from the Azure portal. The replacement is
# System-assigned or User-assigned Managed Identity on the Automation Account, which
# requires no credentials to manage and follows least-privilege RBAC.
#
# This snippet demonstrates both the modern (Managed Identity) and legacy (certificate
# service principal) patterns. Use the Managed Identity block for all new runbooks.
#
# Prerequisites (Managed Identity path):
#   1. Enable System-assigned Managed Identity on the Automation Account
#      (Automation Account > Identity > System assigned > On)
#   2. Grant the Managed Identity the required Azure RBAC role(s) on the target scope
#      (e.g., Reader on the subscription, Contributor on a resource group)
#   3. Az.Accounts module must be imported into the Automation Account (v2.x or later)
#
# Blog Post: https://sharepointyankee.com/azure-automation-managed-identity
# Author   : Geoff Varosky
# Website  : https://sharepointyankee.com
# Updated  : 2024 — migrated from AzureRM Run As Account to Az module + Managed Identity

#Requires -Modules @{ ModuleName = 'Az.Accounts'; ModuleVersion = '2.0.0' }

#region ConnectToAzure — Managed Identity (recommended)
# In Azure Automation, Connect-AzAccount -Identity resolves the system-assigned managed
# identity automatically. No secrets, no certificates, no rotation required.

try {
    Write-Output "Connecting to Azure using System-assigned Managed Identity..."
    Connect-AzAccount -Identity -ErrorAction Stop
    Write-Output "Connected. Subscription: $((Get-AzContext).Subscription.Name)"
} catch {
    Write-Error "Managed Identity authentication failed: $($_.Exception.Message)"
    throw
}
#endregion ConnectToAzure — Managed Identity

<#
#region ConnectToAzure — User-assigned Managed Identity (alternative)
# Use this block instead when you need a user-assigned identity so that multiple
# Automation Accounts can share the same identity and its RBAC assignments.
# Replace the ClientId value with the Client ID of your user-assigned managed identity.

$userAssignedClientId = '<user-assigned-managed-identity-client-id>'

try {
    Write-Output "Connecting via User-assigned Managed Identity ($userAssignedClientId)..."
    Connect-AzAccount -Identity -AccountId $userAssignedClientId -ErrorAction Stop
    Write-Output "Connected. Subscription: $((Get-AzContext).Subscription.Name)"
} catch {
    Write-Error "User-assigned Managed Identity authentication failed: $($_.Exception.Message)"
    throw
}
#endregion ConnectToAzure — User-assigned Managed Identity
#>

<#
#region ConnectToAzure — Certificate Service Principal (legacy fallback only)
# Use ONLY if you cannot enable Managed Identity (e.g., sovereign cloud constraints or
# a very old Automation Account). Store the thumbprint in an Automation variable or
# retrieve it from Key Vault — never hardcode it.
#
# Requires: Az.Accounts module (replaces the retired AzureRM module)

$tenantId       = Get-AutomationVariable -Name 'TenantId'
$applicationId  = Get-AutomationVariable -Name 'ServicePrincipalAppId'
$thumbprint     = Get-AutomationVariable -Name 'ServicePrincipalThumbprint'

try {
    Write-Output "Connecting via certificate service principal..."
    Connect-AzAccount -ServicePrincipal `
                      -TenantId $tenantId `
                      -ApplicationId $applicationId `
                      -CertificateThumbprint $thumbprint `
                      -ErrorAction Stop
    Write-Output "Connected. Subscription: $((Get-AzContext).Subscription.Name)"
} catch {
    Write-Error "Certificate service principal authentication failed: $($_.Exception.Message)"
    throw
}
#endregion ConnectToAzure — Certificate Service Principal (legacy fallback only)
#>
