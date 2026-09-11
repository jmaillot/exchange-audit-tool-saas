<#
.SYNOPSIS
  SharePoint tenant sharing policy for migration prep (admin layer): the full
  Get-SPOTenant surface that Microsoft Graph does NOT expose.

  The SaaS "Sharing Policy" section covers via Graph (one row): TenantName,
  SharingCapability, RequireAcceptingAccountMatchInvitedAccount,
  SharingDomainRestrictionMode, SharingAllowedDomainList,
  SharingBlockedDomainList. Everything below needs SharePoint Online
  Management Shell - which is why it lives here, not in the SaaS worker
  (delegated tokens cannot open the SPO admin connection).

.REQUIREMENTS
  Module Microsoft.Online.SharePoint.PowerShell, SharePoint administrator role.

.EXAMPLE
  .\Get-SPOTenantSharing.ps1 -AdminUrl https://TENANT-admin.sharepoint.com -OutDir .\out

.OUTPUTS (; -delimited UTF-8, same convention as the SaaS audit tool)
  Tenant-Sharing.csv (one row; multi-values joined with ','; empty = not
  exposed by this cmdlet version)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AdminUrl,          # e.g. https://contoso-admin.sharepoint.com
    [string]$OutDir = "."
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

Write-Host 'Connecting SharePoint Online Management Shell...'
Connect-SPOService -Url $AdminUrl

Write-Host 'Reading tenant settings (Get-SPOTenant)...'
$t = Get-SPOTenant

# Multi-valued properties arrive as string arrays - join them for the ';' CSV.
function Join-Multi($v) { ((@($v) | ForEach-Object { [string]$_ }) | Where-Object { $_ -ne '' }) -join ',' }

$row = [pscustomobject]@{
    TenantName                                    = $t.Title
    SharingCapability                             = $t.SharingCapability
    DefaultSharingLinkType                        = $t.DefaultSharingLinkType
    DefaultLinkPermission                         = $t.DefaultLinkPermission
    RequireAnonymousLinksExpireInDays             = $t.RequireAnonymousLinksExpireInDays
    CoreAnyoneSharingLinkMaxExpirationInDays      = $t.CoreAnyoneSharingLinkMaxExpirationInDays
    OneDriveAnyoneSharingLinkMaxExpirationInDays  = $t.OneDriveAnyoneSharingLinkMaxExpirationInDays
    CoreOrganizationSharingLinkMaxExpirationInDays = $t.CoreOrganizationSharingLinkMaxExpirationInDays
    OneDriveOrganizationSharingLinkMaxExpirationInDays = $t.OneDriveOrganizationSharingLinkMaxExpirationInDays
    SharingDomainRestrictionMode                  = $t.SharingDomainRestrictionMode
    SharingAllowedDomainList                      = Join-Multi $t.SharingAllowedDomainList
    SharingBlockedDomainList                      = Join-Multi $t.SharingBlockedDomainList
    RequireAcceptingAccountMatchInvitedAccount    = $t.RequireAcceptingAccountMatchInvitedAccount
    EmailAttestationRequired                      = $t.EmailAttestationRequired
    EmailAttestationReAuthDays                    = $t.EmailAttestationReAuthDays
    BccExternalSharingInvitations                 = $t.BccExternalSharingInvitations
    BccExternalSharingInvitationsList             = Join-Multi $t.BccExternalSharingInvitationsList
    OneDriveForGuestsEnabled                      = $t.OneDriveForGuestsEnabled
    ODBMembersCanShare                            = $t.ODBMembersCanShare
    ODBAccessRequests                             = $t.ODBAccessRequests
    ShowPeoplePickerSuggestionsForGuestUsers      = $t.ShowPeoplePickerSuggestionsForGuestUsers
}
@($row) | Export-Csv (Join-Path $OutDir 'Tenant-Sharing.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-Host 'Done: Tenant-Sharing.csv'
