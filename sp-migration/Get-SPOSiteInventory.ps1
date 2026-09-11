<#
.SYNOPSIS
  SharePoint tenant inventory for tenant-to-tenant migration prep (admin layer).

  Complements the SaaS Graph sections (enumeration, live storage, usage-report
  data) with SharePoint-admin-only properties: LockState, SharingCapability,
  hub association, declared owner/quota, deleted sites.

.REQUIREMENTS
  Module Microsoft.Online.SharePoint.PowerShell, SharePoint administrator role.

.EXAMPLE
  .\Get-SPOSiteInventory.ps1 -AdminUrl https://TENANT-admin.sharepoint.com -OutDir .\out

.OUTPUTS (; -delimited UTF-8, same convention as the SaaS audit tool)
  Sites-Admin.csv, Sites-Deleted.csv
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

Write-Host 'Querying all site collections...'
$sites = @(Get-SPOSite -Limit All -IncludePersonalSite $false -Detailed)
Write-Host ("Retrieved {0} site(s)." -f $sites.Count)

$rows = foreach ($s in $sites) {
    # GroupId: property name varies by module version - try both, first wins.
    $gid = @($s.RelatedGroupId, $s.GroupId | Where-Object { $_ } | Select-Object -First 1) -join ''
    [pscustomobject]@{
        Title           = $s.Title
        Url             = $s.Url
        Template        = $s.Template
        Owner           = $s.Owner
        StorageUsedMB   = [math]::Round([double]$s.StorageUsageCurrent, 2)
        StorageQuotaMB  = $s.StorageQuota
        LockState       = $s.LockState
        SharingCapability = $s.SharingCapability
        IsHubSite       = $s.IsHubSite
        HubSiteId       = $s.HubSiteId
        GroupId         = $gid
        LastContentModifiedDate = $s.LastContentModifiedDate
    }
}
$rows | Export-Csv (Join-Path $OutDir 'Sites-Admin.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'

Write-Host 'Querying deleted sites...'
$del = @(Get-SPODeletedSite -Limit All | Select-Object Url, DaysRemaining, DeletionTime, SiteId)
$del | Export-Csv (Join-Path $OutDir 'Sites-Deleted.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'

Write-Host ("Done: {0} site(s), {1} deleted." -f $rows.Count, $del.Count)
