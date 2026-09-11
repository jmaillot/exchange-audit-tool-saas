<#
.SYNOPSIS
  Per-site deep dive for SharePoint migration prep (PnP layer): webs, lists,
  fields, content types, modern pages + webparts, navigation, groups,
  unique-permissions detail, features, regional settings.

  Complements the SaaS Graph sections (which cover enumeration + storage +
  report data). Run AFTER filling MigrationWave/Decision in the SaaS CSV and
  filtering the wave to process.

.REQUIREMENTS
  Module PnP.PowerShell 2.x/3.x; SharePoint admin or site collection admin.
  Auth per site below is Interactive; for unattended runs switch to
  certificate auth: Connect-PnPOnline -Url $url -ClientId <id> -Tenant <t>
  -CertificatePath <pfx> (see PnP docs).

.EXAMPLE
  .\Get-PnPWebInventory.ps1 -SiteUrls 'https://contoso.sharepoint.com/sites/A' -OutDir .\out
  .\Get-PnPWebInventory.ps1 -InputCsv .\SharePointSites-20240101.csv -OutDir .\out

.OUTPUTS (; -delimited UTF-8, same convention as the SaaS audit tool)
  Webs.csv Lists.csv Fields.csv ContentTypes.csv Pages.csv WebParts.csv
  Navigation.csv Groups.csv Permissions.csv Features.csv
#>
[CmdletBinding()]
param(
    [string[]]$SiteUrls = @(),
    [string]$InputCsv = "",     # SaaS SharePointSites export (must contain Url)
    [string]$OutDir = "."
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
if ($InputCsv) {
    $SiteUrls += @((Import-Csv $InputCsv -Delimiter ';' | Where-Object { $_.Url } | ForEach-Object { $_.Url.Trim() }))
}
$SiteUrls = @($SiteUrls | Where-Object { $_ } | Select-Object -Unique)
if ($SiteUrls.Count -eq 0) { throw 'No site URLs: pass -SiteUrls and/or -InputCsv.' }

$webs = New-Object System.Collections.Generic.List[object]
$lists = New-Object System.Collections.Generic.List[object]
$fields = New-Object System.Collections.Generic.List[object]
$ctypes = New-Object System.Collections.Generic.List[object]
$pages = New-Object System.Collections.Generic.List[object]
$parts = New-Object System.Collections.Generic.List[object]
$navs = New-Object System.Collections.Generic.List[object]
$grps = New-Object System.Collections.Generic.List[object]
$perms = New-Object System.Collections.Generic.List[object]
$feats = New-Object System.Collections.Generic.List[object]

$i = 0
foreach ($url in $SiteUrls) {
    $i++
    Write-Host ("[{0}/{1}] {2}" -f $i, $SiteUrls.Count, $url)
    try {
        Connect-PnPOnline -Url $url -Interactive -ErrorAction Stop
    } catch { Write-Host ("  WARNING: connect failed: " + $_.Exception.Message); continue }

    try {
        # ---- Webs (root + recursive subsites) ----
        $allWebs = @()
        try {
            $root = Get-PnPWeb -Includes WebTemplate, Configuration, Language, HasUniqueRoleAssignments, Created, LastItemModifiedDate, Id, RegionalSettings, ParentWeb -ErrorAction Stop
            $allWebs += $root
            $allWebs += @(Get-PnPSubWeb -Recurse -Includes WebTemplate, Configuration, Language, HasUniqueRoleAssignments, Created, LastItemModifiedDate, Id, RegionalSettings, ParentWeb -ErrorAction Stop)
        } catch { Write-Host ("  WARNING: webs failed: " + $_.Exception.Message) }

        foreach ($w in $allWebs) {
            $pid = ''
            try { if ($w.ParentWeb -and $w.ParentWeb.Id) { $pid = [string]$w.ParentWeb.Id } } catch { }
            $loc = ''; $tz = ''
            try { if ($w.RegionalSettings) { $loc = [string]$w.RegionalSettings.LocaleId; $tz = [string]$w.RegionalSettings.TimeZone.Description } } catch { }
            $webs.Add([pscustomobject]@{
                SiteUrl = $url; WebId = [string]$w.Id; ParentWebId = $pid
                WebUrl = $w.Url; Title = $w.Title
                WebTemplate = $w.WebTemplate; Configuration = $w.Configuration
                Language = $w.Language; HasUniquePermissions = $w.HasUniqueRoleAssignments
                Created = $w.Created; LastItemModifiedDate = $w.LastItemModifiedDate
                LocaleId = $loc; TimeZone = $tz
            }) | Out-Null
            # Unique-permissions detail (unique webs only).
            if ($w.HasUniqueRoleAssignments) {
                try {
                    foreach ($ra in @($w.RoleAssignments)) {
                        $principal = ''
                        try { $principal = [string]$ra.Member.LoginName; if (-not $principal) { $principal = [string]$ra.Member.Title } } catch { }
                        $roles = ''
                        try { $roles = ((@($ra.RoleDefinitionBindings) | ForEach-Object { [string]$_.Name }) -join ',') } catch { }
                        $perms.Add([pscustomobject]@{ WebUrl = $w.Url; Principal = $principal; Roles = $roles }) | Out-Null
                    }
                } catch { }
            }
        }

        # ---- Lists / fields / content types ----
        $allLists = @()
        try { $allLists = @(Get-PnPList -ErrorAction Stop) } catch { Write-Host ("  WARNING: lists failed: " + $_.Exception.Message) }
        foreach ($l in $allLists) {
            $lists.Add([pscustomobject]@{
                WebUrl = $url; List = $l.Title; TemplateId = $l.BaseTemplate
                ItemCount = $l.ItemCount; Hidden = $l.Hidden; Versioning = $l.EnableVersioning
            }) | Out-Null
            try {
                foreach ($f in @(Get-PnPField -List $l -ErrorAction Stop)) {
                    $fields.Add([pscustomobject]@{
                        WebUrl = $url; List = $l.Title; Field = $f.InternalName
                        Type = $f.TypeAsString; Hidden = $f.Hidden; Required = $f.Required; Group = $f.Group
                    }) | Out-Null
                }
            } catch { }
            try {
                foreach ($c in @(Get-PnPContentType -List $l -ErrorAction Stop)) {
                    $cid = ''
                    try { $cid = [string]$c.Id.StringValue } catch { try { $cid = [string]$c.Id } catch { } }
                    $ctypes.Add([pscustomobject]@{
                        WebUrl = $url; List = $l.Title; ContentType = $c.Name; Id = $cid; Group = $c.Group
                    }) | Out-Null
                }
            } catch { }
        }

        # ---- Modern pages + webparts ----
        $allPages = @()
        try { $allPages = @(Get-PnPClientSidePage -ErrorAction Stop) } catch { }
        foreach ($p in $allPages) {
            $pages.Add([pscustomobject]@{
                WebUrl = $url; Page = $p.Name; Title = $p.Title
                LayoutType = $p.LayoutType; PromoteAs = $p.PromoteAs
            }) | Out-Null
            try {
                foreach ($c in @(Get-PnPPageComponent -Page $p.Name -ErrorAction Stop)) {
                    $parts.Add([pscustomobject]@{
                        WebUrl = $url; Page = $p.Name
                        InstanceId = $c.InstanceId; Name = $c.Name
                    }) | Out-Null
                }
            } catch { }
        }

        # ---- Navigation ----
        foreach ($locName in @('QuickLaunch', 'TopNavigationBar')) {
            try {
                foreach ($n in @(Get-PnPNavigationNode -Location $locName -ErrorAction Stop)) {
                    $navs.Add([pscustomobject]@{
                        WebUrl = $url; Location = $locName; Title = $n.Title; Url = $n.Url
                    }) | Out-Null
                }
            } catch { }
        }

        # ---- SharePoint groups + members ----
        try {
            foreach ($gg in @(Get-PnPGroup -ErrorAction Stop)) {
                try {
                    foreach ($m in @(Get-PnPGroupMember -Group $gg.LoginName -ErrorAction Stop)) {
                        $grps.Add([pscustomobject]@{
                            WebUrl = $url; Group = $gg.Title
                            Login = $m.LoginName; Title = $m.Title; PrincipalType = $m.PrincipalType
                        }) | Out-Null
                    }
                } catch { }
            }
        } catch { Write-Host ("  WARNING: groups failed: " + $_.Exception.Message) }

        # ---- Features ----
        foreach ($scope in @('Web', 'Site')) {
            try {
                foreach ($f in @(Get-PnPFeature -Scope $scope -ErrorAction Stop)) {
                    $feats.Add([pscustomobject]@{
                        WebUrl = $url; Scope = $scope
                        Name = $f.DisplayName; DefinitionId = $f.DefinitionId
                    }) | Out-Null
                }
            } catch { }
        }
    } catch { Write-Host ("  WARNING: site failed: " + $_.Exception.Message) }
}

$webs | Export-Csv (Join-Path $OutDir 'Webs.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$lists | Export-Csv (Join-Path $OutDir 'Lists.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$fields | Export-Csv (Join-Path $OutDir 'Fields.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$ctypes | Export-Csv (Join-Path $OutDir 'ContentTypes.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$pages | Export-Csv (Join-Path $OutDir 'Pages.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$parts | Export-Csv (Join-Path $OutDir 'WebParts.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$navs | Export-Csv (Join-Path $OutDir 'Navigation.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$grps | Export-Csv (Join-Path $OutDir 'Groups.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$perms | Export-Csv (Join-Path $OutDir 'Permissions.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$feats | Export-Csv (Join-Path $OutDir 'Features.csv') -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-Host 'Done.'
