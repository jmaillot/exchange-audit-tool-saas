using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "SharePoint" category (Microsoft Graph, NOT Exchange Online) for
    // tenant-to-tenant migration prep:
    //   sharepoint-sites    -> usage report (full site enumeration: id, URL,
    //     template, owner, file count, deleted flag) + GET /sites/{id} (live
    //     title/URL/dates) + /drive (live quota + owning group) + optional
    //     group owners / group sensitivity label + Teams-membership join.
    //     One row per site collection. OneDrive personal sites are skipped.
    //   sharepoint-subsites -> same enumeration, then recursive walk of
    //     /sites/{id}/sites. One row per web (root webs excluded: they are in
    //     Site Collections). Web-level config (template, language, unique
    //     permissions...) is NOT exposed by Graph - see sp-migration/.
    // The report is the only Graph-native full enumeration and needs
    // Reports.Read.All; without it these sections yield nothing (warning).
    // Reading non-group sites/drives needs Sites.Read.All.
    internal static class SectionsSharePoint
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildSitesSection());
            AuditRegistry.Register(BuildSubsitesSection());
            AuditRegistry.Register(BuildSharingSection());
        }

        private static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            return PsScriptHelpers.Collect(sel, groupKeys);
        }

        private const string GraphHeaders =
            "$headers = @{ Authorization = ('Bearer ' + $env:EAT_GRAPH_TOKEN); ConsistencyLevel = 'eventual'; Prefer = 'include-unknown-enum-members' }";

        private static string SizePrelude(string size)
        {
            // NOTE: plain concatenation, NOT string.Format: the PowerShell body
            // contains literal { } blocks that string.Format would try to parse.
            return "$size = '" + size + "'\n$maxSites = -1\nif ($size -eq '100') { $maxSites = 100 } elseif ($size -eq '1000') { $maxSites = 1000 }";
        }

        // Tenant usage report (D30): the only Graph-native full site
        // enumeration. Stale up to 48h. $repRows on success, empty + warning
        // otherwise (Report Site URL is often empty - join on Site Id).
        private const string FetchUsageReport =
            "$repRows = @()\n" +
            "try {\n" +
            "    $repRaw = Invoke-WebRequest -Uri \"https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')\" -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop\n" +
            "    $repTmp = [System.IO.Path]::GetTempFileName()\n" +
            "    if ($repRaw.Content -is [string] -and ([string]$repRaw.Content).Length -gt 0) {\n" +
            "        [System.IO.File]::WriteAllText($repTmp, [string]$repRaw.Content)\n" +
            "    } else {\n" +
            "        $rs = $repRaw.RawContentStream\n" +
            "        $rs.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null\n" +
            "        $fs = [System.IO.File]::Create($repTmp)\n" +
            "        try { $rs.CopyTo($fs) } finally { $fs.Dispose() }\n" +
            "    }\n" +
            "    Write-Host (\"Report download: HTTP {0}.\" -f [int]$repRaw.StatusCode)\n" +
            "    $repBytes = (Get-Item -LiteralPath $repTmp).Length\n" +
            "    Write-Host (\"Report file: {0} bytes.\" -f $repBytes)\n" +
            "    $repHead = ((Get-Content -LiteralPath $repTmp -TotalCount 2) -join '|')\n" +
            "    Write-Host (\"Report head: \" + $repHead.Substring(0, [Math]::Min(300, $repHead.Length)))\n" +
            "    $repRows = @(Import-Csv -LiteralPath $repTmp)\n" +
            "    Remove-Item -LiteralPath $repTmp -Force -ErrorAction SilentlyContinue\n" +
            "    Write-Host (\"Loaded SharePoint usage report ({0} rows).\" -f $repRows.Count)\n" +
            "} catch { Write-Host (\"WARNING: SharePoint usage report failed (consent for Reports.Read.All?) - no sites to process: \" + $_.Exception.Message) }\n";

        // Column-name-tolerant accessor: report headers sometimes arrive with
        // BOM/whitespace/case variations. Direct $rr.'Site Id' lookups would
        // then silently miss and rows would be skipped without a warning.
        private const string ColHelper =
            "function Get-Col($o, $name) {\n" +
            "    foreach ($p in @($o.PSObject.Properties)) { if ([string]$p.Name -eq $name) { return ([string]$p.Value).Trim() } }\n" +
            "    foreach ($p in @($o.PSObject.Properties)) { if (([string]$p.Name).Trim().Trim([char]0xFEFF).ToLower() -eq ([string]$name).ToLower()) { return ([string]$p.Value).Trim() } }\n" +
            "    return ''\n" +
            "}\n";

        // OneDrive personal sites never appear here (groups have none).

        // ============================================================ 1. SITE COLLECTIONS
        private static AuditSection BuildSitesSection()
        {
            var section = new AuditSection(
                "sharepoint-sites",
                "Site Collections",
                "SharePoint site collections export",
                "Audit group-connected site collections via Graph (one row per site) + usage-report enrichment where matchable. Non-group sites (communication/classic) need sp-migration/Get-SPOSiteInventory.ps1.",
                "site",
                AuditScope.Graph);
            section.Category = "SharePoint";
            section.Product = "SharePoint";
            section.DefaultFileName = "SharePointSites.csv";

            var site = new AuditOptionGroup("site", "Site columns", GroupMode.MultiCheck); site.Columns = 2;
            site.AddProp("SiteId", true);
            site.AddProp("Url", true);
            site.AddProp("Title", true);
            site.AddProp("Template", true);
            site.AddProp("GroupId", true);
            site.AddProp("PrimaryOwner", true);
            site.AddProp("Owners", true);
            site.AddProp("IsTeamsConnected", true);
            site.AddProp("SensitivityLabelId", true);
            site.AddProp("StorageUsedMB", true);
            site.AddProp("StorageQuotaMB", true);
            site.AddProp("CreatedDate", true);
            site.AddProp("LastContentModifiedDate", true);
            site.AddProp("Status", true);
            site.AddProp("MigrationWave", true);
            site.AddProp("MigrationDecision", true);
            section.AddGroup(site);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "site");
                if (chosen.Count == 0) chosen.Add("Url");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                bool wantOwners = chosen.Contains("Owners");
                // Template/Status/PrimaryOwner come from the usage report (matched
                // by site id/URL where the report carries usable keys).
                bool wantReport = chosen.Contains("Template") || chosen.Contains("Status")
                    || chosen.Contains("PrimaryOwner");
                bool wantDrive = chosen.Contains("StorageUsedMB") || chosen.Contains("StorageQuotaMB");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Microsoft 365 groups (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine("$groups = @()");
                sb.AppendLine("$guri = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mail,resourceProvisioningOptions,assignedLabels&$top=999'");
                sb.AppendLine("while ($guri -and ($maxSites -lt 0 -or $groups.Count -lt $maxSites)) {");
                sb.AppendLine("    $gr = Invoke-RestMethod -Uri $guri -Headers $headers -Method Get -ErrorAction Stop");
                sb.AppendLine("    $groups += @($gr.value)");
                sb.AppendLine("    $guri = $gr.'@odata.nextLink'");
                sb.AppendLine("}");
                sb.AppendLine("if ($maxSites -ge 0 -and $groups.Count -gt $maxSites) { $groups = $groups[0..($maxSites - 1)] }");
                sb.AppendLine("Write-Host (\"Retrieved {0} group(s).\" -f $groups.Count)");
                if (wantReport)
                {
                    sb.AppendLine(FetchUsageReport);
                    sb.AppendLine(ColHelper);
                    sb.AppendLine("$repById = @{}; $repByUrl = @{}");
                    sb.AppendLine("foreach ($rr in $repRows) {");
                    sb.AppendLine("    $ik = (Get-Col $rr 'Site Id').ToLower(); if ($ik -and $ik -ne '00000000-0000-0000-0000-000000000000' -and -not $repById.ContainsKey($ik)) { $repById[$ik] = $rr }");
                    sb.AppendLine("    $uk = (Get-Col $rr 'Site URL').TrimEnd('/').ToLower(); if ($uk -and -not $repByUrl.ContainsKey($uk)) { $repByUrl[$uk] = $rr }");
                    sb.AppendLine("}");
                    sb.AppendLine("Write-Host (\"Indexed {0} report row(s) for enrichment.\" -f $repById.Count)");
                }
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; if (($i % 20) -eq 0) { Write-Host (\"  ... {0}/{1}\" -f $i, $groups.Count) }");
                sb.AppendLine("    $gid = [string]$g.id; $gname = [string]$g.displayName");
                sb.AppendLine("    $su = ''; $ti = ''; $cd = ''; $lm = ''; $sumb = ''; $sqmb = ''; $owns = ''; $team = 'No'");
                sb.AppendLine("    $tpl = ''; $stat = 'Active'; $own = ''");
                sb.AppendLine("    $slid = ((@($g.assignedLabels) | ForEach-Object { [string]$_.labelId }) | Where-Object { $_ -ne '' }) -join ','");
                sb.AppendLine("    if (@($g.resourceProvisioningOptions) -contains 'Team') { $team = 'Yes' }");
                sb.AppendLine("    $st = $null");
                sb.AppendLine("    try { $st = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/sites/root?`$select=id,displayName,webUrl,lastModifiedDateTime\") -Headers $headers -Method Get -ErrorAction Stop } catch { if ($_.Exception.Message -notmatch '404') { Write-Host (\"  WARNING: site of group '{0}' failed: \" -f $gname + $_.Exception.Message) } }");
                sb.AppendLine("    if (-not $st) { continue }");
                sb.AppendLine("    $coll = ''");
                sb.AppendLine("    if ($st) {");
                sb.AppendLine("        if ($st.webUrl) { $su = [string]$st.webUrl }");
                sb.AppendLine("        $ti = [string]$st.displayName; $cd = [string]$st.createdDateTime; $lm = [string]$st.lastModifiedDateTime");
                sb.AppendLine("        $cp = ([string]$st.id).Split(','); if ($cp.Count -ge 3) { $coll = $cp[1] }");
                if (wantDrive)
                {
                    sb.AppendLine("        try {");
                    sb.AppendLine("            $dr = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/drive?`$select=quota\") -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("            if ($dr.quota) {");
                    sb.AppendLine("                try { if ($null -ne $dr.quota.used) { $sumb = [math]::Round([double]$dr.quota.used / 1MB, 2) } } catch { }");
                    sb.AppendLine("                try { if ($null -ne $dr.quota.total) { $sqmb = [math]::Round([double]$dr.quota.total / 1MB, 2) } } catch { }");
                    sb.AppendLine("            }");
                    sb.AppendLine("        } catch { Write-Host (\"  WARNING: drive of site '{0}' failed: \" -f $su + $_.Exception.Message) }");
                }
                if (wantReport)
                {
                    sb.AppendLine("        $rm = $null");
                    sb.AppendLine("        $ck = $coll.Trim().ToLower(); if ($ck -and $repById.ContainsKey($ck)) { $rm = $repById[$ck] }");
                    sb.AppendLine("        if (-not $rm) { $ukey = $su.Trim().TrimEnd('/').ToLower(); if ($ukey -and $repByUrl.ContainsKey($ukey)) { $rm = $repByUrl[$ukey] } }");
                    sb.AppendLine("        if ($rm) { $tpl = Get-Col $rm 'Root Web Template'; if ((Get-Col $rm 'Is Deleted') -eq 'True') { $stat = 'Deleted' }; $own = Get-Col $rm 'Owner Principal Name'; if (-not $own) { $own = Get-Col $rm 'Owner Display Name' } }");
                }
                sb.AppendLine("    }");
                if (wantOwners)
                {
                    sb.AppendLine("    try {");
                    sb.AppendLine("        $ol = @(); $ouri = \"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/owners?`$select=id,userPrincipalName,displayName\"");
                    sb.AppendLine("        while ($ouri) { $orr = Invoke-RestMethod -Uri $ouri -Headers $headers -Method Get -ErrorAction Stop; $ol += @($orr.value); $ouri = $orr.'@odata.nextLink' }");
                    sb.AppendLine("        $owns = (($ol | ForEach-Object { $u = [string]$_.userPrincipalName; if (-not $u) { $u = [string]$_.displayName }; $u }) | Where-Object { $_ -ne '' }) -join ','");
                    sb.AppendLine("    } catch { Write-Host (\"  WARNING: owners of group '{0}' failed: \" -f $gname + $_.Exception.Message) }");
                }
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        SiteId = $coll; Url = $su; Title = $ti; Template = $tpl; GroupId = $gid;");
                sb.AppendLine("        PrimaryOwner = $own; Owners = $owns; IsTeamsConnected = $team; SensitivityLabelId = $slid;");
                sb.AppendLine("        StorageUsedMB = $sumb; StorageQuotaMB = $sqmb;");
                sb.AppendLine("        CreatedDate = $cd; LastContentModifiedDate = $lm; Status = $stat;");
                sb.AppendLine("        MigrationWave = ''; MigrationDecision = ''");
                sb.AppendLine("    })");
                sb.AppendLine("}");
                sb.AppendLine("$rows = $rows | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 2. SUBSITES
        private static AuditSection BuildSubsitesSection()
        {
            var section = new AuditSection(
                "sharepoint-subsites",
                "Subsites",
                "SharePoint subsites export",
                "Audit webs below group-connected sites via Graph: one row per web (root webs excluded). Other sites + web-level config (template, language, unique permissions) need sp-migration/.",
                "site",
                AuditScope.Graph);
            section.Category = "SharePoint";
            section.Product = "SharePoint";
            section.DefaultFileName = "SharePointSubsites.csv";

            var sub = new AuditOptionGroup("subsite", "Subsite columns", GroupMode.MultiCheck); sub.Columns = 2;
            sub.AddProp("SiteId", true);
            sub.AddProp("WebId", true);
            sub.AddProp("ParentWebId", true);
            sub.AddProp("Url", true);
            sub.AddProp("Title", true);
            sub.AddProp("Created", true);
            sub.AddProp("LastItemModifiedDate", true);
            section.AddGroup(sub);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "subsite");
                if (chosen.Count == 0) chosen.Add("Url");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying SharePoint subsites (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine("$groups = @()");
                sb.AppendLine("$guri = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName&$top=999'");
                sb.AppendLine("while ($guri -and ($maxSites -lt 0 -or $groups.Count -lt $maxSites)) {");
                sb.AppendLine("    $gr = Invoke-RestMethod -Uri $guri -Headers $headers -Method Get -ErrorAction Stop");
                sb.AppendLine("    $groups += @($gr.value)");
                sb.AppendLine("    $guri = $gr.'@odata.nextLink'");
                sb.AppendLine("}");
                sb.AppendLine("if ($maxSites -ge 0 -and $groups.Count -gt $maxSites) { $groups = $groups[0..($maxSites - 1)] }");
                sb.AppendLine("Write-Host (\"Retrieved {0} group(s).\" -f $groups.Count)");
                sb.AppendLine("function Get-Kids($sid) {");
                sb.AppendLine("    $k = @()");
                sb.AppendLine("    try {");
                sb.AppendLine("        $u = \"https://graph.microsoft.com/v1.0/sites/\" + $sid + \"/sites\"");
                sb.AppendLine("        while ($u) { $r = Invoke-RestMethod -Uri $u -Headers $headers -Method Get -ErrorAction Stop; $k += @($r.value); $u = $r.'@odata.nextLink' }");
                sb.AppendLine("    } catch { }");
                sb.AppendLine("    return $k");
                sb.AppendLine("}");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; if (($i % 20) -eq 0) { Write-Host (\"  ... {0}/{1}\" -f $i, $groups.Count) }");
                sb.AppendLine("    $gid = [string]$g.id; $gname = [string]$g.displayName");
                sb.AppendLine("    $rt = $null");
                sb.AppendLine("    try { $rt = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/sites/root?`$select=id\") -Headers $headers -Method Get -ErrorAction Stop } catch { if ($_.Exception.Message -notmatch '404') { Write-Host (\"  WARNING: site of group '{0}' failed: \" -f $gname + $_.Exception.Message) } }");
                sb.AppendLine("    if (-not $rt) { continue }");
                sb.AppendLine("    $rp = ([string]$rt.id).Split(',')");
                sb.AppendLine("    $rweb = ''; $coll = ''; if ($rp.Count -ge 3) { $rweb = $rp[2]; $coll = $rp[1] }");
                sb.AppendLine("    $stack = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("    $stack.Add([pscustomobject]@{ Sid = [string]$rt.id; Parent = $rweb })");
                sb.AppendLine("    while ($stack.Count -gt 0) {");
                sb.AppendLine("        $tp = $stack[$stack.Count - 1]; $stack.RemoveAt($stack.Count - 1)");
                sb.AppendLine("        foreach ($kd in @(Get-Kids $tp.Sid)) {");
                sb.AppendLine("            $kcid = [string]$kd.id; $kp = $kcid.Split(',')");
                sb.AppendLine("            $kw = ''; if ($kp.Count -ge 3) { $kw = $kp[2] }");
                sb.AppendLine("            $rows.Add([pscustomobject]@{");
                sb.AppendLine("                SiteId = $coll; WebId = $kw; ParentWebId = $tp.Parent;");
                sb.AppendLine("                Url = $kd.webUrl; Title = $kd.displayName;");
                sb.AppendLine("                Created = $kd.createdDateTime; LastItemModifiedDate = $kd.lastModifiedDateTime");
                sb.AppendLine("            })");
                sb.AppendLine("            if ($kw) { $stack.Add([pscustomobject]@{ Sid = $kcid; Parent = $kw }) }");
                sb.AppendLine("        }");
                sb.AppendLine("    }");
                sb.AppendLine("}");
                sb.AppendLine("$rows = $rows | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 3. SHARING POLICY
        // Tenant-level sharing settings via Graph (/admin/sharepoint/settings).
        // Only a subset is exposed there (sharing capability, domain lists,
        // accept-match flag); the full Get-SPOTenant surface (default link
        // types, expiration windows, attestation, BCC, OneDrive flags...)
        // needs SharePoint Online Management Shell - see sp-migration/
        // Get-SPOTenantSharing.ps1. One row per tenant. Reading requires the
        // SharePoint Administrator or Global Reader role (delegated).
        private static AuditSection BuildSharingSection()
        {
            var section = new AuditSection(
                "sharepoint-sharing",
                "Sharing Policy",
                "SharePoint sharing policy export",
                "Audit tenant sharing settings via Graph: one row. Needs the SharePoint Administrator or Global Reader role; remaining Get-SPOTenant settings are covered by sp-migration/Get-SPOTenantSharing.ps1.",
                "shield",
                AuditScope.Graph);
            section.Category = "SharePoint";
            section.Product = "SharePoint";
            section.DefaultFileName = "SharePointSharing.csv";

            var tenant = new AuditOptionGroup("tenant", "Tenant", GroupMode.MultiCheck); tenant.Columns = 1;
            tenant.AddProp("TenantName", true);
            section.AddGroup(tenant);

            var sharing = new AuditOptionGroup("sharing", "Sharing policy", GroupMode.MultiCheck); sharing.Columns = 2;
            sharing.Hint = "Graph names differ slightly from Get-SPOTenant (RequireAcceptingAccountMatchInvitedAccount = isRequireAcceptingUserToMatchInvitedUserEnabled). Full detail (link defaults, expirations, attestation...): run sp-migration/Get-SPOTenantSharing.ps1 on an admin workstation.";
            sharing.AddProp("SharingCapability", true);
            sharing.AddProp("RequireAcceptingAccountMatchInvitedAccount", true);
            sharing.AddProp("SharingDomainRestrictionMode", true);
            sharing.AddProp("SharingAllowedDomainList", true);
            sharing.AddProp("SharingBlockedDomainList", true);
            section.AddGroup(sharing);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "tenant", "sharing");
                if (chosen.Count == 0) chosen.Add("TenantName");
                string selectList = string.Join(", ", chosen.ToArray());

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying SharePoint tenant settings (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine("$tname = ''");
                sb.AppendLine("try { $org = Invoke-RestMethod -Uri \"https://graph.microsoft.com/v1.0/organization?`$select=displayName\" -Headers $headers -Method Get -ErrorAction Stop; $tname = [string](@($org.value)[0].displayName) } catch { Write-Host (\"WARNING: organization failed: \" + $_.Exception.Message) }");
                sb.AppendLine("$ss = $null");
                sb.AppendLine("try {");
                sb.AppendLine("    $sresp = Invoke-RestMethod -Uri \"https://graph.microsoft.com/v1.0/admin/sharepoint/settings\" -Headers $headers -Method Get -ErrorAction Stop");
                sb.AppendLine("    if ($sresp.value) { $ss = $sresp.value } else { $ss = $sresp }");
                sb.AppendLine("    Write-Host 'Retrieved SharePoint tenant settings.'");
                sb.AppendLine("} catch { Write-Host (\"WARNING: SharePoint tenant settings failed (consent for SharePointTenantSettings.Read.All? SharePoint Administrator or Global Reader role?): \" + $_.Exception.Message) }");
                sb.AppendLine("$sc = ''; $ra = ''; $drm = ''; $al = ''; $bl = ''");
                sb.AppendLine("if ($ss) {");
                sb.AppendLine("    $sc = [string]$ss.sharingCapability");
                sb.AppendLine("    if ($null -ne $ss.isRequireAcceptingUserToMatchInvitedUserEnabled) { if ($ss.isRequireAcceptingUserToMatchInvitedUserEnabled) { $ra = 'Yes' } else { $ra = 'No' } }");
                sb.AppendLine("    $drm = [string]$ss.sharingDomainRestrictionMode");
                sb.AppendLine("    $al = ((@($ss.sharingAllowedDomainList) | ForEach-Object { [string]$_ }) -join ',')");
                sb.AppendLine("    $bl = ((@($ss.sharingBlockedDomainList) | ForEach-Object { [string]$_ }) -join ',')");
                sb.AppendLine("}");
                sb.AppendLine("$rows = @([pscustomobject]@{");
                sb.AppendLine("    TenantName = $tname; SharingCapability = $sc;");
                sb.AppendLine("    RequireAcceptingAccountMatchInvitedAccount = $ra; SharingDomainRestrictionMode = $drm;");
                sb.AppendLine("    SharingAllowedDomainList = $al; SharingBlockedDomainList = $bl");
                sb.AppendLine("})");
                sb.AppendLine("$rows = $rows | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }
    }
}
