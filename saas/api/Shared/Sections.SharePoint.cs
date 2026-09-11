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
            "    $repRows = @(($repRaw.Content -split \"`n\" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) | ConvertFrom-Csv)\n" +
            "    Write-Host (\"Loaded SharePoint usage report ({0} rows).\" -f $repRows.Count)\n" +
            "} catch { Write-Host (\"WARNING: SharePoint usage report failed (consent for Reports.Read.All?) - no sites to process: \" + $_.Exception.Message) }\n";

        // Non-personal sites only (OneDrive excluded by URL or SPSPERS
        // template), then apply the size cap. Deleted sites are KEPT
        // (Status column marks them).
        private const string FilterWorkRows =
            "$work = @($repRows | Where-Object { (-not (([string]$_.'Site URL').ToLower() -match '/personal/')) -and (-not (([string]$_.'Root Web Template').Trim().ToLower().StartsWith('spspers'))) })\n" +
            "if ($maxSites -ge 0 -and $work.Count -gt $maxSites) { $work = $work[0..($maxSites - 1)] }\n" +
            "Write-Host (\"Processing {0} site(s).\" -f $work.Count)\n";

        // ============================================================ 1. SITE COLLECTIONS
        private static AuditSection BuildSitesSection()
        {
            var section = new AuditSection(
                "sharepoint-sites",
                "Site Collections",
                "SharePoint site collections export",
                "Audit every site collection via Graph (usage report + live lookups): one row per site. MigrationWave/Decision are empty working columns (uncheck Smart mode to keep them).",
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
                bool wantLabel = chosen.Contains("SensitivityLabelId");
                bool wantTeamsConn = chosen.Contains("IsTeamsConnected");
                bool wantDrive = chosen.Contains("StorageUsedMB") || chosen.Contains("StorageQuotaMB")
                    || chosen.Contains("GroupId") || wantOwners || wantLabel || wantTeamsConn;

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying SharePoint sites (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchUsageReport);
                sb.AppendLine(FilterWorkRows);
                if (wantTeamsConn)
                {
                    sb.AppendLine("$teamIds = @{}");
                    sb.AppendLine("try {");
                    sb.AppendLine("    $turi = 'https://graph.microsoft.com/v1.0/groups?$count=true&$filter=resourceProvisioningOptions/Any(x:x eq ''Team'')&$select=id&$top=999'");
                    sb.AppendLine("    while ($turi) { $tr = Invoke-RestMethod -Uri $turi -Headers $headers -Method Get -ErrorAction Stop; foreach ($tt in @($tr.value)) { $k = ([string]$tt.id).Trim().ToLower(); if ($k) { $teamIds[$k] = $true } }; $turi = $tr.'@odata.nextLink' }");
                    sb.AppendLine("    Write-Host (\"Loaded {0} team group(s).\" -f $teamIds.Count)");
                    sb.AppendLine("} catch { Write-Host (\"WARNING: team groups failed - IsTeamsConnected will be blank: \" + $_.Exception.Message) }");
                }
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($rr in $work) {");
                sb.AppendLine("    $i++; if (($i % 20) -eq 0) { Write-Host (\"  ... {0}/{1}\" -f $i, $work.Count) }");
                sb.AppendLine("    $sidRaw = ([string]$rr.'Site Id').Trim()");
                sb.AppendLine("    if (-not $sidRaw) { continue }");
                sb.AppendLine("    $st = $null");
                sb.AppendLine("    try { $st = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/sites/\" + $sidRaw) -Headers $headers -Method Get -ErrorAction Stop } catch { Write-Host (\"  WARNING: site {0} failed: \" -f $sidRaw + $_.Exception.Message) }");
                sb.AppendLine("    $su = ''; $ti = ''; $cd = ''; $lm = ''; $sumb = ''; $sqmb = ''; $gid = ''; $owns = ''; $slid = ''; $team = 'No'");
                sb.AppendLine("    $tpl = [string]$rr.'Root Web Template'");
                sb.AppendLine("    $stat = 'Active'; if ([string]$rr.'Is Deleted' -eq 'True') { $stat = 'Deleted' }");
                sb.AppendLine("    $own = ([string]$rr.'Owner Principal Name').Trim(); if (-not $own) { $own = ([string]$rr.'Owner Display Name').Trim() }");
                sb.AppendLine("    if ($st) {");
                sb.AppendLine("        if ($st.webUrl) { $su = [string]$st.webUrl }");
                sb.AppendLine("        $ti = [string]$st.displayName; $cd = [string]$st.createdDateTime; $lm = [string]$st.lastModifiedDateTime");
                if (wantDrive)
                {
                    sb.AppendLine("        try {");
                    sb.AppendLine("            $dr = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/sites/\" + $sidRaw + \"/drive?`$select=quota,owner\") -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("            if ($dr.quota) {");
                    sb.AppendLine("                try { if ($null -ne $dr.quota.used) { $sumb = [math]::Round([double]$dr.quota.used / 1MB, 2) } } catch { }");
                    sb.AppendLine("                try { if ($null -ne $dr.quota.total) { $sqmb = [math]::Round([double]$dr.quota.total / 1MB, 2) } } catch { }");
                    sb.AppendLine("            }");
                    sb.AppendLine("            if ($dr.owner -and $dr.owner.group -and $dr.owner.group.id) { $gid = [string]$dr.owner.group.id }");
                    sb.AppendLine("        } catch { Write-Host (\"  WARNING: drive of site '{0}' failed: \" -f $su + $_.Exception.Message) }");
                }
                if (wantOwners)
                {
                    sb.AppendLine("        if ($gid) {");
                    sb.AppendLine("            try {");
                    sb.AppendLine("                $ol = @(); $ouri = \"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/owners?`$select=id,userPrincipalName,displayName\"");
                    sb.AppendLine("                while ($ouri) { $orr = Invoke-RestMethod -Uri $ouri -Headers $headers -Method Get -ErrorAction Stop; $ol += @($orr.value); $ouri = $orr.'@odata.nextLink' }");
                    sb.AppendLine("                $owns = (($ol | ForEach-Object { $u = [string]$_.userPrincipalName; if (-not $u) { $u = [string]$_.displayName }; $u }) | Where-Object { $_ -ne '' }) -join ','");
                    sb.AppendLine("            } catch { }");
                    sb.AppendLine("        }");
                }
                if (wantLabel)
                {
                    sb.AppendLine("        if ($gid) {");
                    sb.AppendLine("            try {");
                    sb.AppendLine("                $gg = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"?`$select=assignedLabels\") -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("                $slid = ((@($gg.assignedLabels) | ForEach-Object { [string]$_.labelId }) | Where-Object { $_ -ne '' }) -join ','");
                    sb.AppendLine("            } catch { }");
                    sb.AppendLine("        }");
                }
                if (wantTeamsConn)
                    sb.AppendLine("        if ($gid -and $teamIds.ContainsKey($gid.Trim().ToLower())) { $team = 'Yes' }");
                sb.AppendLine("    }");
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        SiteId = $sidRaw; Url = $su; Title = $ti; Template = $tpl; GroupId = $gid;");
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
                "Audit webs below every site collection via Graph: one row per web (root webs excluded). Web-level config (template, language, unique permissions) is NOT exposed by Graph - see sp-migration/.",
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
                sb.AppendLine(FetchUsageReport);
                sb.AppendLine(FilterWorkRows);
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
                sb.AppendLine("foreach ($rr in $work) {");
                sb.AppendLine("    $i++; if (($i % 20) -eq 0) { Write-Host (\"  ... {0}/{1}\" -f $i, $work.Count) }");
                sb.AppendLine("    $sidRaw = ([string]$rr.'Site Id').Trim()");
                sb.AppendLine("    if (-not $sidRaw) { continue }");
                sb.AppendLine("    $rt = $null");
                sb.AppendLine("    try { $rt = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/sites/\" + $sidRaw) -Headers $headers -Method Get -ErrorAction Stop } catch { Write-Host (\"  WARNING: site {0} failed: \" -f $sidRaw + $_.Exception.Message) }");
                sb.AppendLine("    if (-not $rt) { continue }");
                sb.AppendLine("    $rp = ([string]$rt.id).Split(',')");
                sb.AppendLine("    $rweb = ''; if ($rp.Count -ge 3) { $rweb = $rp[2] }");
                sb.AppendLine("    $stack = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("    $stack.Add([pscustomobject]@{ Sid = [string]$rt.id; Parent = $rweb })");
                sb.AppendLine("    while ($stack.Count -gt 0) {");
                sb.AppendLine("        $tp = $stack[$stack.Count - 1]; $stack.RemoveAt($stack.Count - 1)");
                sb.AppendLine("        foreach ($kd in @(Get-Kids $tp.Sid)) {");
                sb.AppendLine("            $kcid = [string]$kd.id; $kp = $kcid.Split(',')");
                sb.AppendLine("            $kw = ''; if ($kp.Count -ge 3) { $kw = $kp[2] }");
                sb.AppendLine("            $rows.Add([pscustomobject]@{");
                sb.AppendLine("                SiteId = $sidRaw; WebId = $kw; ParentWebId = $tp.Parent;");
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
    }
}
