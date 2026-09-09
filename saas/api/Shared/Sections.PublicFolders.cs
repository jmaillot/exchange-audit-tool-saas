using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    internal static class SectionsPublicFolders
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildPfMailboxSection());
            AuditRegistry.Register(BuildPfHierarchySection());
            AuditRegistry.Register(BuildMailPfSection());
        }

        private static readonly List<string> RecipientRefs = new List<string>(new string[]
        {
            "GrantSendOnBehalfTo", "AcceptMessagesOnlyFromSendersOrMembers", "RejectMessagesFromSendersOrMembers", "ModeratedBy", "ForwardingAddress"
        });

        // ============================================================ 1. PUBLIC FOLDER MAILBOXES
        private static AuditSection BuildPfMailboxSection()
        {
            var section = new AuditSection(
                "pf-mailboxes",
                "PF mailboxes",
                "Public folder mailbox export",
                "Audit public folder mailboxes (host the hierarchy and content).",
                "folder",
                AuditScope.Both);
            section.Category = "Public Folders";
            section.DefaultFileName = "PublicFolderMailboxes.csv";

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("Name", false);
            identity.AddProp("Identity", false);
            identity.AddProp("PrimarySmtpAddress", true);
            identity.AddProp("UserPrincipalName", false);
            identity.AddProp("EmailAddresses", false);
            identity.AddProp("RecipientTypeDetails", true);
            identity.AddProp("ExchangeGuid", false);
            section.AddGroup(identity);

            var role = new AuditOptionGroup("role", "Hierarchy role", GroupMode.MultiCheck); role.Columns = 2;
            role.Hint = "IsRootPublicFolderMailbox marks the primary hierarchy mailbox.";
            role.AddProp("IsRootPublicFolderMailbox", true);
            role.AddProp("IsHierarchyReady", false);
            role.AddProp("IsHierarchySyncEnabled", false);
            role.AddProp("IsExcludedFromServingHierarchy", true);
            role.AddProp("IsPublicFolderSystemMailbox", false);
            section.AddGroup(role);

            var quotas = new AuditOptionGroup("quotas", "Quotas", GroupMode.MultiCheck); quotas.Columns = 2;
            quotas.AddProp("ProhibitSendQuota", false);
            quotas.AddProp("ProhibitSendReceiveQuota", false);
            quotas.AddProp("IssueWarningQuota", false);
            quotas.AddProp("RecoverableItemsQuota", false);
            quotas.AddProp("RecoverableItemsWarningQuota", false);
            quotas.AddProp("UseDatabaseQuotaDefaults", false);
            quotas.AddProp("MaxSendSize", false);
            quotas.AddProp("MaxReceiveSize", false);
            section.AddGroup(quotas);

            var retention = new AuditOptionGroup("retention", "Retention / compliance", GroupMode.MultiCheck); retention.Columns = 2;
            retention.AddProp("RetentionPolicy", false);
            retention.AddProp("RetainDeletedItemsFor", false);
            retention.AddProp("LitigationHoldEnabled", false);
            retention.AddProp("SingleItemRecoveryEnabled", false);
            retention.AddProp("InPlaceHolds", false);
            retention.AddProp("AuditEnabled", false);
            section.AddGroup(retention);

            var visibility = new AuditOptionGroup("visibility", "Visibility / address policy", GroupMode.MultiCheck); visibility.Columns = 2;
            visibility.AddProp("HiddenFromAddressListsEnabled", false);
            visibility.AddProp("EmailAddressPolicyEnabled", false);
            visibility.AddProp("PoliciesIncluded", false);
            section.AddGroup(visibility);

            var location = new AuditOptionGroup("location", "Location", GroupMode.MultiCheck); location.Columns = 2;
            location.Add(AuditOption.PropScoped("Database", false, true));
            location.Add(AuditOption.PropScoped("ServerName", false, true));
            section.AddGroup(location);

            var extra = new AuditOptionGroup("complementary", "Complementary data", GroupMode.MultiCheck); extra.Columns = 1;
            extra.Add(new AuditOption("mailboxsize", "Mailbox size in MB (Get-MailboxStatistics)", "mailboxsize", false).MarkSlow());
            extra.Add(new AuditOption("mailboxitemcount", "Item count (Get-MailboxStatistics)", "mailboxitemcount", false).MarkSlow());
            section.AddGroup(extra);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");

                var chosen = new List<string>();
                foreach (string gk in new string[] { "identity", "role", "quotas", "retention", "visibility", "location" })
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("DisplayName");

                var exprs = new List<string>();
                foreach (string v in chosen)
                {
                    if (v == "EmailAddresses")
                        exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                    else if (v == "InPlaceHolds" || v == "PoliciesIncluded")
                        exprs.Add("@{Name='" + v + "';Expression={($_." + v + " | ForEach-Object { [string]$_ }) -join ','}}");
                    else exprs.Add(v);
                }
                string selectList = string.Join(", ", exprs.ToArray());

                bool needSize = sel.IsSelected("complementary", "mailboxsize");
                bool needCount = sel.IsSelected("complementary", "mailboxitemcount");
                bool online = ConnectionSettings.IsOnline;
                string getStats = online ? "Get-EXOMailboxStatistics" : "Get-MailboxStatistics";

                var sb = new StringBuilder();
                if (needSize) PsScriptHelpers.EmitSizeHelper(sb, getStats);
                if (needCount) PsScriptHelpers.EmitCountHelper(sb, getStats);

                sb.AppendLine("Write-Host 'Querying public folder mailboxes...'");
                sb.AppendLine("$mbx = @(Get-Mailbox -PublicFolder -ResultSize " + resultSize + ")");
                sb.AppendLine("Write-Host (\"Retrieved {0} public folder mailbox(es).\" -f $mbx.Count)");
                sb.AppendLine();

                if (!needSize && !needCount)
                {
                    sb.AppendLine("$rows = $mbx | Select-Object " + selectList);
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($m in $mbx) {");
                    sb.AppendLine("    $obj = $m | Select-Object " + selectList);
                    if (needSize)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName MailboxSizeMB -NotePropertyValue (Get-SizeMB $m.PrimarySmtpAddress) -Force");
                    if (needCount)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName MailboxItemCount -NotePropertyValue (Get-ItemCount $m.PrimarySmtpAddress) -Force");
                    sb.AppendLine("    $obj");
                    sb.AppendLine("}");
                }

                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 2. PUBLIC FOLDER HIERARCHY
        private static AuditSection BuildPfHierarchySection()
        {
            var section = new AuditSection(
                "pf-hierarchy",
                "PF hierarchy",
                "Public folder hierarchy export",
                "Audit the public folder tree, statistics and client permissions.",
                "folder",
                AuditScope.Both);
            section.Category = "Public Folders";
            section.DefaultFileName = "PublicFolderHierarchy.csv";

            var tree = new AuditOptionGroup("tree", "Tree", GroupMode.MultiCheck); tree.Columns = 2;
            tree.Hint = "Identity is the folder path (\\Root\\Sub...).";
            tree.AddProp("Identity", true);
            tree.AddProp("Name", true);
            tree.AddProp("ParentPath", false);
            tree.AddProp("FolderClass", true);
            tree.AddProp("MailEnabled", true);
            tree.AddProp("ContentMailboxName", true);
            section.AddGroup(tree);

            var quotas = new AuditOptionGroup("quotas", "Quotas (Get-PublicFolder)", GroupMode.MultiCheck); quotas.Columns = 2;
            quotas.AddProp("ProhibitPostQuota", false);
            quotas.AddProp("IssueWarningQuota", false);
            section.AddGroup(quotas);

            var retention = new AuditOptionGroup("retention", "Retention (Get-PublicFolder)", GroupMode.MultiCheck); retention.Columns = 2;
            retention.AddProp("AgeLimit", false);
            retention.AddProp("RetainDeletedItemsFor", false);
            section.AddGroup(retention);

            var stats = new AuditOptionGroup("stats", "Statistics (Get-PublicFolderStatistics, slower)", GroupMode.MultiCheck).MarkSlow(); stats.Columns = 2;
            stats.Hint = "One call per folder; can be slow on large trees. MaxItemSize is read from the folder object.";
            stats.Add(new AuditOption("ItemCount", "Item count", "ItemCount", false));
            stats.Add(new AuditOption("SizeMB", "Total size (MB)", "SizeMB", false));
            stats.Add(new AuditOption("MaxItemSize", "Max item size", "MaxItemSize", false));
            stats.Add(new AuditOption("LastModificationTime", "Last modified", "LastModificationTime", false));
            section.AddGroup(stats);

            var perms = new AuditOptionGroup("perms", "Client permissions (Get-PublicFolderClientPermission, slower)", GroupMode.MultiCheck).MarkSlow();
            perms.Hint = "THE critical data for migration. Inherited is a heuristic (same user+rights present on the parent folder).";
            perms.Columns = 1;
            perms.Add(new AuditOption("clientperms", "Include client permissions", "clientperms", false));
            perms.Add(new AuditOption("inherited", "Add 'Inherited' indicator (compares with parent folder)", "inherited", false));
            section.AddGroup(perms);

            var mode = new AuditOptionGroup("mode", "Output mode", GroupMode.SingleChoice); mode.Columns = 1;
            mode.Add(new AuditOption("folders", "Folder rows only", "folders", true));
            mode.Add(new AuditOption("expandperms", "Expand permissions (one row per folder + user)", "expandperms", false));
            section.AddGroup(mode);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");
                string mode2 = sel.First("mode", "folders");

                var chosen = new List<string>();
                // tree + quotas + retention are all direct Get-PublicFolder properties.
                foreach (string gk in new string[] { "tree", "quotas", "retention" })
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("Identity");
                var baseSelect = new List<string>(chosen);
                // MaxItemSize sits in the Statistics group but is a direct folder property -> add it here.
                bool wantMaxItemSize = sel.IsSelected("stats", "MaxItemSize");
                if (wantMaxItemSize && !baseSelect.Contains("MaxItemSize")) baseSelect.Add("MaxItemSize");
                if (!baseSelect.Contains("Identity")) baseSelect.Add("Identity");
                string selectList = string.Join(", ", baseSelect.ToArray());

                List<string> stat = sel.Selected("stats");
                bool clientPerms = sel.IsSelected("perms", "clientperms");
                bool inherited = sel.IsSelected("perms", "inherited");
                bool expandPerms = mode2 == "expandperms";

                var sb = new StringBuilder();

                sb.AppendLine("Write-Host 'Querying public folder hierarchy (recursive)...'");
                sb.AppendLine("$folders = @(Get-PublicFolder -Recurse -ResultSize " + resultSize + " -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} public folder(s).\" -f $folders.Count)");
                sb.AppendLine();

                // Build a permission index (path -> user -> rights) so inheritance can be derived
                // by comparing each folder's entries with its parent folder's entries.
                if (clientPerms)
                {
                    sb.AppendLine("Write-Host 'Indexing client permissions...'");
                    sb.AppendLine("$permIndex = @{}");
                    sb.AppendLine("foreach ($f in $folders) {");
                    sb.AppendLine("    $m = @{}");
                    sb.AppendLine("    try {");
                    sb.AppendLine("        Get-PublicFolderClientPermission -Identity $f.Identity -ErrorAction Stop |");
                    sb.AppendLine("            Where-Object { $_.User.DisplayName -ne 'Default' -and $_.User.DisplayName -ne 'Anonymous' } |");
                    sb.AppendLine("            ForEach-Object { $m[[string]$_.User.DisplayName] = (($_.AccessRights) -join ',') }");
                    sb.AppendLine("    } catch { }");
                    sb.AppendLine("    $permIndex[[string]$f.Identity] = $m");
                    sb.AppendLine("}");
                    sb.AppendLine("function Get-PfParent { param($p); $i = $p.LastIndexOf('\\'); if ($i -le 0) { return '\\' } else { return $p.Substring(0,$i) } }");
                    sb.AppendLine("function Test-PfInherited { param($path,$user,$rights)");
                    sb.AppendLine("    $pp = Get-PfParent $path");
                    sb.AppendLine("    if ($permIndex.ContainsKey($pp)) { $pm = $permIndex[$pp]; if ($pm.ContainsKey($user) -and $pm[$user] -eq $rights) { return 'Yes' } }");
                    sb.AppendLine("    return 'No' }");
                    sb.AppendLine();
                }

                if (expandPerms && clientPerms)
                {
                    sb.AppendLine("$rows = foreach ($f in $folders) {");
                    sb.AppendLine("    $pmap = $permIndex[[string]$f.Identity]");
                    sb.AppendLine("    if (-not $pmap -or $pmap.Count -eq 0) {");
                    sb.AppendLine("        [PSCustomObject]@{ FolderPath=[string]$f.Identity; Name=$f.Name; FolderClass=$f.FolderClass; MailEnabled=$f.MailEnabled; ContentMailbox=$f.ContentMailboxName; User=''; AccessRights=''" + (inherited ? "; Inherited=''" : "") + " }");
                    sb.AppendLine("    } else {");
                    sb.AppendLine("        foreach ($usr in $pmap.Keys) {");
                    sb.AppendLine("            $rights = $pmap[$usr]");
                    if (inherited)
                        sb.AppendLine("            [PSCustomObject]@{ FolderPath=[string]$f.Identity; Name=$f.Name; FolderClass=$f.FolderClass; MailEnabled=$f.MailEnabled; ContentMailbox=$f.ContentMailboxName; User=$usr; AccessRights=$rights; Inherited=(Test-PfInherited ([string]$f.Identity) $usr $rights) }");
                    else
                        sb.AppendLine("            [PSCustomObject]@{ FolderPath=[string]$f.Identity; Name=$f.Name; FolderClass=$f.FolderClass; MailEnabled=$f.MailEnabled; ContentMailbox=$f.ContentMailboxName; User=$usr; AccessRights=$rights }");
                    sb.AppendLine("        }");
                    sb.AppendLine("    }");
                    sb.AppendLine("}");
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($f in $folders) {");
                    sb.AppendLine("    $obj = $f | Select-Object " + selectList);
                    // Only ItemCount/SizeMB/LastModificationTime require Get-PublicFolderStatistics.
                    bool needStatsCall = stat.Contains("ItemCount") || stat.Contains("SizeMB") || stat.Contains("LastModificationTime");
                    if (needStatsCall)
                    {
                        sb.AppendLine("    try { $st = Get-PublicFolderStatistics -Identity $f.Identity -ErrorAction Stop } catch { $st = $null }");
                        if (stat.Contains("ItemCount"))
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName ItemCount -NotePropertyValue ($(if ($st) { $st.ItemCount } else { '' })) -Force");
                        if (stat.Contains("SizeMB"))
                        {
                            sb.AppendLine("    $szMB = ''");
                            sb.AppendLine("    if ($st -and $st.TotalItemSize) { $s=$st.TotalItemSize.ToString(); if ($s -match '\\(([\\d,]+) bytes\\)') { $szMB=[math]::Round(([double]($matches[1] -replace ',','')) / 1MB, 2) } }");
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName SizeMB -NotePropertyValue $szMB -Force");
                        }
                        if (stat.Contains("LastModificationTime"))
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName LastModificationTime -NotePropertyValue ($(if ($st) { [string]$st.LastModificationTime } else { '' })) -Force");
                    }
                    if (clientPerms)
                    {
                        // Combined "User=Rights" entries; when Inherited requested, add "(inherited)" marker.
                        sb.AppendLine("    $pmap = $permIndex[[string]$f.Identity]");
                        sb.AppendLine("    $entries = New-Object System.Collections.Generic.List[string]");
                        sb.AppendLine("    if ($pmap) { foreach ($usr in $pmap.Keys) {");
                        sb.AppendLine("        $rights = $pmap[$usr]");
                        if (inherited)
                            sb.AppendLine("        $inh = Test-PfInherited ([string]$f.Identity) $usr $rights; $entries.Add($usr + '=' + $rights + '(' + $(if ($inh -eq 'Yes') { 'inherited' } else { 'explicit' }) + ')')");
                        else
                            sb.AppendLine("        $entries.Add($usr + '=' + $rights)");
                        sb.AppendLine("    } }");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName ClientPermissions -NotePropertyValue ($entries -join ' | ') -Force");
                    }
                    sb.AppendLine("    $obj");
                    sb.AppendLine("}");
                }

                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 3. MAIL-ENABLED PUBLIC FOLDERS
        private static AuditSection BuildMailPfSection()
        {
            var section = new AuditSection(
                "mail-pf",
                "Mail-enabled PF",
                "Mail-enabled public folder export",
                "Audit mail-enabled public folders (they have an SMTP address, like recipients).",
                "folder",
                AuditScope.Both);
            section.Category = "Public Folders";
            section.DefaultFileName = "MailEnabledPublicFolders.csv";

            var identity = new AuditOptionGroup("identity", "Identity / addressing", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("PrimarySmtpAddress", true);
            identity.AddProp("EmailAddresses", true);
            identity.AddProp("HiddenFromAddressListsEnabled", true);
            section.AddGroup(identity);

            var location = new AuditOptionGroup("location", "Location", GroupMode.MultiCheck); location.Columns = 2;
            location.Hint = "ContentMailbox = the public folder mailbox hosting this folder's content.";
            location.AddProp("ContentMailbox", true);
            section.AddGroup(location);

            var mailflow = new AuditOptionGroup("mailflow", "Mail flow", GroupMode.MultiCheck); mailflow.Columns = 2;
            mailflow.AddProp("RequireSenderAuthenticationEnabled", false);
            mailflow.AddProp("AcceptMessagesOnlyFromSendersOrMembers", false);
            mailflow.AddProp("RejectMessagesFromSendersOrMembers", false);
            section.AddGroup(mailflow);

            var routing = new AuditOptionGroup("routing", "Routing", GroupMode.MultiCheck); routing.Columns = 2;
            routing.AddProp("DeliverToMailboxAndForward", false);
            routing.AddProp("ForwardingAddress", false);
            routing.AddProp("ExternalEmailAddress", false);
            section.AddGroup(routing);

            var quotas = new AuditOptionGroup("quotas", "Quotas", GroupMode.MultiCheck); quotas.Columns = 2;
            quotas.AddProp("MaxSendSize", false);
            quotas.AddProp("MaxReceiveSize", false);
            section.AddGroup(quotas);

            var moderation = new AuditOptionGroup("moderation", "Moderation", GroupMode.MultiCheck); moderation.Columns = 2;
            moderation.AddProp("ModerationEnabled", false);
            moderation.AddProp("ModeratedBy", false);
            section.AddGroup(moderation);

            var permissions = new AuditOptionGroup("permissions", "Permissions", GroupMode.MultiCheck); permissions.Columns = 1;
            permissions.Hint = "Send As (EXO: RecipientPermission / on-prem: ADPermission); Send on Behalf from GrantSendOnBehalfTo. Both resolved to SMTP.";
            permissions.Add(new AuditOption("sendas", "Send As delegates (EXO: RecipientPermission / on-prem: ADPermission)", "sendas", false));
            permissions.Add(new AuditOption("sendonbehalf", "Send on Behalf (GrantSendOnBehalfTo)", "sendonbehalf", false));
            section.AddGroup(permissions);

            var custom = new AuditOptionGroup("custom", "Custom attributes", GroupMode.MultiCheck); custom.Columns = 2;
            custom.Add(new AuditOption("customattr", "CustomAttribute1-15", "CustomAttribute1-15", false));
            section.AddGroup(custom);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");

                var chosen = new List<string>();
                foreach (string gk in new string[] { "identity", "location", "mailflow", "routing", "quotas", "moderation", "custom" })
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("DisplayName");

                bool sa = sel.IsSelected("permissions", "sendas");
                bool sob = sel.IsSelected("permissions", "sendonbehalf");

                bool needResolver = sob;
                var exprs = new List<string>();
                foreach (string v in chosen)
                {
                    if (v == "EmailAddresses")
                        exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                    else if (RecipientRefs.Contains(v)) { needResolver = true; exprs.Add("@{Name='" + v + "';Expression={ Resolve-Recip $_." + v + " }}"); }
                    else if (v == "CustomAttribute1-15")
                        for (int i = 1; i <= 15; i++) exprs.Add("CustomAttribute" + i);
                    else exprs.Add(v);
                }
                string selectList = string.Join(", ", exprs.ToArray());

                bool online = ConnectionSettings.IsOnline;

                var sb = new StringBuilder();
                if (needResolver) { string getRecip = online ? "Get-EXORecipient" : "Get-Recipient"; PsScriptHelpers.EmitResolver(sb, getRecip, false); }

                sb.AppendLine("Write-Host 'Querying mail-enabled public folders...'");
                sb.AppendLine("$items = @(Get-MailPublicFolder -ResultSize " + resultSize + ")");
                sb.AppendLine("Write-Host (\"Retrieved {0} mail-enabled public folder(s).\" -f $items.Count)");
                sb.AppendLine();

                if (!sa && !sob)
                {
                    sb.AppendLine("$rows = $items | Select-Object " + selectList);
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($m in $items) {");
                    sb.AppendLine("    $obj = $m | Select-Object " + selectList);
                    if (sa)
                    {
                        sb.AppendLine("    try {");
                        if (online)
                        {
                            // Exchange Online: Send As via (EXO)RecipientPermission.
                            sb.AppendLine("        $saList = @(Get-EXORecipientPermission -Identity $m.Identity -ErrorAction Stop |");
                            sb.AppendLine("            Where-Object { $_.AccessRights -contains 'SendAs' -and $_.Trustee -notlike 'NT AUTHORITY\\*' } |");
                            sb.AppendLine("            Select-Object -ExpandProperty Trustee)");
                        }
                        else
                        {
                            // On-premises: Send As is the AD 'Send-As' extended right (Get-ADPermission).
                            sb.AppendLine("        $saList = @(Get-ADPermission -Identity $m.Identity -ErrorAction Stop |");
                            sb.AppendLine("            Where-Object { ($_.ExtendedRights -like '*Send-As*') -and (-not $_.IsInherited) -and (-not $_.Deny) -and ($_.User -notlike 'NT AUTHORITY\\*') } |");
                            sb.AppendLine("            Select-Object -ExpandProperty User)");
                        }
                        sb.AppendLine("    } catch { $saList = @() }");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName SendAs -NotePropertyValue (($saList | ForEach-Object { [string]$_ }) -join ',') -Force");
                    }
                    if (sob)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName SendOnBehalf -NotePropertyValue (Resolve-Recip $m.GrantSendOnBehalfTo) -Force");
                    sb.AppendLine("    $obj");
                    sb.AppendLine("}");
                }

                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }
    }
}
