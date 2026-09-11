using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "Teams" category (Microsoft Graph, NOT Exchange Online):
    //   teams-inventory            -> GET /groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')
    //     (one row per team) + GET /teams/{id} (Archived) + owners/members/$count
    //     (Owner/Member/Guest counts) + /groups/{id}/sites/root (SiteUrl).
    //   teams-channels             -> GET /teams/{id}/channels (one row per channel)
    //     + /channels/{id}/filesFolder (SharePointUrl) + /channels/{id}/members (counts).
    //   teams-channel-memberships  -> private/shared channels only, one row per
    //     channel member (user lookup cached for UPN/UserType).
    //   teams-members              -> GET /groups/{id}/owners + /groups/{id}/members
    //     (one row per owner/member, guests flagged with external domain).
    // All emitted cells are pre-joined scalars, so Select-Object takes plain
    // property names and the ';' CSV never clashes (multi-values are ','-joined).
    // Required Graph delegated scopes: Group.Read.All, Team.ReadBasic.All,
    // Channel.ReadBasic.All, TeamMember.Read.All, ChannelMember.Read.All
    // (plus User.Read.All already used by Licensing).
    internal static class SectionsTeams
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildInventorySection());
            AuditRegistry.Register(BuildChannelsSection());
            AuditRegistry.Register(BuildChannelMembershipsSection());
            AuditRegistry.Register(BuildMembersSection());
            AuditRegistry.Register(BuildAppsSection());
            AuditRegistry.Register(BuildTabsSection());
            AuditRegistry.Register(BuildSitesSection());
            AuditRegistry.Register(BuildActivitySection());
            AuditRegistry.Register(BuildComplianceSection());
        }

        private static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            return PsScriptHelpers.Collect(sel, groupKeys);
        }

        // Prefer header: membershipType is an evolvable enum - without it Graph
        // returns 'unknownFutureValue' instead of 'shared' for shared channels.
        private const string GraphHeaders =
            "$headers = @{ Authorization = ('Bearer ' + $env:EAT_GRAPH_TOKEN); ConsistencyLevel = 'eventual'; Prefer = 'include-unknown-enum-members' }";

        private const string FetchTeamsGroups =
            "$groups = @()\n" +
            "$uri = 'https://graph.microsoft.com/v1.0/groups?$count=true&$filter=resourceProvisioningOptions/Any(x:x eq ''Team'')&$select=id,displayName,mail,description,visibility,groupTypes,createdDateTime,classification,expirationDateTime,assignedLabels&$top=999'\n" +
            "while ($uri -and ($maxTeams -lt 0 -or $groups.Count -lt $maxTeams)) {\n" +
            "    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop\n" +
            "    $groups += @($r.value)\n" +
            "    $uri = $r.'@odata.nextLink'\n" +
            "}\n" +
            "if ($maxTeams -ge 0 -and $groups.Count -gt $maxTeams) { $groups = $groups[0..($maxTeams - 1)] }";

        private static string SizePrelude(string size)
        {
            // NOTE: plain concatenation, NOT string.Format: the PowerShell body
            // contains literal { } blocks that string.Format would try to parse.
            return "$size = '" + size + "'\n$maxTeams = -1\nif ($size -eq '100') { $maxTeams = 100 } elseif ($size -eq '1000') { $maxTeams = 1000 }";
        }

        // ============================================================ 1. INVENTORY
        private static AuditSection BuildInventorySection()
        {
            var section = new AuditSection(
                "teams-inventory",
                "List",
                "List export",
                "Audit Microsoft Teams teams via Graph: identity, visibility, archive state plus owner/member/guest counts.",
                "group",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsInventory.csv";

            var identity = new AuditOptionGroup("identity", "Identity & addressing", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("GroupId", true);
            identity.AddProp("PrimarySMTP", true);
            identity.AddProp("Description", true);
            section.AddGroup(identity);

            var team = new AuditOptionGroup("team", "Team properties", GroupMode.MultiCheck); team.Columns = 2;
            team.AddProp("Visibility", true);
            team.AddProp("IsMembershipDynamic", true);
            team.AddProp("SiteUrl", true);
            team.AddProp("Archived", true);
            section.AddGroup(team);

            var counts = new AuditOptionGroup("counts", "Membership counts (per-team lookups)", GroupMode.MultiCheck); counts.Columns = 3;
            counts.Hint = "Counts run one owners/members query per team - slower on large tenants. Uncheck for a fast first pass.";
            counts.AddProp("OwnerCount", true);
            counts.AddProp("MemberCount", true);
            counts.AddProp("GuestCount", true);
            section.AddGroup(counts);

            var stats = new AuditOptionGroup("stats", "Team stats (per-team lookups)", GroupMode.MultiCheck); stats.Columns = 3;
            stats.Hint = "Channels/apps/tabs are counted per team (tabs: one query per channel). Uncheck for a fast first pass.";
            stats.AddProp("ChannelsCount", true);
            stats.AddProp("PrivateChannelsCount", true);
            stats.AddProp("SharedChannelsCount", true);
            stats.AddProp("AppsCount", true);
            stats.AddProp("TabsCount", true);
            stats.AddProp("HasExternalGuests", true);
            section.AddGroup(stats);

            var meta = new AuditOptionGroup("meta", "Classification & lifecycle", GroupMode.MultiCheck); meta.Columns = 2;
            meta.AddProp("CreatedDate", true);
            section.AddGroup(meta);

            var settings = new AuditOptionGroup("settings", "Teams settings (same per-team lookup as Archived)", GroupMode.MultiCheck); settings.Columns = 2;
            settings.Hint = "Guest/member channel rights, messaging rights and fun settings, read from the team object.";
            settings.AddProp("AllowGuestCreateUpdateChannels", true);
            settings.AddProp("AllowGuestDeleteChannels", true);
            settings.AddProp("AllowMemberCreatePrivateChannels", true);
            settings.AddProp("AllowCreateUpdateChannels", true);
            settings.AddProp("AllowDeleteChannels", true);
            settings.AddProp("AllowAddRemoveApps", true);
            settings.AddProp("AllowCreateUpdateRemoveTabs", true);
            settings.AddProp("AllowCreateUpdateRemoveConnectors", true);
            settings.AddProp("AllowMemberDeleteMessages", true);
            settings.AddProp("AllowMemberEditMessages", true);
            settings.AddProp("AllowOwnerDeleteMessages", true);
            settings.AddProp("AllowTeamMentions", true);
            settings.AddProp("AllowChannelMentions", true);
            settings.AddProp("GiphyEnabled", true);
            settings.AddProp("GiphyContentRating", true);
            settings.AddProp("AllowStickersAndMemes", true);
            settings.AddProp("AllowCustomMemes", true);
            section.AddGroup(settings);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "team", "counts", "stats", "meta", "settings");
                if (chosen.Count == 0) chosen.Add("DisplayName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                bool wantOwners = chosen.Contains("OwnerCount");
                bool wantMembers = chosen.Contains("MemberCount");
                bool wantGuestN = chosen.Contains("GuestCount") || chosen.Contains("HasExternalGuests");
                bool wantCounts = wantOwners || wantMembers || wantGuestN;
                bool wantArchived = chosen.Contains("Archived");
                bool wantSite = chosen.Contains("SiteUrl");
                bool wantSettings = Collect(sel, "settings").Count > 0;
                bool wantTeam = wantArchived || wantSettings;
                bool wantChannels = chosen.Contains("ChannelsCount") || chosen.Contains("PrivateChannelsCount")
                    || chosen.Contains("SharedChannelsCount") || chosen.Contains("TabsCount");
                bool wantTabs = chosen.Contains("TabsCount");
                bool wantApps = chosen.Contains("AppsCount");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                if (wantCounts || wantApps || wantChannels) sb.AppendLine("$countFail = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; if (($i % 20) -eq 0) { Write-Host (\"  ... {0}/{1}\" -f $i, $groups.Count) }");
                sb.AppendLine("    $gid = [string]$g.id");
                sb.AppendLine("    $dyn = 'No'; if (@($g.groupTypes) -contains 'DynamicMembership') { $dyn = 'Yes' }");
                if (wantTeam)
                {
                    sb.AppendLine("    $arch = ''; $sgcu = ''; $sgd = ''; $smpc = ''; $smcuc = ''; $smdc = ''; $smara = ''; $smct = ''; $smcc = ''; $smdm = ''; $smem = ''; $smod = ''; $smtm = ''; $smcm = ''; $fg = ''; $fgr = ''; $fsm = ''; $fcm = ''");
                    sb.AppendLine("    try {");
                    sb.AppendLine("        $td = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/teams/\" + $gid) -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("        if ($td.isArchived) { $arch = 'Yes' } else { $arch = 'No' }");
                    sb.AppendLine("        if ($td.guestSettings.allowCreateUpdateChannels) { $sgcu = 'Yes' } else { $sgcu = 'No' }");
                    sb.AppendLine("        if ($td.guestSettings.allowDeleteChannels) { $sgd = 'Yes' } else { $sgd = 'No' }");
                    sb.AppendLine("        if ($td.memberSettings.allowCreatePrivateChannels) { $smpc = 'Yes' } else { $smpc = 'No' }");
                    sb.AppendLine("        if ($td.memberSettings.allowCreateUpdateChannels) { $smcuc = 'Yes' } else { $smcuc = 'No' }");
                    sb.AppendLine("        if ($td.memberSettings.allowDeleteChannels) { $smdc = 'Yes' } else { $smdc = 'No' }");
                    sb.AppendLine("        if ($td.memberSettings.allowAddRemoveApps) { $smara = 'Yes' } else { $smara = 'No' }");
                    sb.AppendLine("        if ($td.memberSettings.allowCreateUpdateRemoveTabs) { $smct = 'Yes' } else { $smct = 'No' }");
                    sb.AppendLine("        if ($td.memberSettings.allowCreateUpdateRemoveConnectors) { $smcc = 'Yes' } else { $smcc = 'No' }");
                    sb.AppendLine("        if ($td.messagingSettings.allowUserDeleteMessages) { $smdm = 'Yes' } else { $smdm = 'No' }");
                    sb.AppendLine("        if ($td.messagingSettings.allowUserEditMessages) { $smem = 'Yes' } else { $smem = 'No' }");
                    sb.AppendLine("        if ($td.messagingSettings.allowOwnerDeleteMessages) { $smod = 'Yes' } else { $smod = 'No' }");
                    sb.AppendLine("        if ($td.messagingSettings.allowTeamMentions) { $smtm = 'Yes' } else { $smtm = 'No' }");
                    sb.AppendLine("        if ($td.messagingSettings.allowChannelMentions) { $smcm = 'Yes' } else { $smcm = 'No' }");
                    sb.AppendLine("        if ($td.funSettings.allowGiphy) { $fg = 'Yes' } else { $fg = 'No' }");
                    sb.AppendLine("        $fgr = ''; if ($td.funSettings.giphyContentRating) { $fgr = [string]$td.funSettings.giphyContentRating }");
                    sb.AppendLine("        if ($td.funSettings.allowStickersAndMemes) { $fsm = 'Yes' } else { $fsm = 'No' }");
                    sb.AppendLine("        if ($td.funSettings.allowCustomMemes) { $fcm = 'Yes' } else { $fcm = 'No' }");
                    sb.AppendLine("    } catch { }");
                }
                if (wantSite)
                {
                    sb.AppendLine("    $site = ''");
                    sb.AppendLine("    try { $sd = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/sites/root?`$select=webUrl\") -Headers $headers -Method Get -ErrorAction Stop; $site = [string]$sd.webUrl } catch { $site = '' }");
                }
                if (wantCounts)
                {
                    sb.AppendLine("    $oc = 0; $mc = 0; $gc = 0; $heg = 'No'");
                    if (wantOwners)
                        sb.AppendLine("    try { $ow = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/owners?`$count=true&`$top=1\") -Headers $headers -Method Get -ErrorAction Stop; $oc = [int]$ow.'@odata.count' } catch { $countFail++ }");
                    if (wantMembers)
                        sb.AppendLine("    try { $mb = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/members?`$count=true&`$top=1\") -Headers $headers -Method Get -ErrorAction Stop; $mc = [int]$mb.'@odata.count' } catch { $countFail++ }");
                    if (wantGuestN)
                        sb.AppendLine("    try { $gu = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/members?`$count=true&`$filter=userType%20eq%20'Guest'&`$top=1\") -Headers $headers -Method Get -ErrorAction Stop; $gc = [int]$gu.'@odata.count'; if ($gc -gt 0) { $heg = 'Yes' } } catch { $countFail++ }");
                }
                if (wantChannels)
                {
                    sb.AppendLine("    $cht = ''; $chp = ''; $chs = ''; $tc = ''");
                    sb.AppendLine("    try {");
                    sb.AppendLine("        $chl = @(); $churi = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels\"");
                    sb.AppendLine("        while ($churi) { $chr = Invoke-RestMethod -Uri $churi -Headers $headers -Method Get -ErrorAction Stop; $chl += @($chr.value); $churi = $chr.'@odata.nextLink' }");
                    sb.AppendLine("        $cht = $chl.Count");
                    sb.AppendLine("        $chp = @($chl | Where-Object { [string]$_.membershipType -eq 'private' }).Count");
                    sb.AppendLine("        $chs = @($chl | Where-Object { ([string]$_.membershipType -eq 'shared') -or ([string]$_.membershipType -eq 'unknownFutureValue') }).Count");
                    if (wantTabs)
                    {
                        sb.AppendLine("        $tc = 0");
                        sb.AppendLine("        foreach ($cc in $chl) {");
                        sb.AppendLine("            try { $tbr = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + [string]$cc.id + \"/tabs\") -Headers $headers -Method Get -ErrorAction Stop; $tc += @($tbr.value).Count } catch { }");
                        sb.AppendLine("        }");
                    }
                    sb.AppendLine("    } catch { $countFail++ }");
                }
                if (wantApps)
                {
                    sb.AppendLine("    $ac = ''");
                    sb.AppendLine("    try { $ap = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/installedApps?`$count=true&`$top=1\") -Headers $headers -Method Get -ErrorAction Stop; $ac = [int]$ap.'@odata.count' } catch { $countFail++ }");
                }
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        DisplayName = $g.displayName; GroupId = $gid;");
                sb.AppendLine("        PrimarySMTP = $g.mail; Description = $g.description;");
                sb.AppendLine("        Visibility = $g.visibility; IsMembershipDynamic = $dyn;");
                sb.AppendLine("        SiteUrl = " + (wantSite ? "$site" : "''") + "; Archived = " + (wantArchived ? "$arch" : "''") + ";");
                sb.AppendLine("        OwnerCount = " + (wantOwners ? "$oc" : "''") + "; MemberCount = " + (wantMembers ? "$mc" : "''") + "; GuestCount = " + (chosen.Contains("GuestCount") ? "$gc" : "''") + ";");
                sb.AppendLine("        ChannelsCount = " + (chosen.Contains("ChannelsCount") ? "$cht" : "''") + "; PrivateChannelsCount = " + (chosen.Contains("PrivateChannelsCount") ? "$chp" : "''") + "; SharedChannelsCount = " + (chosen.Contains("SharedChannelsCount") ? "$chs" : "''") + ";");
                sb.AppendLine("        AppsCount = " + (wantApps ? "$ac" : "''") + "; TabsCount = " + (wantTabs ? "$tc" : "''") + "; HasExternalGuests = " + (chosen.Contains("HasExternalGuests") ? "$heg" : "''") + ";");
                sb.AppendLine("        CreatedDate = $g.createdDateTime;");
                sb.AppendLine("        AllowGuestCreateUpdateChannels = " + (chosen.Contains("AllowGuestCreateUpdateChannels") ? "$sgcu" : "''") + "; AllowGuestDeleteChannels = " + (chosen.Contains("AllowGuestDeleteChannels") ? "$sgd" : "''") + ";");
                sb.AppendLine("        AllowMemberCreatePrivateChannels = " + (chosen.Contains("AllowMemberCreatePrivateChannels") ? "$smpc" : "''") + ";");
                sb.AppendLine("        AllowMemberDeleteMessages = " + (chosen.Contains("AllowMemberDeleteMessages") ? "$smdm" : "''") + "; AllowMemberEditMessages = " + (chosen.Contains("AllowMemberEditMessages") ? "$smem" : "''") + "; AllowOwnerDeleteMessages = " + (chosen.Contains("AllowOwnerDeleteMessages") ? "$smod" : "''") + ";");
                sb.AppendLine("        AllowCreateUpdateChannels = " + (chosen.Contains("AllowCreateUpdateChannels") ? "$smcuc" : "''") + "; AllowDeleteChannels = " + (chosen.Contains("AllowDeleteChannels") ? "$smdc" : "''") + ";");
                sb.AppendLine("        AllowAddRemoveApps = " + (chosen.Contains("AllowAddRemoveApps") ? "$smara" : "''") + "; AllowCreateUpdateRemoveTabs = " + (chosen.Contains("AllowCreateUpdateRemoveTabs") ? "$smct" : "''") + "; AllowCreateUpdateRemoveConnectors = " + (chosen.Contains("AllowCreateUpdateRemoveConnectors") ? "$smcc" : "''") + ";");
                sb.AppendLine("        AllowTeamMentions = " + (chosen.Contains("AllowTeamMentions") ? "$smtm" : "''") + "; AllowChannelMentions = " + (chosen.Contains("AllowChannelMentions") ? "$smcm" : "''") + ";");
                sb.AppendLine("        GiphyEnabled = " + (chosen.Contains("GiphyEnabled") ? "$fg" : "''") + "; GiphyContentRating = " + (chosen.Contains("GiphyContentRating") ? "$fgr" : "''") + ";");
                sb.AppendLine("        AllowStickersAndMemes = " + (chosen.Contains("AllowStickersAndMemes") ? "$fsm" : "''") + "; AllowCustomMemes = " + (chosen.Contains("AllowCustomMemes") ? "$fcm" : "''"));
                sb.AppendLine("    })");
                sb.AppendLine("}");
                if (wantCounts || wantApps || wantChannels) sb.AppendLine("if ($countFail -gt 0) { Write-Host (\"WARNING: {0} count querie(s) failed (consent for Group.Read.All / TeamsAppInstallation.ReadForTeam / TeamsTab.Read.All?); affected counts show 0/blank.\" -f $countFail) }");
                sb.AppendLine("$rows = $rows | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 2. CHANNELS
        private static AuditSection BuildChannelsSection()
        {
            var section = new AuditSection(
                "teams-channels",
                "Channels",
                "Channels export",
                "Audit channels per team via Graph: type, SharePoint folder plus owner/member/guest counts.",
                "channel",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsChannels.csv";

            var identity = new AuditOptionGroup("identity", "Channel identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("TeamName", true);
            identity.AddProp("ChannelName", true);
            identity.AddProp("ChannelType", true);
            section.AddGroup(identity);

            var loc = new AuditOptionGroup("location", "Location & lifecycle", GroupMode.MultiCheck); loc.Columns = 2;
            loc.AddProp("SharePointUrl", true);
            loc.AddProp("CreatedDate", true);
            section.AddGroup(loc);

            var counts = new AuditOptionGroup("counts", "Membership counts (per-channel lookups)", GroupMode.MultiCheck); counts.Columns = 3;
            counts.Hint = "Counts query members + summary of every channel - slower on tenants with many channels.";
            counts.AddProp("OwnerCount", true);
            counts.AddProp("MemberCount", true);
            counts.AddProp("GuestCount", true);
            section.AddGroup(counts);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "location", "counts");
                if (chosen.Count == 0) chosen.Add("ChannelName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                bool wantSite = chosen.Contains("SharePointUrl");
                bool wantCounts = chosen.Contains("OwnerCount") || chosen.Contains("MemberCount") || chosen.Contains("GuestCount");
                bool wantGuests = chosen.Contains("GuestCount");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; Write-Host (\"  [{0}/{1}] {2}\" -f $i, $groups.Count, $g.displayName)");
                sb.AppendLine("    $gid = [string]$g.id; $tname = [string]$g.displayName");
                sb.AppendLine("    $channels = @()");
                sb.AppendLine("    try {");
                sb.AppendLine("        $curi = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels\"");
                sb.AppendLine("        while ($curi) { $cr = Invoke-RestMethod -Uri $curi -Headers $headers -Method Get -ErrorAction Stop; $channels += @($cr.value); $curi = $cr.'@odata.nextLink' }");
                sb.AppendLine("    } catch { Write-Host (\"  WARNING: channels of '{0}' failed: \" -f $tname + $_.Exception.Message); continue }");
                sb.AppendLine("    foreach ($c in $channels) {");
                sb.AppendLine("        $cid = [string]$c.id; $mt = [string]$c.membershipType");
                sb.AppendLine("        if ($mt -eq 'unknownFutureValue') { $mt = 'shared' }");
                if (wantSite)
                {
                    sb.AppendLine("        $sp = ''");
                    sb.AppendLine("        try { $fd = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + $cid + \"/filesFolder?`$select=webUrl\") -Headers $headers -Method Get -ErrorAction Stop; $sp = [string]$fd.webUrl } catch { $sp = '' }");
                }
                if (wantCounts)
                {
                    sb.AppendLine("        $oc = ''; $mc = ''; $gc = ''");
                    sb.AppendLine("        try { $cm = @(); $muri = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + $cid + \"/members\"; while ($muri) { $mr = Invoke-RestMethod -Uri $muri -Headers $headers -Method Get -ErrorAction Stop; $cm += @($mr.value); $muri = $mr.'@odata.nextLink' }; $oc = @($cm | Where-Object { @($_.roles) -contains 'owner' }).Count; $mc = $cm.Count } catch { }");
                    if (wantGuests)
                        sb.AppendLine("        try { $cs = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + $cid + \"?`$select=summary\") -Headers $headers -Method Get -ErrorAction Stop; $gc = $cs.summary.guestCount } catch { $gc = '' }");
                }
                sb.AppendLine("        $rows.Add([pscustomobject]@{");
                sb.AppendLine("            TeamName = $tname; ChannelName = $c.displayName;");
                sb.AppendLine("            ChannelType = $mt;");
                sb.AppendLine("            SharePointUrl = " + (wantSite ? "$sp" : "''") + "; CreatedDate = $c.createdDateTime;");
                sb.AppendLine("            OwnerCount = " + (chosen.Contains("OwnerCount") ? "$oc" : "''") + "; MemberCount = " + (chosen.Contains("MemberCount") ? "$mc" : "''") + "; GuestCount = " + (wantGuests ? "$gc" : "''"));
                sb.AppendLine("        })");
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

        // ============================================================ 3. PRIVATE & SHARED CHANNEL MEMBERSHIPS
        private static AuditSection BuildChannelMembershipsSection()
        {
            var section = new AuditSection(
                "teams-channel-memberships",
                "Channels (Private/Shared)",
                "Channels (private/shared) members export",
                "Audit members of private/shared channels via Graph (standard channels inherit team membership and are skipped).",
                "users",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsChannelMemberships.csv";

            var channel = new AuditOptionGroup("channel", "Channel", GroupMode.MultiCheck); channel.Columns = 2;
            channel.AddProp("TeamName", true);
            channel.AddProp("ChannelName", true);
            channel.AddProp("ChannelType", true);
            channel.AddProp("GuestCount", true);
            section.AddGroup(channel);

            var member = new AuditOptionGroup("member", "Member", GroupMode.MultiCheck); member.Columns = 2;
            member.AddProp("DisplayName", true);
            member.AddProp("UserPrincipalName", true);
            member.AddProp("Role", true);
            member.AddProp("UserType", true);
            section.AddGroup(member);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "channel", "member");
                if (chosen.Count == 0) chosen.Add("UserPrincipalName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                bool wantUpn = chosen.Contains("UserPrincipalName");
                bool wantType = chosen.Contains("UserType");
                bool wantName = chosen.Contains("DisplayName");
                bool wantGuests = chosen.Contains("GuestCount");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                if (wantUpn || wantType || wantName || wantGuests)
                {
                    sb.AppendLine("$userCache = @{}");
                    sb.AppendLine("function Get-CachedUser($uid) {");
                    sb.AppendLine("    if (-not $uid) { return $null }");
                    sb.AppendLine("    $k = [string]$uid");
                    sb.AppendLine("    if ($userCache.ContainsKey($k)) { return $userCache[$k] }");
                    sb.AppendLine("    try { $u = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/users/\" + $k + \"?`$select=displayName,userPrincipalName,userType\") -Headers $headers -Method Get -ErrorAction Stop; $userCache[$k] = $u; return $u } catch { return $null }");
                    sb.AppendLine("}");
                }
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; Write-Host (\"  [{0}/{1}] {2}\" -f $i, $groups.Count, $g.displayName)");
                sb.AppendLine("    $gid = [string]$g.id; $tname = [string]$g.displayName");
                sb.AppendLine("    $channels = @()");
                sb.AppendLine("    try {");
                sb.AppendLine("        $curi = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels\"");
                sb.AppendLine("        while ($curi) { $cr = Invoke-RestMethod -Uri $curi -Headers $headers -Method Get -ErrorAction Stop; $channels += @($cr.value); $curi = $cr.'@odata.nextLink' }");
                sb.AppendLine("    } catch { Write-Host (\"  WARNING: channels of '{0}' failed: \" -f $tname + $_.Exception.Message); continue }");
                sb.AppendLine("    foreach ($c in $channels) {");
                sb.AppendLine("        $mt = [string]$c.membershipType");
                sb.AppendLine("        if ($mt -eq 'unknownFutureValue') { $mt = 'shared' }");
                sb.AppendLine("        if ($mt -eq 'standard' -or $mt -eq '') { continue }");
                sb.AppendLine("        $cid = [string]$c.id; $cname = [string]$c.displayName");
                sb.AppendLine("        $cm = @()");
                sb.AppendLine("        try { $muri = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + $cid + \"/members\"; while ($muri) { $mr = Invoke-RestMethod -Uri $muri -Headers $headers -Method Get -ErrorAction Stop; $cm += @($mr.value); $muri = $mr.'@odata.nextLink' } } catch { Write-Host (\"  WARNING: members of channel '{0}' failed: \" -f $cname + $_.Exception.Message); continue }");
                sb.AppendLine("        $infos = @()");
                sb.AppendLine("        foreach ($m in $cm) {");
                sb.AppendLine("            $role = 'Member'; if (@($m.roles) -contains 'owner') { $role = 'Owner' }");
                if (wantUpn || wantType || wantName || wantGuests)
                {
                    sb.AppendLine("            $uu = Get-CachedUser $m.userId");
                    sb.AppendLine("            $dn = [string]$m.displayName; if ($uu -and $uu.displayName) { $dn = [string]$uu.displayName }");
                    sb.AppendLine("            $upn = [string]$m.email; if ($uu -and $uu.userPrincipalName) { $upn = [string]$uu.userPrincipalName }");
                    sb.AppendLine("            $ut = ''; if ($uu -and $uu.userType) { $ut = [string]$uu.userType }");
                }
                else
                {
                    sb.AppendLine("            $dn = [string]$m.displayName; $upn = [string]$m.email; $ut = ''");
                }
                sb.AppendLine("            $infos += [pscustomobject]@{ Role = $role; DisplayName = $dn; UserPrincipalName = $upn; UserType = $ut }");
                sb.AppendLine("        }");
                if (wantGuests)
                    sb.AppendLine("        $gcount = @($infos | Where-Object { $_.UserType -eq 'Guest' }).Count");
                else
                    sb.AppendLine("        $gcount = ''");
                sb.AppendLine("        foreach ($x in $infos) {");
                sb.AppendLine("            $rows.Add([pscustomobject]@{");
                sb.AppendLine("                TeamName = $tname; ChannelName = $cname; ChannelType = $mt; GuestCount = $gcount;");
                sb.AppendLine("                DisplayName = $x.DisplayName; UserPrincipalName = $x.UserPrincipalName; Role = $x.Role; UserType = $x.UserType");
                sb.AppendLine("            })");
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

        // ============================================================ 4. MEMBERS & OWNERS
        private static AuditSection BuildMembersSection()
        {
            var section = new AuditSection(
                "teams-members",
                "Users",
                "Users export",
                "Audit owners and members of every team via Graph: one row per user with guest flag and external domain.",
                "user",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsMembers.csv";

            var team = new AuditOptionGroup("team", "Team", GroupMode.MultiCheck); team.Columns = 1;
            team.AddProp("TeamName", true);
            team.AddProp("Role", true);
            section.AddGroup(team);

            var user = new AuditOptionGroup("user", "User", GroupMode.MultiCheck); user.Columns = 2;
            user.AddProp("DisplayName", true);
            user.AddProp("UserPrincipalName", true);
            user.AddProp("AccountEnabled", true);
            user.AddProp("UserType", true);
            user.AddProp("IsGuest", true);
            user.AddProp("ExternalDomain", true);
            section.AddGroup(user);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "team", "user");
                if (chosen.Count == 0) chosen.Add("UserPrincipalName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; Write-Host (\"  [{0}/{1}] {2}\" -f $i, $groups.Count, $g.displayName)");
                sb.AppendLine("    $gid = [string]$g.id; $tname = [string]$g.displayName");
                sb.AppendLine("    $owners = @(); $members = @()");
                sb.AppendLine("    try { $ouri = \"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/owners?`$select=id,displayName,userPrincipalName,accountEnabled,userType\"; while ($ouri) { $orr = Invoke-RestMethod -Uri $ouri -Headers $headers -Method Get -ErrorAction Stop; $owners += @($orr.value); $ouri = $orr.'@odata.nextLink' } } catch { Write-Host (\"  WARNING: owners of '{0}' failed: \" -f $tname + $_.Exception.Message) }");
                sb.AppendLine("    try { $muri = \"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/members?`$select=id,displayName,userPrincipalName,accountEnabled,userType\"; while ($muri) { $mrr = Invoke-RestMethod -Uri $muri -Headers $headers -Method Get -ErrorAction Stop; $members += @($mrr.value); $muri = $mrr.'@odata.nextLink' } } catch { Write-Host (\"  WARNING: members of '{0}' failed: \" -f $tname + $_.Exception.Message) }");
                sb.AppendLine("    $seen = @{}");
                sb.AppendLine("    foreach ($o in $owners) {");
                sb.AppendLine("        $oid = [string]$o.id; if ($oid) { $seen[$oid] = $true }");
                sb.AppendLine("        $ut = [string]$o.userType; $guest = 'No'; if ($ut -eq 'Guest') { $guest = 'Yes' }");
                sb.AppendLine("        $dom = ''; $u = [string]$o.userPrincipalName; if ($guest -eq 'Yes' -and $u -match '@') { $dom = $u.Substring($u.LastIndexOf('@') + 1) }");
                sb.AppendLine("        $rows.Add([pscustomobject]@{ TeamName = $tname; Role = 'Owner'; DisplayName = $o.displayName; UserPrincipalName = $o.userPrincipalName; AccountEnabled = $o.accountEnabled; UserType = $ut; IsGuest = $guest; ExternalDomain = $dom })");
                sb.AppendLine("    }");
                sb.AppendLine("    foreach ($m in $members) {");
                sb.AppendLine("        $mid = [string]$m.id; if ($mid -and $seen.ContainsKey($mid)) { continue }");
                sb.AppendLine("        $ut = [string]$m.userType; $guest = 'No'; if ($ut -eq 'Guest') { $guest = 'Yes' }");
                sb.AppendLine("        $dom = ''; $u = [string]$m.userPrincipalName; if ($guest -eq 'Yes' -and $u -match '@') { $dom = $u.Substring($u.LastIndexOf('@') + 1) }");
                sb.AppendLine("        $rows.Add([pscustomobject]@{ TeamName = $tname; Role = 'Member'; DisplayName = $m.displayName; UserPrincipalName = $m.userPrincipalName; AccountEnabled = $m.accountEnabled; UserType = $ut; IsGuest = $guest; ExternalDomain = $dom })");
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

        // ============================================================ 5. TEAM APPS
        private static AuditSection BuildAppsSection()
        {
            var section = new AuditSection(
                "teams-apps",
                "Apps",
                "Apps export",
                "Audit apps installed in each team via Graph: one row per team x app with publisher and version.",
                "app",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsApps.csv";

            var team = new AuditOptionGroup("team", "Team", GroupMode.MultiCheck); team.Columns = 1;
            team.AddProp("TeamName", true);
            section.AddGroup(team);

            var app = new AuditOptionGroup("app", "App", GroupMode.MultiCheck); app.Columns = 2;
            app.AddProp("AppName", true);
            app.AddProp("Publisher", true);
            app.AddProp("AppId", true);
            app.AddProp("Version", true);
            app.AddProp("DistributionMethod", true);
            section.AddGroup(app);

            // Filter only (never a CSV column): uncheck Store to keep org-deployed
            // (organization) + sideloaded apps and drop built-in defaults.
            var dist = new AuditOptionGroup("distribution", "Distribution filter", GroupMode.MultiCheck); dist.Columns = 3;
            dist.Hint = "Every row below IS installed in its team - use this to narrow to the apps you deployed (Organization = your catalog apps).";
            dist.Add(new AuditOption("organization", "Organization", "organization", true));
            dist.Add(new AuditOption("store", "Store", "store", true));
            dist.Add(new AuditOption("sideloaded", "Sideloaded", "sideloaded", true));
            section.AddGroup(dist);

            // Publisher filter (never a CSV column, unchecked by default):
            // hides everything published by Microsoft - including apps you may
            // use deliberately (Planner, Approvals, Whiteboard...).
            var pubf = new AuditOptionGroup("publisher", "Publisher filter", GroupMode.MultiCheck); pubf.Columns = 1;
            pubf.Hint = "Hides ALL Microsoft-published apps, deliberate or default. Leave unchecked to see everything.";
            pubf.Add(new AuditOption("hidems", "Hide Microsoft-published apps", "hidems", false));
            section.AddGroup(pubf);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "team", "app");
                if (chosen.Count == 0) chosen.Add("AppName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                // Filter values must NOT enter the Select-Object list: read separately.
                var distWant = sel.Selected("distribution");
                if (distWant.Count == 0) distWant = new List<string> { "organization", "store", "sideloaded" };
                bool hideMs = sel.IsSelected("publisher", "hidems");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                sb.AppendLine("$wantDist = @('" + string.Join("','", distWant.ToArray()) + "')");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; Write-Host (\"  [{0}/{1}] {2}\" -f $i, $groups.Count, $g.displayName)");
                sb.AppendLine("    $gid = [string]$g.id; $tname = [string]$g.displayName");
                sb.AppendLine("    $apps = @()");
                sb.AppendLine("    try {");
                sb.AppendLine("        $auri = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/installedApps?`$expand=teamsApp,teamsAppDefinition\"");
                sb.AppendLine("        while ($auri) { $ar = Invoke-RestMethod -Uri $auri -Headers $headers -Method Get -ErrorAction Stop; $apps += @($ar.value); $auri = $ar.'@odata.nextLink' }");
                sb.AppendLine("    } catch { Write-Host (\"  WARNING: apps of '{0}' failed: \" -f $tname + $_.Exception.Message); continue }");
                sb.AppendLine("    foreach ($a in $apps) {");
                sb.AppendLine("        $an = ''; if ($a.teamsApp -and $a.teamsApp.displayName) { $an = [string]$a.teamsApp.displayName }");
                sb.AppendLine("        $aid = ''; if ($a.teamsApp -and $a.teamsApp.id) { $aid = [string]$a.teamsApp.id }");
                sb.AppendLine("        $pub = ''; if ($a.teamsAppDefinition -and $a.teamsAppDefinition.publisher -and $a.teamsAppDefinition.publisher.displayName) { $pub = [string]$a.teamsAppDefinition.publisher.displayName }");
                sb.AppendLine("        $ver = ''; if ($a.teamsAppDefinition -and $a.teamsAppDefinition.version) { $ver = [string]$a.teamsAppDefinition.version }");
                sb.AppendLine("        $dm = ''; if ($a.teamsApp -and $a.teamsApp.distributionMethod) { $dm = [string]$a.teamsApp.distributionMethod }");
                sb.AppendLine("        if ($dm -ne '' -and $wantDist -notcontains $dm) { continue }");
                if (hideMs)
                    sb.AppendLine("        if ($pub -like 'Microsoft*') { continue }");
                sb.AppendLine("        $rows.Add([pscustomobject]@{ TeamName = $tname; AppName = $an; Publisher = $pub; AppId = $aid; Version = $ver; DistributionMethod = $dm })");
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

        // ============================================================ 6. TEAM TABS
        private static AuditSection BuildTabsSection()
        {
            var section = new AuditSection(
                "teams-tabs",
                "Tabs",
                "Tabs export",
                "Audit pinned tabs per channel via Graph: one row per team x channel x tab. The native Files tab is not returned by the API.",
                "tab",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsTabs.csv";

            var team = new AuditOptionGroup("team", "Team & channel", GroupMode.MultiCheck); team.Columns = 2;
            team.AddProp("TeamName", true);
            team.AddProp("ChannelName", true);
            section.AddGroup(team);

            var tab = new AuditOptionGroup("tab", "Tab", GroupMode.MultiCheck); tab.Columns = 2;
            tab.AddProp("TabName", true);
            tab.AddProp("TabType", true);
            section.AddGroup(tab);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "team", "tab");
                if (chosen.Count == 0) chosen.Add("TabName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; Write-Host (\"  [{0}/{1}] {2}\" -f $i, $groups.Count, $g.displayName)");
                sb.AppendLine("    $gid = [string]$g.id; $tname = [string]$g.displayName");
                sb.AppendLine("    $channels = @()");
                sb.AppendLine("    try {");
                sb.AppendLine("        $curi = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels\"");
                sb.AppendLine("        while ($curi) { $cr = Invoke-RestMethod -Uri $curi -Headers $headers -Method Get -ErrorAction Stop; $channels += @($cr.value); $curi = $cr.'@odata.nextLink' }");
                sb.AppendLine("    } catch { Write-Host (\"  WARNING: channels of '{0}' failed: \" -f $tname + $_.Exception.Message); continue }");
                sb.AppendLine("    foreach ($c in $channels) {");
                sb.AppendLine("        $cid = [string]$c.id; $cname = [string]$c.displayName");
                sb.AppendLine("        $tabs = @()");
                sb.AppendLine("        try {");
                sb.AppendLine("            $turi = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + $cid + \"/tabs?`$expand=teamsApp\"");
                sb.AppendLine("            while ($turi) { $tr = Invoke-RestMethod -Uri $turi -Headers $headers -Method Get -ErrorAction Stop; $tabs += @($tr.value); $turi = $tr.'@odata.nextLink' }");
                sb.AppendLine("        } catch {");
                sb.AppendLine("            try {");
                sb.AppendLine("                $turi = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + $cid + \"/tabs\"");
                sb.AppendLine("                while ($turi) { $tr = Invoke-RestMethod -Uri $turi -Headers $headers -Method Get -ErrorAction Stop; $tabs += @($tr.value); $turi = $tr.'@odata.nextLink' }");
                sb.AppendLine("            } catch { Write-Host (\"  WARNING: tabs of channel '{0}' failed: \" -f $cname + $_.Exception.Message); continue }");
                sb.AppendLine("        }");
                sb.AppendLine("        foreach ($t in $tabs) {");
                sb.AppendLine("            $ta = ''; if ($t.teamsApp -and $t.teamsApp.displayName) { $ta = [string]$t.teamsApp.displayName }");
                sb.AppendLine("            $rows.Add([pscustomobject]@{ TeamName = $tname; ChannelName = $cname; TabName = $t.displayName; TabType = $ta })");
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

        // ============================================================ 7. SHAREPOINT SITES
        private static AuditSection BuildSitesSection()
        {
            var section = new AuditSection(
                "teams-sites",
                "SharePoint Sites",
                "SharePoint sites export",
                "Audit the SharePoint root site of each team via Graph: one row per team (storage live; template + file count from the SharePoint usage report, up to 48h stale).",
                "site",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsSites.csv";

            var team = new AuditOptionGroup("team", "Team", GroupMode.MultiCheck); team.Columns = 1;
            team.AddProp("TeamName", true);
            section.AddGroup(team);

            var site = new AuditOptionGroup("site", "SharePoint site", GroupMode.MultiCheck); site.Columns = 2;
            site.AddProp("SiteUrl", true);
            site.AddProp("Template", true);
            site.AddProp("StorageUsedGB", true);
            site.AddProp("StorageQuotaGB", true);
            site.AddProp("FileCount", true);
            site.AddProp("LastContentModified", true);
            section.AddGroup(site);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "team", "site");
                if (chosen.Count == 0) chosen.Add("SiteUrl");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                // Template + FileCount come from the tenant usage report (fetched
                // once); everything else is live per-team data.
                bool wantReport = chosen.Contains("Template") || chosen.Contains("FileCount");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                if (wantReport)
                {
                    sb.AppendLine("$repById = @{}; $repByUrl = @{}");
                    sb.AppendLine("try {");
                    sb.AppendLine("    $repRaw = Invoke-WebRequest -Uri \"https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')\" -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop");
                    sb.AppendLine("    $repRows = @(($repRaw.Content -split \"`n\" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) | ConvertFrom-Csv)");
                    sb.AppendLine("    foreach ($rr in $repRows) {");
                    sb.AppendLine("        $ik = ([string]$rr.'Site Id').Trim().ToLower(); if ($ik -and -not $repById.ContainsKey($ik)) { $repById[$ik] = $rr }");
                    sb.AppendLine("        $uk = ([string]$rr.'Site URL').Trim().TrimEnd('/').ToLower(); if ($uk -and -not $repByUrl.ContainsKey($uk)) { $repByUrl[$uk] = $rr }");
                    sb.AppendLine("    }");
                    sb.AppendLine("    Write-Host (\"Loaded SharePoint usage report ({0} rows).\" -f $repRows.Count)");
                    sb.AppendLine("} catch { Write-Host (\"WARNING: SharePoint usage report failed (consent for Reports.Read.All?) - Template/FileCount will be blank: \" + $_.Exception.Message) }");
                }
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; if (($i % 20) -eq 0) { Write-Host (\"  ... {0}/{1}\" -f $i, $groups.Count) }");
                sb.AppendLine("    $gid = [string]$g.id; $tname = [string]$g.displayName");
                sb.AppendLine("    $st = $null");
                sb.AppendLine("    try { $st = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/sites/root?`$select=id,displayName,webUrl,lastModifiedDateTime\") -Headers $headers -Method Get -ErrorAction Stop } catch { Write-Host (\"  WARNING: site of '{0}' failed: \" -f $tname + $_.Exception.Message) }");
                sb.AppendLine("    $su = ''; $lm = ''; $sugb = ''; $sqgb = ''; $tpl = ''; $fc = ''");
                sb.AppendLine("    if ($st) {");
                sb.AppendLine("        $su = [string]$st.webUrl; $lm = [string]$st.lastModifiedDateTime");
                sb.AppendLine("        try { $dr = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/groups/\" + $gid + \"/drive?`$select=quota\") -Headers $headers -Method Get -ErrorAction Stop } catch { Write-Host (\"  WARNING: drive quota of '{0}' failed: \" -f $tname + $_.Exception.Message) }");
                sb.AppendLine("        if ($dr -and $dr.quota) {");
                sb.AppendLine("            try { if ($null -ne $dr.quota.used) { $sugb = [math]::Round([double]$dr.quota.used / 1GB, 2) } } catch { }");
                sb.AppendLine("            try { if ($null -ne $dr.quota.total) { $sqgb = [math]::Round([double]$dr.quota.total / 1GB, 2) } } catch { }");
                sb.AppendLine("        }");
                if (wantReport)
                {
                    sb.AppendLine("        $rm = $null");
                    sb.AppendLine("        $ukey = ([string]$st.webUrl).Trim().TrimEnd('/').ToLower()");
                    sb.AppendLine("        if ($ukey -and $repByUrl.ContainsKey($ukey)) { $rm = $repByUrl[$ukey] }");
                    sb.AppendLine("        else { $tRaw = [string]$st.id; if ($tRaw -match ',') { $ik2 = ($tRaw.Split(',')[1]).Trim().ToLower(); if ($repById.ContainsKey($ik2)) { $rm = $repById[$ik2] } } }");
                    sb.AppendLine("        if ($rm) { $tpl = [string]$rm.'Root Web Template'; $fc = [string]$rm.'File Count' }");
                }
                sb.AppendLine("    }");
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        TeamName = $tname; SiteUrl = $su; Template = $tpl;");
                sb.AppendLine("        StorageUsedGB = $sugb; StorageQuotaGB = $sqgb; FileCount = $fc; LastContentModified = $lm");
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

        // ============================================================ 8. TEAM ACTIVITY
        private static AuditSection BuildActivitySection()
        {
            var section = new AuditSection(
                "teams-activity",
                "Activity",
                "Activity export",
                "Audit per-team activity via Graph: latest channel message (top-level thread + replies). One row per team.",
                "activity",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsActivity.csv";

            var team = new AuditOptionGroup("team", "Team", GroupMode.MultiCheck); team.Columns = 1;
            team.AddProp("TeamName", true);
            section.AddGroup(team);

            var activity = new AuditOptionGroup("activity", "Latest activity (per-channel message lookups)", GroupMode.MultiCheck); activity.Columns = 2;
            activity.Hint = "Reads the latest thread of every channel (one query per channel) - slower on large tenants. LastChannelMessage shows 'Channel: message snippet'.";
            activity.AddProp("LastChannelMessage", true);
            activity.AddProp("LastActivityDate", true);
            section.AddGroup(activity);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "team", "activity");
                if (chosen.Count == 0) chosen.Add("TeamName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                bool wantMsg = chosen.Contains("LastChannelMessage") || chosen.Contains("LastActivityDate");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                // Channel messages come back newest-thread-first, so $top=1 returns
                // the most recently active thread; $expand=replies covers answers.
                sb.AppendLine("function Get-MsgText { param($m)");
                sb.AppendLine("    $t = ''; try { $t = [string]$m.body.content } catch { }");
                sb.AppendLine("    $t = $t -replace '<[^>]+>',' '");
                sb.AppendLine("    $t = ($t -replace '\\s+',' ').Trim()");
                sb.AppendLine("    if ($t.Length -gt 250) { $t = $t.Substring(0,250) + '...' }");
                sb.AppendLine("    return $t");
                sb.AppendLine("}");
                if (wantMsg) sb.AppendLine("$msgFail = 0");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $i++; Write-Host (\"  [{0}/{1}] {2}\" -f $i, $groups.Count, $g.displayName)");
                sb.AppendLine("    $gid = [string]$g.id; $tname = [string]$g.displayName");
                sb.AppendLine("    $teamBest = ''; $teamTxt = ''");
                if (wantMsg)
                {
                    sb.AppendLine("    try {");
                    sb.AppendLine("        $channels = @(); $curi = \"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels\"");
                    sb.AppendLine("        while ($curi) { $cr = Invoke-RestMethod -Uri $curi -Headers $headers -Method Get -ErrorAction Stop; $channels += @($cr.value); $curi = $cr.'@odata.nextLink' }");
                    sb.AppendLine("        foreach ($c in $channels) {");
                    sb.AppendLine("            try {");
                    sb.AppendLine("                $cid = [string]$c.id; $cname = [string]$c.displayName");
                    sb.AppendLine("                $m0 = $null");
                    sb.AppendLine("                $msgr = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/teams/\" + $gid + \"/channels/\" + $cid + \"/messages?`$top=1&`$expand=replies\") -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("                $m0 = @($msgr.value)[0]");
                    sb.AppendLine("                if ($m0) {");
                    sb.AppendLine("                    $cand = [string]$m0.lastModifiedDateTime; if (-not $cand) { $cand = [string]$m0.createdDateTime }");
                    sb.AppendLine("                    $ctxt = Get-MsgText $m0");
                    sb.AppendLine("                    foreach ($rp in @($m0.replies)) {");
                    sb.AppendLine("                        $rdt = [string]$rp.createdDateTime");
                    sb.AppendLine("                        if ($rdt -and $rdt -gt $cand) { $cand = $rdt; $ctxt = Get-MsgText $rp }");
                    sb.AppendLine("                    }");
                    sb.AppendLine("                    if ($cand -and $cand -gt $teamBest) { $teamBest = $cand; $teamTxt = $cname + ': ' + $ctxt }");
                    sb.AppendLine("                }");
                    sb.AppendLine("            } catch { $msgFail++ }");
                    sb.AppendLine("        }");
                    sb.AppendLine("    } catch { Write-Host (\"  WARNING: channels of '{0}' failed: \" -f $tname + $_.Exception.Message) }");
                }
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        TeamName = $tname; LastChannelMessage = $teamTxt; LastActivityDate = $teamBest");
                sb.AppendLine("    })");
                sb.AppendLine("}");
                if (wantMsg) sb.AppendLine("if ($msgFail -gt 0) { Write-Host (\"WARNING: {0} channel message querie(s) failed (consent for ChannelMessage.Read.All?) - affected teams show no message.\" -f $msgFail) }");
                sb.AppendLine("$rows = $rows | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 9. SENSITIVITY & COMPLIANCE
        // NOTE: RetentionPolicy is deliberately absent - Graph v1.0 exposes no
        // per-team retention property (retention lives in the Purview compliance
        // center, which needs a different connection than these Graph sections).
        private static AuditSection BuildComplianceSection()
        {
            var section = new AuditSection(
                "teams-compliance",
                "Sensitivity & Compliance",
                "Sensitivity & compliance export",
                "Audit sensitivity and lifecycle signals per team via Graph: one row per team, no per-team lookups.",
                "shield",
                AuditScope.Graph);
            section.Category = "Teams";
            section.Product = "Teams";
            section.DefaultFileName = "TeamsCompliance.csv";

            var team = new AuditOptionGroup("team", "Team", GroupMode.MultiCheck); team.Columns = 1;
            team.AddProp("TeamName", true);
            section.AddGroup(team);

            var compliance = new AuditOptionGroup("compliance", "Sensitivity & lifecycle", GroupMode.MultiCheck); compliance.Columns = 2;
            compliance.Hint = "If your tenant assigns no labels, classifications or expiry, these columns come back empty and Smart mode hides them - uncheck Smart mode on a test run to confirm.";
            compliance.AddProp("SensitivityLabel", true);
            compliance.AddProp("Classification", true);
            compliance.AddProp("ExpirationPolicy", true);
            section.AddGroup(compliance);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "team", "compliance");
                if (chosen.Count == 0) chosen.Add("TeamName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Teams (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine(FetchTeamsGroups);
                sb.AppendLine("Write-Host (\"Retrieved {0} team(s).\" -f $groups.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("foreach ($g in $groups) {");
                sb.AppendLine("    $lbl = ((@($g.assignedLabels) | ForEach-Object { if ($_.displayName) { [string]$_.displayName } elseif ($_.labelId) { [string]$_.labelId } else { [string]$_ } }) -join ',')");
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        TeamName = $g.displayName; SensitivityLabel = $lbl;");
                sb.AppendLine("        Classification = $g.classification; ExpirationPolicy = $g.expirationDateTime");
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
    }
}
