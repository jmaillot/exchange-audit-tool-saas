using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    internal static class SectionsMailboxTypes
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildSharedMailboxSection());
            AuditRegistry.Register(BuildResourceSection(true));
            AuditRegistry.Register(BuildResourceSection(false));
        }

        private static readonly List<string> SharedRecipientRefs = new List<string>(new string[]
        {
            "ForwardingAddress", "GrantSendOnBehalfTo",
            "AcceptMessagesOnlyFromSendersOrMembers", "RejectMessagesFromSendersOrMembers"
        });

        private static readonly List<string> SharedMultiValued = new List<string>(new string[]
        {
            "InPlaceHolds", "EmailAddresses"
        });

        private static readonly string[] SharedPropGroups =
        {
            "identity", "type", "visibility", "quotas", "retention", "mailflow", "misc"
        };

        private static readonly List<string> CalProcMultiValued = new List<string>(new string[]
        {
            "ResourceDelegates", "BookInPolicy", "AllBookInPolicy", "RequestOutOfPolicy", "RequestInPolicy"
        });

        private static AuditSection BuildSharedMailboxSection()
        {
            var section = new AuditSection(
                "shared-mailboxes",
                "Shared mailboxes",
                "Shared mailbox export",
                "Audit shared mailboxes - permissions-focused, since access relies entirely on delegation.",
                "mailbox",
                AuditScope.Both);
            section.Category = "Mailboxes";
            section.DefaultFileName = "SharedMailboxes.csv";

            var identity = new AuditOptionGroup("identity", "Identity / addressing", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("PrimarySmtpAddress", true);
            identity.AddProp("EmailAddresses", true);
            identity.AddProp("ExchangeGuid", false);
            section.AddGroup(identity);

            var type = new AuditOptionGroup("type", "Type", GroupMode.MultiCheck); type.Columns = 2;
            type.AddProp("RecipientTypeDetails", true);
            section.AddGroup(type);

            var visibility = new AuditOptionGroup("visibility", "Visibility", GroupMode.MultiCheck); visibility.Columns = 2;
            visibility.AddProp("HiddenFromAddressListsEnabled", false);
            section.AddGroup(visibility);

            var quotas = new AuditOptionGroup("quotas", "Quotas", GroupMode.MultiCheck); quotas.Columns = 2;
            quotas.AddProp("ProhibitSendQuota", false);
            quotas.AddProp("ProhibitSendReceiveQuota", false);
            quotas.AddProp("IssueWarningQuota", false);
            quotas.AddProp("UseDatabaseQuotaDefaults", false);
            section.AddGroup(quotas);

            var retention = new AuditOptionGroup("retention", "Retention / compliance", GroupMode.MultiCheck); retention.Columns = 2;
            retention.AddProp("RetentionPolicy", true);
            retention.AddProp("LitigationHoldEnabled", true);
            retention.AddProp("SingleItemRecoveryEnabled", false);
            retention.AddProp("AuditEnabled", false);
            retention.AddProp("InPlaceHolds", false);
            section.AddGroup(retention);

            var mailflow = new AuditOptionGroup("mailflow", "Mail flow", GroupMode.MultiCheck); mailflow.Columns = 2;
            mailflow.AddProp("ForwardingAddress", true);
            mailflow.AddProp("ForwardingSmtpAddress", true);
            mailflow.AddProp("DeliverToMailboxAndForward", false);
            mailflow.AddProp("AcceptMessagesOnlyFromSendersOrMembers", false);
            mailflow.AddProp("RejectMessagesFromSendersOrMembers", false);
            mailflow.AddProp("MessageCopyForSentAsEnabled", false);
            mailflow.AddProp("MessageCopyForSendOnBehalfEnabled", false);
            section.AddGroup(mailflow);

            var misc = new AuditOptionGroup("misc", "Misc", GroupMode.MultiCheck); misc.Columns = 2;
            misc.AddProp("WhenCreated", false);
            misc.AddProp("Database", false);
            misc.AddProp("ServerName", false);
            section.AddGroup(misc);

            var perms = new AuditOptionGroup("permissions", "Permissions (critical - who can use the mailbox)", GroupMode.MultiCheck);
            perms.Hint = "FullAccess & SendAs are bulk-indexed (one pass each); AutoMapping reads msExchDelegateListLink (may be blank on EXO).";
            perms.Columns = 2;
            perms.Add(new AuditOption("fullaccess", "Full Access (Get-MailboxPermission)", "fullaccess", true));
            perms.Add(new AuditOption("sendas", "Send As (EXO: RecipientPermission / on-prem: ADPermission)", "sendas", true));
            perms.Add(new AuditOption("sendonbehalf", "Send on Behalf (GrantSendOnBehalfTo)", "sendonbehalf", true));
            perms.Add(new AuditOption("automapping", "AutoMapping delegates (msExchDelegateListLink)", "automapping", false));
            section.AddGroup(perms);

            var folder = new AuditOptionGroup("folderperms", "Folder permissions (per-mailbox, slower)", GroupMode.MultiCheck).MarkSlow();
            folder.Hint = "Get-MailboxFolderPermission on the chosen folder(s); non-default entries only.";
            folder.Columns = 2;
            folder.Add(new AuditOption("calendar", "Calendar", "Calendar", false));
            folder.Add(new AuditOption("inbox", "Inbox", "Inbox", false));
            folder.Add(new AuditOption("sentitems", "Sent Items", "SentItems", false));
            folder.Add(new AuditOption("contacts", "Contacts", "Contacts", false));
            folder.Add(new AuditOption("tasks", "Tasks", "Tasks", false));
            section.AddGroup(folder);

            var extra = new AuditOptionGroup("complementary", "Complementary data", GroupMode.MultiCheck);
            extra.Columns = 1;
            extra.Add(new AuditOption("mailboxsize", "Mailbox size in MB (Get-MailboxStatistics)", "mailboxsize", false).MarkSlow());
            extra.Add(new AuditOption("mailboxitemcount", "Item count (Get-MailboxStatistics)", "mailboxitemcount", false).MarkSlow());
            extra.Add(new AuditOption("archivesize", "Archive size in MB (per-mailbox, slower)", "archivesize", false).MarkSlow());
            extra.Add(new AuditOption("accountstatus", "Account activated / disabled (Get-User)", "accountstatus", false));
            extra.Add(new AuditOption("regional", "Regional config (Language, TimeZone) - per-mailbox, slower", "regional", false).MarkSlow());
            section.AddGroup(extra);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");

                var chosen = new List<string>();
                foreach (string gk in SharedPropGroups)
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("DisplayName");

                bool needResolver = false;
                var exprs = new List<string>();
                foreach (string v in chosen)
                {
                    if (v == "EmailAddresses")
                        exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                    else if (SharedRecipientRefs.Contains(v)) { needResolver = true; exprs.Add("@{Name='" + v + "';Expression={ Resolve-Recip $_." + v + " }}"); }
                    else if (SharedMultiValued.Contains(v))
                        exprs.Add("@{Name='" + v + "';Expression={($_." + v + " | ForEach-Object { [string]$_ }) -join ','}}");
                    else exprs.Add(v);
                }
                string selectList = string.Join(", ", exprs.ToArray());

                bool fa = sel.IsSelected("permissions", "fullaccess");
                bool sa = sel.IsSelected("permissions", "sendas");
                bool sob = sel.IsSelected("permissions", "sendonbehalf");
                bool automap = sel.IsSelected("permissions", "automapping");
                if (sob) needResolver = true;
                List<string> folders = sel.Selected("folderperms");
                bool needSize = sel.IsSelected("complementary", "mailboxsize");
                bool needCount = sel.IsSelected("complementary", "mailboxitemcount");
                bool needArchive = sel.IsSelected("complementary", "archivesize");
                bool acctStatus = sel.IsSelected("complementary", "accountstatus");
                bool regional = sel.IsSelected("complementary", "regional");

                bool online = ConnectionSettings.IsOnline;
                string getMbx = "Get-Mailbox";
                string getRecip = online ? "Get-EXORecipient" : "Get-Recipient";
                string getStats = online ? "Get-EXOMailboxStatistics" : "Get-MailboxStatistics";
                string getMbxPerm = "Get-MailboxPermission";
                string getRecPerm = online ? "Get-EXORecipientPermission" : "Get-RecipientPermission";

                bool perRow = fa || sa || sob || automap || folders.Count > 0 || needSize || needCount || needArchive || acctStatus || regional;

                var sb = new StringBuilder();
                if (needSize) PsScriptHelpers.EmitSizeHelper(sb, getStats);
                if (needCount) PsScriptHelpers.EmitCountHelper(sb, getStats);
                if (needArchive) PsScriptHelpers.EmitArchiveMBHelper(sb, getStats);
                if (needResolver || automap) PsScriptHelpers.EmitResolver(sb, getRecip, false);

                sb.AppendLine("Write-Host 'Querying shared mailboxes...'");
                sb.AppendLine("$mbx = @(" + getMbx + " -ResultSize " + resultSize + " -RecipientTypeDetails SharedMailbox)");
                sb.AppendLine("Write-Host (\"Retrieved {0} shared mailbox(es).\" -f $mbx.Count)");
                sb.AppendLine();

                if (acctStatus)
                {
                    sb.AppendLine("Write-Host 'Indexing Get-User (bulk)...'");
                    sb.AppendLine("$userIndex = @{}");
                    sb.AppendLine("Get-User -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object { $userIndex[[string]$_.Guid] = $_ }");
                    sb.AppendLine();
                }
                // Bulk-index FullAccess / SendAs in one pipeline pass each (was per-mailbox, now O(1) lookups).
                if (fa)
                {
                    sb.AppendLine("Write-Host 'Indexing Full Access (bulk)...'");
                    sb.AppendLine("$faIndex = @{}");
                    sb.AppendLine("$mbx | " + getMbxPerm + " -ErrorAction SilentlyContinue |");
                    sb.AppendLine("    Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\\*' } |");
                    sb.AppendLine("    ForEach-Object { $k=[string]$_.Identity; if ($faIndex.ContainsKey($k)) { $faIndex[$k]=$faIndex[$k]+','+[string]$_.User } else { $faIndex[$k]=[string]$_.User } }");
                    sb.AppendLine();
                }
                if (sa)
                {
                    sb.AppendLine("Write-Host 'Indexing Send As (bulk)...'");
                    sb.AppendLine("$saIndex = @{}");
                    if (online)
                    {
                        // Exchange Online: Send As is exposed via (EXO)RecipientPermission.
                        sb.AppendLine("$mbx | " + getRecPerm + " -ErrorAction SilentlyContinue |");
                        sb.AppendLine("    Where-Object { $_.AccessRights -contains 'SendAs' -and $_.Trustee -notlike 'NT AUTHORITY\\*' } |");
                        sb.AppendLine("    ForEach-Object { $k=[string]$_.Identity; if ($saIndex.ContainsKey($k)) { $saIndex[$k]=$saIndex[$k]+','+[string]$_.Trustee } else { $saIndex[$k]=[string]$_.Trustee } }");
                    }
                    else
                    {
                        // On-premises: Get-RecipientPermission does not exist. Send As lives in
                        // Active Directory as the 'Send-As' extended right (Get-ADPermission).
                        sb.AppendLine("foreach ($__m in $mbx) {");
                        sb.AppendLine("    $k = [string]$__m.Identity");
                        sb.AppendLine("    $__t = Get-ADPermission -Identity $__m.Identity -ErrorAction SilentlyContinue |");
                        sb.AppendLine("        Where-Object { ($_.ExtendedRights -like '*Send-As*') -and (-not $_.IsInherited) -and (-not $_.Deny) -and ($_.User -notlike 'NT AUTHORITY\\*') } |");
                        sb.AppendLine("        ForEach-Object { [string]$_.User }");
                        sb.AppendLine("    if ($__t) { $saIndex[$k] = (@($__t) -join ',') }");
                        sb.AppendLine("}");
                    }
                    sb.AppendLine();
                }

                if (!perRow)
                {
                    sb.AppendLine("$rows = $mbx | Select-Object " + selectList);
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($m in $mbx) {");
                    sb.AppendLine("    $obj = $m | Select-Object " + selectList);
                    sb.AppendLine("    $key = [string]$m.Identity");
                    if (fa)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName FullAccess -NotePropertyValue ($(if ($faIndex.ContainsKey($key)) { $faIndex[$key] } else { '' })) -Force");
                    if (sa)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName SendAs -NotePropertyValue ($(if ($saIndex.ContainsKey($key)) { $saIndex[$key] } else { '' })) -Force");
                    if (sob)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName SendOnBehalf -NotePropertyValue (Resolve-Recip $m.GrantSendOnBehalfTo) -Force");
                    if (automap)
                    {
                        sb.AppendLine("    $am = ''");
                        sb.AppendLine("    try { if ($m.PSObject.Properties['msExchDelegateListLink'] -and $m.msExchDelegateListLink) { $am = Resolve-Recip $m.msExchDelegateListLink } } catch { }");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName AutoMappingDelegates -NotePropertyValue $am -Force");
                    }
                    foreach (string f in folders)
                    {
                        string colName = "FolderPerm_" + f;
                        sb.AppendLine("    $fp = ''");
                        sb.AppendLine("    try {");
                        sb.AppendLine("        $fp = @(Get-MailboxFolderPermission -Identity ($m.PrimarySmtpAddress.ToString() + ':\\" + f + "') -ErrorAction Stop |");
                        sb.AppendLine("            Where-Object { $_.User.DisplayName -ne 'Default' -and $_.User.DisplayName -ne 'Anonymous' } |");
                        sb.AppendLine("            ForEach-Object { [string]$_.User.DisplayName + '=' + (($_.AccessRights) -join ',') }) -join ' | '");
                        sb.AppendLine("    } catch { $fp = '' }");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName " + colName + " -NotePropertyValue $fp -Force");
                    }
                    if (acctStatus)
                    {
                        sb.AppendLine("    $acct = ''; $ukey=[string]$m.Guid; if ($userIndex.ContainsKey($ukey)) { $uu=$userIndex[$ukey]; if ($uu.AccountDisabled) { $acct='Not activated' } else { $acct='Activated' } }");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName AccountStatus -NotePropertyValue $acct -Force");
                    }
                    if (regional)
                    {
                        sb.AppendLine("    try { $rc = Get-MailboxRegionalConfiguration -Identity $m.Identity } catch { $rc = $null }");
                        sb.AppendLine("    $lang=''; $tz=''; if ($rc) { $lang=$rc.Language; $tz=$rc.TimeZone }");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName Language -NotePropertyValue $lang -Force");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName TimeZone -NotePropertyValue $tz -Force");
                    }
                    if (needSize)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName MailboxSizeMB -NotePropertyValue (Get-SizeMB $m.PrimarySmtpAddress) -Force");
                    if (needCount)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName MailboxItemCount -NotePropertyValue (Get-ItemCount $m.PrimarySmtpAddress) -Force");
                    if (needArchive)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName ArchiveSizeMB -NotePropertyValue (Get-ArchiveMB $m.PrimarySmtpAddress) -Force");
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

        private static AuditSection BuildResourceSection(bool isRoom)
        {
            string filterType = isRoom ? "RoomMailbox" : "EquipmentMailbox";
            var section = new AuditSection(
                isRoom ? "room-mailboxes" : "equipment-mailboxes",
                isRoom ? "Room mailboxes" : "Equipment mailboxes",
                isRoom ? "Room mailbox export" : "Equipment mailbox export",
                isRoom
                    ? "Audit room mailboxes - capacity, location and calendar processing."
                    : "Audit equipment mailboxes - resource tags and calendar processing.",
                "mailbox",
                AuditScope.Both);
            section.Category = "Mailboxes";
            section.DefaultFileName = isRoom ? "RoomMailboxes.csv" : "EquipmentMailboxes.csv";

            var identity = new AuditOptionGroup("identity", "Identity / addressing", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("PrimarySmtpAddress", true);
            identity.AddProp("EmailAddresses", false);
            identity.AddProp("ExchangeGuid", false);
            section.AddGroup(identity);

            var type = new AuditOptionGroup("type", "Type / resource", GroupMode.MultiCheck); type.Columns = 2;
            type.AddProp("RecipientTypeDetails", true);
            type.AddProp("IsResource", true);
            section.AddGroup(type);

            var visibility = new AuditOptionGroup("visibility", "Visibility", GroupMode.MultiCheck); visibility.Columns = 2;
            visibility.AddProp("HiddenFromAddressListsEnabled", false);
            section.AddGroup(visibility);

            var quotas = new AuditOptionGroup("quotas", "Quotas", GroupMode.MultiCheck); quotas.Columns = 2;
            quotas.AddProp("ProhibitSendQuota", false);
            quotas.AddProp("ProhibitSendReceiveQuota", false);
            quotas.AddProp("IssueWarningQuota", false);
            section.AddGroup(quotas);

            var retention = new AuditOptionGroup("retention", "Retention / compliance", GroupMode.MultiCheck); retention.Columns = 2;
            retention.AddProp("RetentionPolicy", false);
            retention.AddProp("LitigationHoldEnabled", false);
            retention.AddProp("AuditEnabled", false);
            section.AddGroup(retention);

            var mailflow = new AuditOptionGroup("mailflow", "Mail flow", GroupMode.MultiCheck); mailflow.Columns = 2;
            mailflow.AddProp("ForwardingAddress", false);
            mailflow.AddProp("ForwardingSmtpAddress", false);
            mailflow.AddProp("DeliverToMailboxAndForward", false);
            section.AddGroup(mailflow);

            var account = new AuditOptionGroup("account", "Account", GroupMode.MultiCheck); account.Columns = 2;
            account.AddProp("RoomMailboxAccountEnabled", true);
            account.Add(new AuditOption("accountstatus", "Account activated / disabled (Get-User)", "accountstatus", false));
            account.Add(new AuditOption("mailboxsize", "Mailbox size in MB (Get-MailboxStatistics)", "mailboxsize", false).MarkSlow());
            account.Add(new AuditOption("mailboxitemcount", "Item count (Get-MailboxStatistics)", "mailboxitemcount", false).MarkSlow());
            account.Add(new AuditOption("archivesize", "Archive size in MB (per-mailbox, slower)", "archivesize", false).MarkSlow());
            section.AddGroup(account);

            var custom = new AuditOptionGroup("custom", "Tags / custom", GroupMode.MultiCheck); custom.Columns = 2;
            custom.Add(new AuditOption("ResourceCustom", "ResourceCustom (resource tags)", "ResourceCustom", !isRoom));
            custom.Add(new AuditOption("customattr", "CustomAttribute1-15", "CustomAttribute1-15", false));
            section.AddGroup(custom);

            if (isRoom)
            {
                var roomGrp = new AuditOptionGroup("room", "Room specifics", GroupMode.MultiCheck); roomGrp.Columns = 2;
                roomGrp.AddProp("ResourceCapacity", true);
                roomGrp.AddProp("Office", true);
                section.AddGroup(roomGrp);

                var place = new AuditOptionGroup("place", "Places (Get-Place, EXO only)", GroupMode.MultiCheck);
                place.Hint = "Modern room metadata via Get-Place; blank on-premises or if unavailable.";
                place.Columns = 2;
                place.Add(new AuditOption("Building", "Building", "Building", false));
                place.Add(new AuditOption("Floor", "Floor", "Floor", false));
                place.Add(new AuditOption("FloorLabel", "Floor label", "FloorLabel", false));
                place.Add(new AuditOption("Capacity", "Capacity", "Capacity", false));
                place.Add(new AuditOption("IsWheelChairAccessible", "Wheelchair accessible", "IsWheelChairAccessible", false));
                place.Add(new AuditOption("Phone", "Phone", "Phone", false));
                place.Add(new AuditOption("AudioDeviceName", "Audio device", "AudioDeviceName", false));
                place.Add(new AuditOption("VideoDeviceName", "Video device", "VideoDeviceName", false));
                place.Add(new AuditOption("DisplayDeviceName", "Display device", "DisplayDeviceName", false));
                place.Add(new AuditOption("Tags", "Tags", "Tags", false));
                section.AddGroup(place);
            }

            var calProc = new AuditOptionGroup("calproc", "Calendar processing (Get-CalendarProcessing)", GroupMode.MultiCheck);
            calProc.Hint = "Retrieved per mailbox via Get-CalendarProcessing.";
            calProc.Columns = 2;
            calProc.Add(new AuditOption("AutomateProcessing", "AutomateProcessing", "AutomateProcessing", true));
            calProc.Add(new AuditOption("BookingWindowInDays", "BookingWindowInDays", "BookingWindowInDays", false));
            calProc.Add(new AuditOption("MaximumDurationInMinutes", "MaximumDurationInMinutes", "MaximumDurationInMinutes", false));
            calProc.Add(new AuditOption("AllowConflicts", "AllowConflicts", "AllowConflicts", false));
            calProc.Add(new AuditOption("AllowRecurringMeetings", "AllowRecurringMeetings", "AllowRecurringMeetings", false));
            calProc.Add(new AuditOption("AllBookInPolicy", "AllBookInPolicy", "AllBookInPolicy", false));
            calProc.Add(new AuditOption("BookInPolicy", "BookInPolicy", "BookInPolicy", false));
            calProc.Add(new AuditOption("AllRequestOutOfPolicy", "AllRequestOutOfPolicy", "AllRequestOutOfPolicy", false));
            calProc.Add(new AuditOption("RequestOutOfPolicy", "RequestOutOfPolicy", "RequestOutOfPolicy", false));
            calProc.Add(new AuditOption("ResourceDelegates", "ResourceDelegates", "ResourceDelegates", false));
            calProc.Add(new AuditOption("AddOrganizerToSubject", "AddOrganizerToSubject", "AddOrganizerToSubject", false));
            calProc.Add(new AuditOption("DeleteComments", "DeleteComments", "DeleteComments", false));
            calProc.Add(new AuditOption("DeleteSubject", "DeleteSubject", "DeleteSubject", false));
            section.AddGroup(calProc);

            var contact = new AuditOptionGroup("contact", "Contact info (Get-User)", GroupMode.MultiCheck); contact.Columns = 3;
            contact.Hint = "Contact fields carried on the resource account, pulled from Get-User.";
            contact.Add(new AuditOption("Phone", "Phone", "Phone", false));
            contact.Add(new AuditOption("MobilePhone", "MobilePhone", "MobilePhone", false));
            contact.Add(new AuditOption("Fax", "Fax", "Fax", false));
            contact.Add(new AuditOption("Company", "Company", "Company", false));
            contact.Add(new AuditOption("Department", "Department", "Department", false));
            contact.Add(new AuditOption("StreetAddress", "StreetAddress", "StreetAddress", false));
            contact.Add(new AuditOption("City", "City", "City", false));
            contact.Add(new AuditOption("StateOrProvince", "StateOrProvince", "StateOrProvince", false));
            contact.Add(new AuditOption("PostalCode", "PostalCode", "PostalCode", false));
            contact.Add(new AuditOption("CountryOrRegion", "CountryOrRegion", "CountryOrRegion", false));
            section.AddGroup(contact);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");

                string[] mbxGroups = { "identity", "type", "visibility", "quotas", "retention", "mailflow", "account", "custom", "room" };
                var chosen = new List<string>();
                foreach (string gk in mbxGroups)
                    foreach (string v in sel.Selected(gk))
                        if (v != "accountstatus" && v != "mailboxsize" && !chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("DisplayName");

                bool needResolver = false;
                var exprs = new List<string>();
                foreach (string v in chosen)
                {
                    if (v == "EmailAddresses")
                        exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                    else if (v == "ForwardingAddress" || v == "GrantSendOnBehalfTo") { needResolver = true; exprs.Add("@{Name='" + v + "';Expression={ Resolve-Recip $_." + v + " }}"); }
                    else if (v == "ResourceCustom")
                        exprs.Add("@{Name='ResourceCustom';Expression={($_.ResourceCustom | ForEach-Object { [string]$_ }) -join ','}}");
                    else if (v == "CustomAttribute1-15")
                        for (int i = 1; i <= 15; i++) exprs.Add("CustomAttribute" + i);
                    else exprs.Add(v);
                }
                string selectList = string.Join(", ", exprs.ToArray());

                List<string> cal = sel.Selected("calproc");
                List<string> place = sel.Selected("place");
                List<string> contactProps = sel.Selected("contact");
                bool acctStatus = sel.IsSelected("account", "accountstatus");
                bool needSize = sel.IsSelected("account", "mailboxsize");
                bool needCount = sel.IsSelected("account", "mailboxitemcount");
                bool needArchive = sel.IsSelected("account", "archivesize");
                bool needUser = acctStatus || contactProps.Count > 0;

                if (cal.Contains("ResourceDelegates")) needResolver = true;

                bool online = ConnectionSettings.IsOnline;
                string getMbx = "Get-Mailbox";
                string getRecip = online ? "Get-EXORecipient" : "Get-Recipient";
                string getStats = online ? "Get-EXOMailboxStatistics" : "Get-MailboxStatistics";

                bool perRow = cal.Count > 0 || place.Count > 0 || needUser || needSize || needCount || needArchive;

                var sb = new StringBuilder();
                if (needSize) PsScriptHelpers.EmitSizeHelper(sb, getStats);
                if (needCount) PsScriptHelpers.EmitCountHelper(sb, getStats);
                if (needArchive) PsScriptHelpers.EmitArchiveMBHelper(sb, getStats);
                if (needResolver) PsScriptHelpers.EmitResolver(sb, getRecip, false);

                sb.AppendLine("Write-Host 'Querying " + (isRoom ? "room" : "equipment") + " mailboxes...'");
                sb.AppendLine("$mbx = @(" + getMbx + " -ResultSize " + resultSize + " -RecipientTypeDetails " + filterType + ")");
                sb.AppendLine("Write-Host (\"Retrieved {0} mailbox(es).\" -f $mbx.Count)");
                sb.AppendLine();

                if (needUser)
                {
                    sb.AppendLine("Write-Host 'Indexing Get-User (bulk)...'");
                    sb.AppendLine("$userIndex = @{}");
                    sb.AppendLine("Get-User -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object { $userIndex[[string]$_.Guid] = $_ }");
                    sb.AppendLine();
                }

                if (!perRow)
                {
                    sb.AppendLine("$rows = $mbx | Select-Object " + selectList);
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($m in $mbx) {");
                    sb.AppendLine("    $obj = $m | Select-Object " + selectList);
                    if (cal.Count > 0)
                    {
                        sb.AppendLine("    try { $cp = Get-CalendarProcessing -Identity $m.Identity } catch { $cp = $null }");
                        foreach (string prop in cal)
                        {
                            if (prop == "ResourceDelegates")
                                sb.AppendLine("    $val = if ($cp) { Resolve-Recip $cp.ResourceDelegates } else { '' }");
                            else if (CalProcMultiValued.Contains(prop))
                                sb.AppendLine("    $val = if ($cp) { (@($cp." + prop + " | ForEach-Object { [string]$_ }) -join ',') } else { '' }");
                            else
                                sb.AppendLine("    $val = if ($cp) { [string]$cp." + prop + " } else { '' }");
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName " + prop + " -NotePropertyValue $val -Force");
                        }
                    }
                    if (place.Count > 0)
                    {
                        sb.AppendLine("    $pl = $null");
                        sb.AppendLine("    try { $pl = Get-Place -Identity $m.PrimarySmtpAddress -ErrorAction Stop } catch { $pl = $null }");
                        foreach (string prop in place)
                        {
                            sb.AppendLine("    $pv = if ($pl) { (@($pl." + prop + " | ForEach-Object { [string]$_ }) -join ',') } else { '' }");
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName Place_" + prop + " -NotePropertyValue $pv -Force");
                        }
                    }
                    if (needUser)
                    {
                        sb.AppendLine("    $u = $null; $ukey=[string]$m.Guid; if ($userIndex.ContainsKey($ukey)) { $u = $userIndex[$ukey] }");
                        foreach (string cp2 in contactProps)
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName " + cp2 + " -NotePropertyValue ($(if ($u) { [string]$u." + cp2 + " } else { '' })) -Force");
                        if (acctStatus)
                        {
                            sb.AppendLine("    $acct = ''; if ($u) { if ($u.AccountDisabled) { $acct='Not activated' } else { $acct='Activated' } }");
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName AccountStatus -NotePropertyValue $acct -Force");
                        }
                    }
                    if (needSize)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName MailboxSizeMB -NotePropertyValue (Get-SizeMB $m.PrimarySmtpAddress) -Force");
                    if (needCount)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName MailboxItemCount -NotePropertyValue (Get-ItemCount $m.PrimarySmtpAddress) -Force");
                    if (needArchive)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName ArchiveSizeMB -NotePropertyValue (Get-ArchiveMB $m.PrimarySmtpAddress) -Force");
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
