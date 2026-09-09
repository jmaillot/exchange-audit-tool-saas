using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    internal static class SectionsMailboxes
    {
        public static void RegisterMailboxSections()
        {
            AuditRegistry.Register(BuildMailboxListSection());
            AuditRegistry.Register(BuildMobileDevicesSection());
        }

        private static readonly List<string> MultiValued = new List<string>(new string[]
        {
            "InPlaceHolds", "Languages", "AddressListMembership"
        });

        private static readonly List<string> RecipientMultiValued = new List<string>(new string[]
        {
            "ForwardingAddress", "GrantSendOnBehalfTo", "ModeratedBy",
            "AcceptMessagesOnlyFromDLMembers", "AcceptMessagesOnlyFromSendersOrMembers",
            "RejectMessagesFromDLMembers", "RejectMessagesFromSendersOrMembers"
        });

        // Attributes that are NOT usable / not meaningful in Exchange Online. Even if the user
        // ticks them, they are dropped from the export when connected to EXO.
        private static readonly List<string> OnPremOnly = new List<string>(new string[]
        {
            "Database", "ServerName", "UseDatabaseQuotaDefaults", "UseDatabaseRetentionDefaults", "ThrottlingPolicy"
        });

        private static readonly string[] PropGroupKeys =
        {
            "identity", "addressing", "type", "visibility", "custom", "quotas", "msgsize",
            "retention", "audit", "mailflow", "copies", "archive", "oof", "misc"
        };

        private static readonly string[] UserGroupKeys = { "personal", "hierarchy" };

        private static readonly List<string> UserSimpleProps = new List<string>(new string[]
        {
            "FirstName", "LastName", "Initials", "Title", "Department", "Company",
            "Phone", "MobilePhone", "Fax", "HomePhone", "OtherTelephone", "Pager",
            "StreetAddress", "City", "StateOrProvince", "PostalCode", "CountryOrRegion", "WebPage"
        });

        private static AuditSection BuildMailboxListSection()
        {
            var section = new AuditSection(
                "mailbox-list",
                "User mailboxes",
                "User mailbox export",
                "Export USER mailboxes only. Defaults adapt to Exchange Online vs on-premises.",
                "mailbox",
                AuditScope.Both);
            section.Category = "Mailboxes";
            section.DefaultFileName = "UserMailboxes.csv";
            section.ScopeAwareDefaults = true;

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 3;
            identity.Add(AuditOption.PropScoped("DisplayName", true, true));
            identity.Add(AuditOption.PropScoped("Alias", true, true));
            identity.Add(AuditOption.PropScoped("SamAccountName", false, true));
            identity.Add(AuditOption.PropScoped("UserPrincipalName", true, true));
            identity.Add(AuditOption.PropScoped("PrimarySmtpAddress", true, true));
            identity.Add(AuditOption.PropScoped("ExchangeGuid", true, true));
            section.AddGroup(identity);

            var addressing = new AuditOptionGroup("addressing", "Addressing", GroupMode.MultiCheck); addressing.Columns = 2;
            addressing.Hint = "EmailAddresses is filtered to smtp:/SMTP: entries only (SIP/SPO/X500 excluded).";
            addressing.Add(AuditOption.PropScoped("EmailAddresses", true, true));
            addressing.Add(AuditOption.PropScoped("EmailAddressPolicyEnabled", false, true));
            section.AddGroup(addressing);

            var type = new AuditOptionGroup("type", "Type", GroupMode.MultiCheck); type.Columns = 2;
            type.Add(AuditOption.PropScoped("RecipientTypeDetails", true, true));
            section.AddGroup(type);

            var visibility = new AuditOptionGroup("visibility", "Visibility", GroupMode.MultiCheck); visibility.Columns = 3;
            visibility.Add(AuditOption.PropScoped("HiddenFromAddressListsEnabled", true, true));
            visibility.Add(AuditOption.PropScoped("MailTip", false, false));
            section.AddGroup(visibility);

            var custom = new AuditOptionGroup("custom", "Custom attributes", GroupMode.MultiCheck); custom.Columns = 2;
            custom.Hint = "Ranges expand to every attribute in the range.";
            custom.Add(new AuditOption("customattr", "CustomAttribute1-15", "CustomAttribute1-15", false));
            custom.Add(new AuditOption("extcustomattr", "ExtensionCustomAttribute1-5", "ExtensionCustomAttribute1-5", false));
            section.AddGroup(custom);

            var quotas = new AuditOptionGroup("quotas", "Quotas", GroupMode.MultiCheck); quotas.Columns = 2;
            quotas.Add(AuditOption.PropScoped("ProhibitSendQuota", false, true));
            quotas.Add(AuditOption.PropScoped("ProhibitSendReceiveQuota", false, true));
            quotas.Add(AuditOption.PropScoped("IssueWarningQuota", false, true));
            quotas.Add(AuditOption.PropScoped("UseDatabaseQuotaDefaults", false, true));   // on-prem only
            quotas.Add(AuditOption.PropScoped("RulesQuota", false, false));
            section.AddGroup(quotas);

            var msgsize = new AuditOptionGroup("msgsize", "Message size", GroupMode.MultiCheck); msgsize.Columns = 3;
            msgsize.Add(AuditOption.PropScoped("MaxSendSize", false, false));
            msgsize.Add(AuditOption.PropScoped("MaxReceiveSize", false, false));
            msgsize.Add(AuditOption.PropScoped("RecipientLimits", false, false));
            section.AddGroup(msgsize);

            var retention = new AuditOptionGroup("retention", "Retention / compliance", GroupMode.MultiCheck); retention.Columns = 2;
            retention.Add(AuditOption.PropScoped("RetentionPolicy", true, true));
            retention.Add(AuditOption.PropScoped("UseDatabaseRetentionDefaults", false, true));   // on-prem only
            retention.Add(AuditOption.PropScoped("LitigationHoldEnabled", true, true));
            retention.Add(AuditOption.PropScoped("LitigationHoldDuration", false, false));
            retention.Add(AuditOption.PropScoped("SingleItemRecoveryEnabled", true, true));
            retention.Add(AuditOption.PropScoped("RetentionHoldEnabled", false, false));
            retention.Add(AuditOption.PropScoped("RetainDeletedItemsFor", false, false));
            retention.Add(AuditOption.PropScoped("InPlaceHolds", true, false));
            section.AddGroup(retention);

            var audit = new AuditOptionGroup("audit", "Audit", GroupMode.MultiCheck); audit.Columns = 3;
            audit.Add(AuditOption.PropScoped("AuditEnabled", false, false));
            audit.Add(AuditOption.PropScoped("AuditLogAgeLimit", false, false));
            audit.Add(AuditOption.PropScoped("AuditAdmin", false, false));
            audit.Add(AuditOption.PropScoped("AuditDelegate", false, false));
            audit.Add(AuditOption.PropScoped("AuditOwner", false, false));
            audit.Add(AuditOption.PropScoped("DefaultAuditSet", false, false));
            section.AddGroup(audit);

            var mailflow = new AuditOptionGroup("mailflow", "Mail flow", GroupMode.MultiCheck); mailflow.Columns = 2;
            mailflow.Add(AuditOption.PropScoped("ForwardingAddress", true, true));
            mailflow.Add(AuditOption.PropScoped("ForwardingSmtpAddress", true, true));
            mailflow.Add(AuditOption.PropScoped("DeliverToMailboxAndForward", true, true));
            mailflow.Add(AuditOption.PropScoped("AcceptMessagesOnlyFromDLMembers", false, false));
            mailflow.Add(AuditOption.PropScoped("AcceptMessagesOnlyFromSendersOrMembers", false, false));
            mailflow.Add(AuditOption.PropScoped("RejectMessagesFromDLMembers", false, false));
            mailflow.Add(AuditOption.PropScoped("RejectMessagesFromSendersOrMembers", false, false));
            mailflow.Add(AuditOption.PropScoped("RequireSenderAuthenticationEnabled", false, false));
            mailflow.Add(AuditOption.PropScoped("ModeratedBy", false, false));
            mailflow.Add(AuditOption.PropScoped("ModerationEnabled", false, false));
            section.AddGroup(mailflow);

            var copies = new AuditOptionGroup("copies", "Copies / traceability", GroupMode.MultiCheck); copies.Columns = 1;
            copies.Add(AuditOption.PropScoped("MessageCopyForSentAsEnabled", false, false));
            copies.Add(AuditOption.PropScoped("MessageCopyForSendOnBehalfEnabled", false, false));
            copies.Add(AuditOption.PropScoped("MessageCopyForSMTPClientSubmissionEnabled", false, false));
            section.AddGroup(copies);

            var archive = new AuditOptionGroup("archive", "Archive", GroupMode.MultiCheck); archive.Columns = 2;
            archive.Add(AuditOption.PropScoped("ArchiveStatus", true, true));
            archive.Add(AuditOption.PropScoped("ArchiveState", true, true));
            archive.Add(AuditOption.PropScoped("ArchiveQuota", false, false));
            archive.Add(AuditOption.PropScoped("ArchiveWarningQuota", false, false));
            archive.Add(AuditOption.PropScoped("AutoExpandingArchiveEnabled", true, false));
            archive.Add(AuditOption.PropScoped("ArchiveName", false, false));
            section.AddGroup(archive);

            var oof = new AuditOptionGroup("oof", "OOF / anti-spam", GroupMode.MultiCheck); oof.Columns = 2;
            oof.Add(AuditOption.PropScoped("ExternalOofOptions", false, false));
            oof.Add(AuditOption.PropScoped("AntispamBypassEnabled", false, false));
            section.AddGroup(oof);

            var misc = new AuditOptionGroup("misc", "Misc", GroupMode.MultiCheck); misc.Columns = 2;
            misc.Add(AuditOption.PropScoped("Office", true, true));
            misc.Add(AuditOption.PropScoped("IsDirSynced", true, false));
            misc.Add(AuditOption.PropScoped("Database", false, true));         // on-prem only
            misc.Add(AuditOption.PropScoped("ServerName", false, true));       // on-prem only
            misc.Add(AuditOption.PropScoped("Languages", false, false));
            misc.Add(AuditOption.PropScoped("ThrottlingPolicy", false, false)); // on-prem only
            misc.Add(AuditOption.PropScoped("RoleAssignmentPolicy", false, false));
            misc.Add(AuditOption.PropScoped("SharingPolicy", false, false));
            misc.Add(AuditOption.PropScoped("AddressBookPolicy", false, false));
            misc.Add(AuditOption.PropScoped("OfflineAddressBook", false, false));
            misc.Add(AuditOption.PropScoped("AddressListMembership", false, false));
            section.AddGroup(misc);

            var personal = new AuditOptionGroup("personal", "Personal / contact (Get-User)", GroupMode.MultiCheck);
            personal.Hint = "Pulled from Get-User in one bulk pass and matched by GUID - no Graph/AzureAD module needed.";
            personal.Columns = 3;
            personal.Add(AuditOption.PropScoped("FirstName", false, false));
            personal.Add(AuditOption.PropScoped("LastName", false, false));
            personal.Add(AuditOption.PropScoped("Initials", false, false));
            personal.Add(AuditOption.PropScoped("Title", false, false));
            personal.Add(AuditOption.PropScoped("Department", false, false));
            personal.Add(AuditOption.PropScoped("Company", false, false));
            personal.Add(AuditOption.PropScoped("Phone", false, false));
            personal.Add(AuditOption.PropScoped("MobilePhone", false, false));
            personal.Add(AuditOption.PropScoped("Fax", false, false));
            personal.Add(AuditOption.PropScoped("HomePhone", false, false));
            personal.Add(AuditOption.PropScoped("OtherTelephone", false, false));
            personal.Add(AuditOption.PropScoped("Pager", false, false));
            personal.Add(AuditOption.PropScoped("StreetAddress", false, false));
            personal.Add(AuditOption.PropScoped("City", false, false));
            personal.Add(AuditOption.PropScoped("StateOrProvince", false, false));
            personal.Add(AuditOption.PropScoped("PostalCode", false, false));
            personal.Add(AuditOption.PropScoped("CountryOrRegion", false, false));
            personal.Add(AuditOption.PropScoped("WebPage", false, false));
            section.AddGroup(personal);

            var hierarchy = new AuditOptionGroup("hierarchy", "Hierarchy (Get-User)", GroupMode.MultiCheck);
            hierarchy.Hint = "Manager and DirectReports are DN references, resolved to the target's SMTP address.";
            hierarchy.Columns = 2;
            hierarchy.Add(AuditOption.PropScoped("Manager", false, false));
            hierarchy.Add(AuditOption.PropScoped("DirectReports", false, false));
            section.AddGroup(hierarchy);

            var status = new AuditOptionGroup("accountstatus", "Account status (Get-User)", GroupMode.MultiCheck);
            status.Hint = "Reports whether the sign-in account is enabled, as 'Activated' / 'Not activated'.";
            status.Columns = 1;
            status.Add(new AuditOption("AccountStatus", "Account activated / disabled", "AccountStatus", false));
            section.AddGroup(status);

            var extra = new AuditOptionGroup("complementary", "Complementary data (bulk-indexed)", GroupMode.MultiCheck);
            extra.Hint = "Permissions are pre-fetched in one bulk pass. Regional config, counts and archive lookups are per-mailbox.";
            extra.Columns = 1;
            extra.Add(new AuditOption("fullaccess", "FullAccess delegates - bulk Get-(EXO)MailboxPermission", "fullaccess", false));
            extra.Add(new AuditOption("sendas", "SendAs delegates - EXO: RecipientPermission / on-prem: ADPermission", "sendas", false));
            extra.Add(new AuditOption("grantsendonbehalf", "Send on Behalf (GrantSendOnBehalfTo, resolved to SMTP)", "grantsendonbehalf", false));
            extra.Add(new AuditOption("mailboxsize", "Mailbox size in MB (Get-MailboxStatistics)", "mailboxsize", false).MarkSlow());
            extra.Add(new AuditOption("mailboxitemcount", "Item count (Get-MailboxStatistics)", "mailboxitemcount", false).MarkSlow());
            extra.Add(new AuditOption("archivesize", "Archive size in MB (per-mailbox, slower)", "archivesize", false).MarkSlow());
            extra.Add(new AuditOption("regional", "Regional config (Language, TimeZone) - per-mailbox, slower", "regional", false).MarkSlow());
            section.AddGroup(extra);

            var folderperms = new AuditOptionGroup("folderperms", "Folder permissions (per-mailbox, slower)", GroupMode.MultiCheck);
            folderperms.Hint = "Get-MailboxFolderPermission on the chosen folder(s); non-default entries only.";
            folderperms.Columns = 2;
            folderperms.MarkSlow();
            folderperms.Add(new AuditOption("calendar", "Calendar", "Calendar", false));
            folderperms.Add(new AuditOption("inbox", "Inbox", "Inbox", false));
            folderperms.Add(new AuditOption("sentitems", "Sent Items", "SentItems", false));
            folderperms.Add(new AuditOption("contacts", "Contacts", "Contacts", false));
            folderperms.Add(new AuditOption("tasks", "Tasks", "Tasks", false));
            section.AddGroup(folderperms);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");
                bool online = ConnectionSettings.IsOnline;

                var chosen = new List<string>();
                foreach (string gk in PropGroupKeys)
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                // Send on Behalf now lives in complementary data.
                if (sel.IsSelected("complementary", "grantsendonbehalf") && !chosen.Contains("GrantSendOnBehalfTo"))
                    chosen.Add("GrantSendOnBehalfTo");
                // Drop attributes that are not usable in Exchange Online.
                if (online) chosen.RemoveAll(delegate (string v) { return OnPremOnly.Contains(v); });
                if (chosen.Count == 0) chosen.Add("DisplayName");

                bool needResolver = false;
                var rawNames = new List<string>();
                var exprs = new List<string>();
                foreach (string v in chosen)
                {
                    if (v == "EmailAddresses")
                    {
                        rawNames.Add("EmailAddresses");
                        exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                    }
                    else if (v == "MailTip")
                    {
                        rawNames.Add("MailTip");
                        exprs.Add("@{Name='MailTip';Expression={ if ([string]::IsNullOrWhiteSpace([string]$_.MailTip)) { 'Not activated' } else { 'Activated' } }}");
                    }
                    else if (v == "CustomAttribute1-15")
                        for (int i = 1; i <= 15; i++) { rawNames.Add("CustomAttribute" + i); exprs.Add("CustomAttribute" + i); }
                    else if (v == "ExtensionCustomAttribute1-5")
                        for (int i = 1; i <= 5; i++)
                        {
                            rawNames.Add("ExtensionCustomAttribute" + i);
                            exprs.Add("@{Name='ExtensionCustomAttribute" + i + "';Expression={($_.ExtensionCustomAttribute" + i + " | ForEach-Object { [string]$_ }) -join ','}}");
                        }
                    else if (RecipientMultiValued.Contains(v))
                    {
                        needResolver = true;
                        rawNames.Add(v);
                        exprs.Add("@{Name='" + v + "';Expression={ Resolve-Recip $_." + v + " }}");
                    }
                    else if (MultiValued.Contains(v))
                    {
                        rawNames.Add(v);
                        exprs.Add("@{Name='" + v + "';Expression={($_." + v + " | ForEach-Object { [string]$_ }) -join ','}}");
                    }
                    else { rawNames.Add(v); exprs.Add(v); }
                }
                if (!rawNames.Contains("Identity")) rawNames.Add("Identity");

                var userSimple = new List<string>();
                foreach (string gk in UserGroupKeys)
                    foreach (string v in sel.Selected(gk))
                        if (UserSimpleProps.Contains(v) && !userSimple.Contains(v)) userSimple.Add(v);
                bool needMgr = sel.IsSelected("hierarchy", "Manager");
                bool needDR = sel.IsSelected("hierarchy", "DirectReports");
                bool needAcctStatus = sel.IsSelected("accountstatus", "AccountStatus");
                bool needUser = userSimple.Count > 0 || needMgr || needDR || needAcctStatus;
                if (needUser && !rawNames.Contains("Guid")) rawNames.Add("Guid");

                // Manager / DirectReports resolve to the target's SMTP address (email), not display name.
                if (needMgr || needDR) needResolver = true;
                string selectList = string.Join(", ", exprs.ToArray());

                bool regional = sel.IsSelected("complementary", "regional");
                bool fa = sel.IsSelected("complementary", "fullaccess");
                bool sa = sel.IsSelected("complementary", "sendas");
                List<string> folders = sel.Selected("folderperms");
                bool needSize = sel.IsSelected("complementary", "mailboxsize");
                bool needCount = sel.IsSelected("complementary", "mailboxitemcount");
                bool needArchive = sel.IsSelected("complementary", "archivesize");

                string getMbx = online ? "Get-EXOMailbox" : "Get-Mailbox";
                string getMbxPerm = online ? "Get-EXOMailboxPermission" : "Get-MailboxPermission";
                string getRecPerm = online ? "Get-EXORecipientPermission" : "Get-RecipientPermission";
                string getRecip = online ? "Get-EXORecipient" : "Get-Recipient";
                string getStats = online ? "Get-EXOMailboxStatistics" : "Get-MailboxStatistics";
                string propsArg = online ? " -Properties " + string.Join(",", rawNames.ToArray()) : "";

                var sb = new StringBuilder();

                if (needResolver) PsScriptHelpers.EmitResolver(sb, getRecip, false);

                if (needSize) PsScriptHelpers.EmitSizeHelper(sb, getStats);
                if (needCount) PsScriptHelpers.EmitCountHelper(sb, getStats);
                if (needArchive) PsScriptHelpers.EmitArchiveMBHelper(sb, getStats);

                sb.AppendLine("Write-Host 'Querying USER mailboxes...'");
                sb.AppendLine("$mbx = @(" + getMbx + " -ResultSize " + resultSize + " -RecipientTypeDetails UserMailbox" + propsArg + ")");
                sb.AppendLine("Write-Host (\"Retrieved {0} user mailbox(es).\" -f $mbx.Count)");
                sb.AppendLine();

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
                if (needUser)
                {
                    sb.AppendLine("Write-Host 'Indexing Get-User (bulk)...'");
                    sb.AppendLine("$userIndex = @{}");
                    sb.AppendLine("Get-User -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object { $userIndex[[string]$_.Guid] = $_ }");
                    sb.AppendLine();
                }

                if (!fa && !sa && !regional && !needUser && !needSize && !needCount && !needArchive && folders.Count == 0)
                {
                    sb.AppendLine("$rows = $mbx | Select-Object " + selectList);
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($m in $mbx) {");
                    sb.AppendLine("    $obj = $m | Select-Object " + selectList);
                    sb.AppendLine("    $key = [string]$m.Identity");
                    if (needUser)
                    {
                        sb.AppendLine("    $u = $null; $ukey = [string]$m.Guid; if ($userIndex.ContainsKey($ukey)) { $u = $userIndex[$ukey] }");
                        foreach (string up in userSimple)
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName " + up + " -NotePropertyValue ($(if ($u) { [string]$u." + up + " } else { '' })) -Force");
                        if (needMgr)
                        {
                            sb.AppendLine("    $mgrVal = ''; if ($u -and $u.Manager) { $mgrVal = Resolve-Recip $u.Manager }");
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName Manager -NotePropertyValue $mgrVal -Force");
                        }
                        if (needDR)
                        {
                            sb.AppendLine("    $drVal = ''; if ($u -and $u.DirectReports) { $drVal = Resolve-Recip $u.DirectReports }");
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName DirectReports -NotePropertyValue $drVal -Force");
                        }
                        if (needAcctStatus)
                        {
                            sb.AppendLine("    $acct = ''; if ($u) { if ($u.AccountDisabled) { $acct = 'Not activated' } else { $acct = 'Activated' } }");
                            sb.AppendLine("    $obj | Add-Member -NotePropertyName AccountStatus -NotePropertyValue $acct -Force");
                        }
                    }
                    if (fa)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName FullAccess -NotePropertyValue ($(if ($faIndex.ContainsKey($key)) { $faIndex[$key] } else { '' })) -Force");
                    if (sa)
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName SendAs -NotePropertyValue ($(if ($saIndex.ContainsKey($key)) { $saIndex[$key] } else { '' })) -Force");
                    if (regional)
                    {
                        sb.AppendLine("    try { $rc = Get-MailboxRegionalConfiguration -Identity $m.Identity } catch { $rc = $null }");
                        sb.AppendLine("    $lang = ''; $tz = ''");
                        sb.AppendLine("    if ($rc) { $lang = $rc.Language; $tz = $rc.TimeZone }");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName Language -NotePropertyValue $lang -Force");
                        sb.AppendLine("    $obj | Add-Member -NotePropertyName TimeZone -NotePropertyValue $tz -Force");
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

        // ============================================================ MOBILE DEVICES
        private static AuditSection BuildMobileDevicesSection()
        {
            var section = new AuditSection(
                "mobile-devices",
                "Mobile devices",
                "Mobile device export",
                "Audit mobile partnerships (Get-MobileDevice): wipe / re-enroll planning. One call per mailbox.",
                "mobile",
                AuditScope.Both);
            section.Category = "Mobile";
            section.DefaultFileName = "MobileDevices.csv";

            var devices = new AuditOptionGroup("devices", "Device properties", GroupMode.MultiCheck); devices.Columns = 2;
            devices.MarkSlow();
            devices.AddProp("DeviceId", true);
            devices.AddProp("DeviceType", true);
            devices.AddProp("DeviceModel", true);
            devices.AddProp("DeviceOS", true);
            devices.AddProp("DeviceAccessState", true);
            devices.AddProp("FirstSyncTime", false);
            devices.AddProp("LastSyncAttemptTime", false);
            devices.AddProp("LastSuccessSync", false);
            section.AddGroup(devices);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");
                var chosen = new List<string>();
                foreach (string v in sel.Selected("devices"))
                    if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("DeviceId");

                bool online = ConnectionSettings.IsOnline;
                string getMbx = online ? "Get-EXOMailbox" : "Get-Mailbox";

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying USER mailboxes...'");
                sb.AppendLine("$mbx = @(" + getMbx + " -ResultSize " + resultSize + " -RecipientTypeDetails UserMailbox)");
                sb.AppendLine("Write-Host (\"Retrieved {0} user mailbox(es).\" -f $mbx.Count)");
                sb.AppendLine("Write-Host 'Querying mobile devices (one call per mailbox)...'");
                sb.AppendLine("$rows = foreach ($m in $mbx) {");
                sb.AppendLine("    $devs = @(Get-MobileDevice -Mailbox $m.Identity -ErrorAction SilentlyContinue)");
                sb.AppendLine("    foreach ($d in $devs) {");
                sb.AppendLine("        $obj = New-Object psobject");
                sb.AppendLine("        $obj | Add-Member -NotePropertyName MailboxDisplayName -NotePropertyValue ([string]$m.DisplayName) -Force");
                sb.AppendLine("        $obj | Add-Member -NotePropertyName PrimarySmtpAddress -NotePropertyValue ([string]$m.PrimarySmtpAddress) -Force");
                foreach (string prop in chosen)
                    sb.AppendLine("        $obj | Add-Member -NotePropertyName " + prop + " -NotePropertyValue ([string]$d." + prop + ") -Force");
                sb.AppendLine("        $obj");
                sb.AppendLine("    }");
                sb.AppendLine("}");
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }
    }
}
