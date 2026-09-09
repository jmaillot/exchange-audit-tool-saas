using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    internal static class SectionsGroups
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildDistLikeSection(false));
            AuditRegistry.Register(BuildDistLikeSection(true));
            AuditRegistry.Register(BuildDynamicGroupSection());
            AuditRegistry.Register(BuildUnifiedGroupSection());
        }

        private static readonly List<string> RecipientRefs = new List<string>(new string[]
        {
            "ManagedBy", "ModeratedBy", "BypassModerationFromSendersOrMembers",
            "AcceptMessagesOnlyFrom", "AcceptMessagesOnlyFromDLMembers", "AcceptMessagesOnlyFromSendersOrMembers",
            "RejectMessagesFrom", "RejectMessagesFromDLMembers", "RejectMessagesFromSendersOrMembers",
            "GrantSendOnBehalfTo"
        });

        private static readonly List<string> MultiValued = new List<string>(new string[]
        {
            "EmailAddresses", "AddressListMembership", "IncludedRecipients",
            "ConditionalCompany", "ConditionalDepartment", "ConditionalStateOrProvince", "InPlaceHolds"
        });

        private static readonly string[] DistPropGroups =
        {
            "identity", "owners", "moderation", "restrictions", "ndr", "limits"
        };

        private static AuditSection BuildDistLikeSection(bool isSecurity)
        {
            string filterType = isSecurity ? "MailUniversalSecurityGroup" : "MailUniversalDistributionGroup";
            var section = new AuditSection(
                isSecurity ? "security-groups" : "distribution-groups",
                isSecurity ? "Security groups" : "Distribution groups",
                isSecurity ? "Mail-enabled security group export" : "Distribution group export",
                isSecurity
                    ? "Audit mail-enabled security groups (Universal, SecurityEnabled). Defaults adapt to EXO vs on-premises."
                    : "Audit distribution groups. Defaults adapt to Exchange Online vs on-premises.",
                "group",
                AuditScope.Both);
            section.Category = "Groups";
            section.DefaultFileName = isSecurity ? "SecurityGroups.csv" : "DistributionGroups.csv";
            section.ScopeAwareDefaults = true;

            var identity = new AuditOptionGroup("identity", "Identity & addressing", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("Name", false);
            identity.AddProp("PrimarySmtpAddress", true);
            identity.AddProp("EmailAddresses", true);
            identity.AddProp("RecipientTypeDetails", true);
            if (isSecurity) identity.AddProp("GroupType", true);
            identity.AddProp("HiddenFromAddressListsEnabled", true);
            identity.AddProp("Description", false);
            identity.Add(new AuditOption("customattr", "CustomAttribute1-15", "CustomAttribute1-15", false));
            identity.Add(new AuditOption("extcustomattr", "ExtensionCustomAttribute1-5", "ExtensionCustomAttribute1-5", false));
            section.AddGroup(identity);

            var owners = new AuditOptionGroup("owners", "Owners & management", GroupMode.MultiCheck); owners.Columns = 2;
            owners.Hint = "ManagedBy is resolved from GUID to SMTP - without an owner the group is unmanageable after migration.";
            owners.AddProp("ManagedBy", true);
            owners.AddProp("MemberJoinRestriction", true);
            owners.AddProp("MemberDepartRestriction", true);
            owners.AddProp("HiddenGroupMembershipEnabled", false);
            section.AddGroup(owners);

            var moderation = new AuditOptionGroup("moderation", "Moderation", GroupMode.MultiCheck); moderation.Columns = 2;
            moderation.AddProp("ModerationEnabled", true);
            moderation.AddProp("ModeratedBy", true);
            moderation.AddProp("BypassModerationFromSendersOrMembers", false);
            moderation.AddProp("BypassNestedModerationEnabled", false);
            moderation.AddProp("SendModerationNotifications", false);
            section.AddGroup(moderation);

            var restrictions = new AuditOptionGroup("restrictions", "Send restrictions / security", GroupMode.MultiCheck); restrictions.Columns = 2;
            restrictions.AddProp("RequireSenderAuthenticationEnabled", true);
            restrictions.AddProp("AcceptMessagesOnlyFrom", false);
            restrictions.AddProp("AcceptMessagesOnlyFromDLMembers", false);
            restrictions.AddProp("AcceptMessagesOnlyFromSendersOrMembers", false);
            restrictions.AddProp("RejectMessagesFrom", false);
            restrictions.AddProp("RejectMessagesFromDLMembers", false);
            restrictions.AddProp("RejectMessagesFromSendersOrMembers", false);
            restrictions.Add(AuditOption.PropScoped("BccBlocked", true, false));
            section.AddGroup(restrictions);

            var ndr = new AuditOptionGroup("ndr", "Delivery reports (NDR)", GroupMode.MultiCheck); ndr.Columns = 3;
            ndr.AddProp("ReportToManagerEnabled", false);
            ndr.AddProp("ReportToOriginatorEnabled", false);
            ndr.AddProp("SendOofMessageToOriginatorEnabled", false);
            section.AddGroup(ndr);

            var limits = new AuditOptionGroup("limits", "Technical limits", GroupMode.MultiCheck); limits.Columns = 3;
            limits.AddProp("MaxSendSize", false);
            limits.AddProp("MaxReceiveSize", false);
            limits.Add(AuditOption.PropScoped("ExpansionServer", false, true));
            section.AddGroup(limits);

            var permissions = new AuditOptionGroup("permissions", "Permissions", GroupMode.MultiCheck); permissions.Columns = 1;
            permissions.Hint = "Send As (EXO: RecipientPermission / on-prem: ADPermission); Send on Behalf from GrantSendOnBehalfTo. Both resolved to SMTP.";
            permissions.Add(new AuditOption("sendas", "Send As delegates (EXO: RecipientPermission / on-prem: ADPermission)", "sendas", false));
            permissions.Add(new AuditOption("sendonbehalf", "Send on Behalf (GrantSendOnBehalfTo)", "sendonbehalf", false));
            section.AddGroup(permissions);

            var membership = new AuditOptionGroup("membership", "Membership (Get-DistributionGroupMember)", GroupMode.SingleChoice);
            membership.Hint = "Expanding produces one row per group+member; members are resolved to name/SMTP/type.";
            membership.Columns = 1;
            membership.Add(new AuditOption("groups", "Group rows only", "groups", true));
            membership.Add(new AuditOption("count", "Add a member-count column", "count", false));
            membership.Add(new AuditOption("expand", "Expand members (one row per member)", "expand", false));
            section.AddGroup(membership);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");
                string membershipMode = sel.First("membership", "groups");
                bool sa = sel.IsSelected("permissions", "sendas");
                bool sob = sel.IsSelected("permissions", "sendonbehalf");

                var chosen = new List<string>();
                foreach (string gk in DistPropGroups)
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("DisplayName");

                bool needResolver;
                string selectList = BuildSelectList(chosen, out needResolver);

                var sb = new StringBuilder();
                if (needResolver || sob || membershipMode == "expand") { string getRecip = ConnectionSettings.IsOnline ? "Get-EXORecipient" : "Get-Recipient"; PsScriptHelpers.EmitResolver(sb, getRecip, false); }

                sb.AppendLine("Write-Host 'Querying groups...'");
                sb.AppendLine("$groups = @(Get-DistributionGroup -ResultSize " + resultSize + " -RecipientTypeDetails " + filterType + ")");
                sb.AppendLine("Write-Host (\"Retrieved {0} group(s).\" -f $groups.Count)");
                sb.AppendLine();
                if (sa) EmitSendAsIndex(sb);

                if (membershipMode == "groups")
                {
                    if (!sa && !sob)
                    {
                        sb.AppendLine("$rows = $groups | Select-Object " + selectList);
                    }
                    else
                    {
                        sb.AppendLine("$rows = foreach ($g in $groups) {");
                        sb.AppendLine("    $o = $g | Select-Object " + selectList);
                        if (sa) sb.AppendLine("    $o | Add-Member -NotePropertyName SendAs -NotePropertyValue " + SaValueExpr + " -Force");
                        if (sob) sb.AppendLine("    $o | Add-Member -NotePropertyName SendOnBehalf -NotePropertyValue " + SobValueExpr + " -Force");
                        sb.AppendLine("    $o");
                        sb.AppendLine("}");
                    }
                }
                else if (membershipMode == "count")
                {
                    sb.AppendLine("$rows = foreach ($g in $groups) {");
                    sb.AppendLine("    try { $members = @(Get-DistributionGroupMember -Identity $g.Identity -ResultSize Unlimited) } catch { $members = @() }");
                    sb.AppendLine("    $o = $g | Select-Object " + selectList);
                    sb.AppendLine("    $o | Add-Member -NotePropertyName MemberCount -NotePropertyValue $members.Count -Force");
                    if (sa) sb.AppendLine("    $o | Add-Member -NotePropertyName SendAs -NotePropertyValue " + SaValueExpr + " -Force");
                    if (sob) sb.AppendLine("    $o | Add-Member -NotePropertyName SendOnBehalf -NotePropertyValue " + SobValueExpr + " -Force");
                    sb.AppendLine("    $o");
                    sb.AppendLine("}");
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($g in $groups) {");
                    sb.AppendLine("    try { $members = @(Get-DistributionGroupMember -Identity $g.Identity -ResultSize Unlimited) } catch { $members = @() }");
                    sb.AppendLine("    $saVal = " + (sa ? SaValueExpr : "''"));
                    sb.AppendLine("    $sobVal = " + (sob ? SobValueExpr : "''"));
                    sb.AppendLine("    if ($members.Count -eq 0) {");
                    sb.AppendLine("        [PSCustomObject]@{ GroupDisplayName=$g.DisplayName; GroupPrimarySmtpAddress=[string]$g.PrimarySmtpAddress; GroupSendAs=$saVal; GroupSendOnBehalf=$sobVal; MemberDisplayName=''; MemberPrimarySmtpAddress=''; MemberRecipientType='' }");
                    sb.AppendLine("    } else {");
                    sb.AppendLine("        foreach ($mem in $members) {");
                    sb.AppendLine("            [PSCustomObject]@{ GroupDisplayName=$g.DisplayName; GroupPrimarySmtpAddress=[string]$g.PrimarySmtpAddress; GroupSendAs=$saVal; GroupSendOnBehalf=$sobVal; MemberDisplayName=$mem.DisplayName; MemberPrimarySmtpAddress=[string]$mem.PrimarySmtpAddress; MemberRecipientType=[string]$mem.RecipientTypeDetails }");
                    sb.AppendLine("        }");
                    sb.AppendLine("    }");
                    sb.AppendLine("}");
                }

                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        private static readonly string[] DynPropGroups =
        {
            "identity", "filter", "owner", "moderation", "restrictions", "ndr"
        };

        private static AuditSection BuildDynamicGroupSection()
        {
            var section = new AuditSection(
                "dynamic-groups",
                "Dynamic groups",
                "Dynamic distribution group export",
                "Audit dynamic distribution groups - membership is a recalculated filter, not a static list.",
                "group",
                AuditScope.Both);
            section.Category = "Groups";
            section.DefaultFileName = "DynamicGroups.csv";
            section.ScopeAwareDefaults = true;

            var identity = new AuditOptionGroup("identity", "Identity & addressing", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("Name", false);
            identity.AddProp("PrimarySmtpAddress", true);
            identity.AddProp("EmailAddresses", true);
            identity.AddProp("RecipientTypeDetails", true);
            identity.AddProp("HiddenFromAddressListsEnabled", true);
            identity.AddProp("Description", false);
            identity.Add(new AuditOption("customattr", "CustomAttribute1-15", "CustomAttribute1-15", false));
            section.AddGroup(identity);

            var filter = new AuditOptionGroup("filter", "Dynamic filter (critical)", GroupMode.MultiCheck); filter.Columns = 1;
            filter.Hint = "RecipientFilter recalculates members live. Do NOT reuse LdapRecipientFilter across environments.";
            filter.AddProp("RecipientFilter", true);
            filter.AddProp("RecipientFilterType", true);
            filter.AddProp("RecipientContainer", true);
            filter.AddProp("IncludedRecipients", true);
            filter.AddProp("ConditionalCompany", false);
            filter.AddProp("ConditionalDepartment", false);
            filter.AddProp("ConditionalStateOrProvince", false);
            filter.Add(new AuditOption("condcustomattr", "ConditionalCustomAttribute1-15", "ConditionalCustomAttribute1-15", false));
            filter.AddProp("LdapRecipientFilter", false);
            section.AddGroup(filter);

            var owner = new AuditOptionGroup("owner", "Owner", GroupMode.MultiCheck); owner.Columns = 2;
            owner.Hint = "ManagedBy resolved from GUID to SMTP. No join/depart restrictions (membership is calculated).";
            owner.AddProp("ManagedBy", true);
            section.AddGroup(owner);

            var moderation = new AuditOptionGroup("moderation", "Moderation", GroupMode.MultiCheck); moderation.Columns = 2;
            moderation.AddProp("ModerationEnabled", true);
            moderation.AddProp("ModeratedBy", true);
            moderation.AddProp("BypassModerationFromSendersOrMembers", false);
            moderation.AddProp("SendModerationNotifications", false);
            section.AddGroup(moderation);

            var restrictions = new AuditOptionGroup("restrictions", "Send restrictions / security", GroupMode.MultiCheck); restrictions.Columns = 2;
            restrictions.AddProp("RequireSenderAuthenticationEnabled", true);
            restrictions.AddProp("AcceptMessagesOnlyFrom", false);
            restrictions.AddProp("AcceptMessagesOnlyFromSendersOrMembers", false);
            restrictions.AddProp("RejectMessagesFromSendersOrMembers", false);
            restrictions.Add(AuditOption.PropScoped("BccBlocked", true, false));
            section.AddGroup(restrictions);

            var ndr = new AuditOptionGroup("ndr", "Delivery reports (NDR)", GroupMode.MultiCheck); ndr.Columns = 3;
            ndr.AddProp("ReportToManagerEnabled", false);
            ndr.AddProp("ReportToOriginatorEnabled", false);
            ndr.AddProp("SendOofMessageToOriginatorEnabled", false);
            section.AddGroup(ndr);

            var permissions = new AuditOptionGroup("permissions", "Permissions", GroupMode.MultiCheck); permissions.Columns = 1;
            permissions.Hint = "Send As (EXO: RecipientPermission / on-prem: ADPermission); Send on Behalf from GrantSendOnBehalfTo. Both resolved to SMTP.";
            permissions.Add(new AuditOption("sendas", "Send As delegates (EXO: RecipientPermission / on-prem: ADPermission)", "sendas", false));
            permissions.Add(new AuditOption("sendonbehalf", "Send on Behalf (GrantSendOnBehalfTo)", "sendonbehalf", false));
            section.AddGroup(permissions);

            var snapshot = new AuditOptionGroup("snapshot", "Members snapshot", GroupMode.SingleChoice);
            snapshot.Hint = "Optional live snapshot via Get-Recipient -RecipientPreviewFilter (documentation, not a static list).";
            snapshot.Columns = 1;
            snapshot.Add(new AuditOption("none", "No snapshot (filter only)", "none", true));
            snapshot.Add(new AuditOption("count", "Add current member-count column", "count", false));
            snapshot.Add(new AuditOption("expand", "Expand current members (one row per member)", "expand", false));
            section.AddGroup(snapshot);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");
                string snap = sel.First("snapshot", "none");
                bool sa = sel.IsSelected("permissions", "sendas");
                bool sob = sel.IsSelected("permissions", "sendonbehalf");

                var chosen = new List<string>();
                foreach (string gk in DynPropGroups)
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0) chosen.Add("DisplayName");

                bool needResolver;
                string selectList = BuildSelectList(chosen, out needResolver);

                var sb = new StringBuilder();
                if (needResolver || sob || snap == "expand") { string getRecip = ConnectionSettings.IsOnline ? "Get-EXORecipient" : "Get-Recipient"; PsScriptHelpers.EmitResolver(sb, getRecip, false); }

                sb.AppendLine("Write-Host 'Querying dynamic distribution groups...'");
                sb.AppendLine("$groups = @(Get-DynamicDistributionGroup -ResultSize " + resultSize + ")");
                sb.AppendLine("Write-Host (\"Retrieved {0} dynamic group(s).\" -f $groups.Count)");
                sb.AppendLine();
                if (sa) EmitSendAsIndex(sb);

                if (snap == "none")
                {
                    if (!sa && !sob)
                    {
                        sb.AppendLine("$rows = $groups | Select-Object " + selectList);
                    }
                    else
                    {
                        sb.AppendLine("$rows = foreach ($g in $groups) {");
                        sb.AppendLine("    $o = $g | Select-Object " + selectList);
                        if (sa) sb.AppendLine("    $o | Add-Member -NotePropertyName SendAs -NotePropertyValue " + SaValueExpr + " -Force");
                        if (sob) sb.AppendLine("    $o | Add-Member -NotePropertyName SendOnBehalf -NotePropertyValue " + SobValueExpr + " -Force");
                        sb.AppendLine("    $o");
                        sb.AppendLine("}");
                    }
                }
                else if (snap == "count")
                {
                    sb.AppendLine("$rows = foreach ($g in $groups) {");
                    sb.AppendLine("    try { $members = @(Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -OrganizationalUnit $g.RecipientContainer -ResultSize Unlimited) } catch { try { $members = @(Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -ResultSize Unlimited) } catch { $members = @() } }");
                    sb.AppendLine("    $o = $g | Select-Object " + selectList);
                    sb.AppendLine("    $o | Add-Member -NotePropertyName CurrentMemberCount -NotePropertyValue $members.Count -Force");
                    if (sa) sb.AppendLine("    $o | Add-Member -NotePropertyName SendAs -NotePropertyValue " + SaValueExpr + " -Force");
                    if (sob) sb.AppendLine("    $o | Add-Member -NotePropertyName SendOnBehalf -NotePropertyValue " + SobValueExpr + " -Force");
                    sb.AppendLine("    $o");
                    sb.AppendLine("}");
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($g in $groups) {");
                    sb.AppendLine("    try { $members = @(Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -OrganizationalUnit $g.RecipientContainer -ResultSize Unlimited) } catch { try { $members = @(Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -ResultSize Unlimited) } catch { $members = @() } }");
                    sb.AppendLine("    $saVal = " + (sa ? SaValueExpr : "''"));
                    sb.AppendLine("    $sobVal = " + (sob ? SobValueExpr : "''"));
                    sb.AppendLine("    if ($members.Count -eq 0) {");
                    sb.AppendLine("        [PSCustomObject]@{ GroupDisplayName=$g.DisplayName; GroupPrimarySmtpAddress=[string]$g.PrimarySmtpAddress; GroupSendAs=$saVal; GroupSendOnBehalf=$sobVal; MemberDisplayName=''; MemberPrimarySmtpAddress=''; MemberRecipientType='' }");
                    sb.AppendLine("    } else {");
                    sb.AppendLine("        foreach ($mem in $members) {");
                    sb.AppendLine("            [PSCustomObject]@{ GroupDisplayName=$g.DisplayName; GroupPrimarySmtpAddress=[string]$g.PrimarySmtpAddress; GroupSendAs=$saVal; GroupSendOnBehalf=$sobVal; MemberDisplayName=$mem.DisplayName; MemberPrimarySmtpAddress=[string]$mem.PrimarySmtpAddress; MemberRecipientType=[string]$mem.RecipientTypeDetails }");
                    sb.AppendLine("        }");
                    sb.AppendLine("    }");
                    sb.AppendLine("}");
                }

                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        private static readonly string[] UnifiedPropGroups =
        {
            "identity", "mailflow", "mailbox", "m365", "sharepoint"
        };

        private static AuditSection BuildUnifiedGroupSection()
        {
            var section = new AuditSection(
                "m365-groups",
                "Microsoft 365 groups",
                "Microsoft 365 group export",
                "Audit Unified (M365) groups. Cloud-only. Optionally flags whether a Teams team is attached.",
                "group",
                AuditScope.ExchangeOnline);
            section.Category = "Groups";
            section.DefaultFileName = "M365Groups.csv";

            var identity = new AuditOptionGroup("identity", "Identity & addressing", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("Name", false);
            identity.AddProp("PrimarySmtpAddress", true);
            identity.AddProp("EmailAddresses", true);
            identity.AddProp("RecipientTypeDetails", true);
            identity.AddProp("Notes", false);
            identity.AddProp("HiddenFromAddressListsEnabled", true);
            identity.AddProp("HiddenFromExchangeClientsEnabled", false);
            identity.AddProp("ExchangeGuid", false);
            identity.Add(new AuditOption("customattr", "CustomAttribute1-15", "CustomAttribute1-15", false));
            identity.Add(new AuditOption("extcustomattr", "ExtensionCustomAttribute1-5", "ExtensionCustomAttribute1-5", false));
            section.AddGroup(identity);

            var teams = new AuditOptionGroup("teams", "Teams & membership", GroupMode.MultiCheck); teams.Columns = 2;
            teams.Hint = "HasTeam is derived from ResourceProvisioningOptions (no Teams module required).";
            teams.Add(new AuditOption("hasteam", "Has Teams team (Yes/No)", "hasteam", true));
            teams.AddProp("GroupMemberCount", false);
            teams.AddProp("GroupExternalMemberCount", false);
            teams.Add(new AuditOption("ManagedBy", "Owners (ManagedBy, resolved)", "ManagedBy", true));
            section.AddGroup(teams);

            var permissions = new AuditOptionGroup("permissions", "Permissions", GroupMode.MultiCheck); permissions.Columns = 1;
            permissions.Hint = "Send As (EXO: RecipientPermission / on-prem: ADPermission); Send on Behalf from GrantSendOnBehalfTo. Both resolved to SMTP.";
            permissions.Add(new AuditOption("sendas", "Send As delegates (EXO: RecipientPermission / on-prem: ADPermission)", "sendas", false));
            permissions.Add(new AuditOption("sendonbehalf", "Send on Behalf (GrantSendOnBehalfTo)", "sendonbehalf", false));
            section.AddGroup(permissions);

            var mailflow = new AuditOptionGroup("mailflow", "Mail behaviour", GroupMode.MultiCheck); mailflow.Columns = 2;
            mailflow.AddProp("RequireSenderAuthenticationEnabled", true);
            mailflow.AddProp("ModerationEnabled", false);
            mailflow.AddProp("ModeratedBy", false);
            mailflow.AddProp("SendModerationNotifications", false);
            mailflow.AddProp("AcceptMessagesOnlyFrom", false);
            mailflow.AddProp("AcceptMessagesOnlyFromSendersOrMembers", false);
            mailflow.AddProp("RejectMessagesFrom", false);
            mailflow.AddProp("RejectMessagesFromSendersOrMembers", false);
            mailflow.AddProp("BypassModerationFromSendersOrMembers", false);
            mailflow.AddProp("BccBlocked", false);
            mailflow.AddProp("ReportToManagerEnabled", false);
            mailflow.AddProp("ReportToOriginatorEnabled", false);
            mailflow.AddProp("SendOofMessageToOriginatorEnabled", false);
            section.AddGroup(mailflow);

            var mailbox = new AuditOptionGroup("mailbox", "Mailbox / content", GroupMode.MultiCheck); mailbox.Columns = 2;
            mailbox.AddProp("MaxSendSize", false);
            mailbox.AddProp("MaxReceiveSize", false);
            mailbox.AddProp("AuditLogAgeLimit", false);
            mailbox.AddProp("InPlaceHolds", false);
            mailbox.AddProp("DataEncryptionPolicy", false);
            section.AddGroup(mailbox);

            var m365 = new AuditOptionGroup("m365", "Microsoft 365 specifics", GroupMode.MultiCheck); m365.Columns = 2;
            m365.Hint = "These have no on-prem equivalent - document, do not migrate 1:1.";
            m365.AddProp("AccessType", true);
            m365.AddProp("AllowAddGuests", false);
            m365.AddProp("AutoSubscribeNewMembers", false);
            m365.AddProp("AlwaysSubscribeMembersToCalendarEvents", false);
            m365.AddProp("WelcomeMessageEnabled", false);
            m365.AddProp("ConnectorsEnabled", false);
            m365.AddProp("SubscriptionEnabled", false);
            m365.AddProp("Language", false);
            m365.AddProp("Classification", false);
            m365.AddProp("SensitivityLabel", false);
            m365.AddProp("InformationBarrierMode", false);
            m365.AddProp("IsMembershipDynamic", false);
            m365.AddProp("ExpirationTime", false);
            section.AddGroup(m365);

            var sharepoint = new AuditOptionGroup("sharepoint", "SharePoint (out of Exchange scope)", GroupMode.MultiCheck); sharepoint.Columns = 1;
            sharepoint.Hint = "URLs present even without Teams - content needs a separate SharePoint migration.";
            sharepoint.AddProp("SharePointSiteUrl", false);
            sharepoint.AddProp("SharePointDocumentsUrl", false);
            sharepoint.AddProp("SharePointNotebookUrl", false);
            section.AddGroup(sharepoint);

            var membership = new AuditOptionGroup("membership", "Membership (Get-UnifiedGroupLinks)", GroupMode.SingleChoice);
            membership.Hint = "Expanding produces one row per group+member/owner; resolved to name/SMTP.";
            membership.Columns = 1;
            membership.Add(new AuditOption("groups", "Group rows only", "groups", true));
            membership.Add(new AuditOption("expandmembers", "Expand members (one row per member)", "expandmembers", false));
            membership.Add(new AuditOption("expandowners", "Expand owners (one row per owner)", "expandowners", false));
            membership.Add(new AuditOption("expandsubscribers", "Expand subscribers (one row per subscriber)", "expandsubscribers", false));
            section.AddGroup(membership);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string resultSize = sel.First("size", "Unlimited");
                string membershipMode = sel.First("membership", "groups");

                var chosen = new List<string>();
                bool hasTeam = sel.IsSelected("teams", "hasteam");
                bool sa = sel.IsSelected("permissions", "sendas");
                bool sob = sel.IsSelected("permissions", "sendonbehalf");
                foreach (string v in sel.Selected("teams"))
                    if (v != "hasteam" && !chosen.Contains(v)) chosen.Add(v);
                foreach (string gk in UnifiedPropGroups)
                    foreach (string v in sel.Selected(gk))
                        if (!chosen.Contains(v)) chosen.Add(v);
                if (chosen.Count == 0 && !hasTeam) chosen.Add("DisplayName");

                bool needResolver;
                string selectList = BuildSelectList(chosen, out needResolver, hasTeam);

                string linkType = null;
                if (membershipMode == "expandmembers") linkType = "Members";
                else if (membershipMode == "expandowners") linkType = "Owners";
                else if (membershipMode == "expandsubscribers") linkType = "Subscribers";

                var sb = new StringBuilder();
                if (needResolver || sob) { string getRecip = ConnectionSettings.IsOnline ? "Get-EXORecipient" : "Get-Recipient"; PsScriptHelpers.EmitResolver(sb, getRecip, false); }

                sb.AppendLine("Write-Host 'Querying Microsoft 365 groups...'");
                sb.AppendLine("$groups = @(Get-UnifiedGroup -ResultSize " + resultSize + ")");
                sb.AppendLine("Write-Host (\"Retrieved {0} M365 group(s).\" -f $groups.Count)");
                sb.AppendLine();
                if (sa) EmitSendAsIndex(sb);

                if (linkType == null)
                {
                    if (!sa && !sob)
                    {
                        sb.AppendLine("$rows = $groups | Select-Object " + selectList);
                    }
                    else
                    {
                        sb.AppendLine("$rows = foreach ($g in $groups) {");
                        sb.AppendLine("    $o = $g | Select-Object " + selectList);
                        if (sa) sb.AppendLine("    $o | Add-Member -NotePropertyName SendAs -NotePropertyValue " + SaValueExpr + " -Force");
                        if (sob) sb.AppendLine("    $o | Add-Member -NotePropertyName SendOnBehalf -NotePropertyValue " + SobValueExpr + " -Force");
                        sb.AppendLine("    $o");
                        sb.AppendLine("}");
                    }
                }
                else
                {
                    sb.AppendLine("$rows = foreach ($g in $groups) {");
                    sb.AppendLine("    try { $links = @(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType " + linkType + " -ResultSize Unlimited) } catch { $links = @() }");
                    sb.AppendLine("    $saVal = " + (sa ? SaValueExpr : "''"));
                    sb.AppendLine("    $sobVal = " + (sob ? SobValueExpr : "''"));
                    sb.AppendLine("    if ($links.Count -eq 0) {");
                    sb.AppendLine("        [PSCustomObject]@{ GroupDisplayName=$g.DisplayName; GroupPrimarySmtpAddress=[string]$g.PrimarySmtpAddress; GroupSendAs=$saVal; GroupSendOnBehalf=$sobVal; LinkType='" + linkType + "'; MemberDisplayName=''; MemberPrimarySmtpAddress=''; MemberRecipientType='' }");
                    sb.AppendLine("    } else {");
                    sb.AppendLine("        foreach ($mem in $links) {");
                    sb.AppendLine("            [PSCustomObject]@{ GroupDisplayName=$g.DisplayName; GroupPrimarySmtpAddress=[string]$g.PrimarySmtpAddress; GroupSendAs=$saVal; GroupSendOnBehalf=$sobVal; LinkType='" + linkType + "'; MemberDisplayName=$mem.DisplayName; MemberPrimarySmtpAddress=[string]$mem.PrimarySmtpAddress; MemberRecipientType=[string]$mem.RecipientTypeDetails }");
                    sb.AppendLine("        }");
                    sb.AppendLine("    }");
                    sb.AppendLine("}");
                }

                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        private static void EmitSendAsIndex(StringBuilder sb)
        {
            bool online = ConnectionSettings.IsOnline;
            sb.AppendLine("Write-Host 'Indexing Send As (bulk)...'");
            sb.AppendLine("$saIndex = @{}");
            if (online)
            {
                // Exchange Online: Send As is exposed via (EXO)RecipientPermission.
                sb.AppendLine("try { $groups | Get-EXORecipientPermission -ErrorAction SilentlyContinue |");
                sb.AppendLine("    Where-Object { $_.AccessRights -contains 'SendAs' -and $_.Trustee -notlike 'NT AUTHORITY\\*' } |");
                sb.AppendLine("    ForEach-Object { $k=[string]$_.Identity; if ($saIndex.ContainsKey($k)) { $saIndex[$k]=$saIndex[$k]+','+[string]$_.Trustee } else { $saIndex[$k]=[string]$_.Trustee } } } catch { }");
            }
            else
            {
                // On-premises: Get-RecipientPermission does not exist. Send As lives in
                // Active Directory as the 'Send-As' extended right (Get-ADPermission).
                sb.AppendLine("foreach ($__g in $groups) {");
                sb.AppendLine("    $k = [string]$__g.Identity");
                sb.AppendLine("    $__t = Get-ADPermission -Identity $__g.Identity -ErrorAction SilentlyContinue |");
                sb.AppendLine("        Where-Object { ($_.ExtendedRights -like '*Send-As*') -and (-not $_.IsInherited) -and (-not $_.Deny) -and ($_.User -notlike 'NT AUTHORITY\\*') } |");
                sb.AppendLine("        ForEach-Object { [string]$_.User }");
                sb.AppendLine("    if ($__t) { $saIndex[$k] = (@($__t) -join ',') }");
                sb.AppendLine("}");
            }
            sb.AppendLine();
        }

        private const string SaValueExpr = "($(if ($saIndex.ContainsKey([string]$g.Identity)) { $saIndex[[string]$g.Identity] } else { '' }))";
        private const string SobValueExpr = "(Resolve-Recip $g.GrantSendOnBehalfTo)";

        private static string BuildSelectList(List<string> chosen, out bool needResolver)
        {
            return BuildSelectList(chosen, out needResolver, false);
        }

        private static string BuildSelectList(List<string> chosen, out bool needResolver, bool hasTeam)
        {
            needResolver = false;
            var exprs = new List<string>();
            if (hasTeam)
                exprs.Add("@{Name='HasTeam';Expression={ if ((@($_.ResourceProvisioningOptions) -contains 'Team')) { 'Yes' } else { 'No' } }}");
            foreach (string v in chosen)
            {
                if (v == "EmailAddresses")
                    exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                else if (RecipientRefs.Contains(v)) { needResolver = true; exprs.Add("@{Name='" + v + "';Expression={ Resolve-Recip $_." + v + " }}"); }
                else if (MultiValued.Contains(v))
                    exprs.Add("@{Name='" + v + "';Expression={($_." + v + " | ForEach-Object { [string]$_ }) -join ','}}");
                else if (v == "CustomAttribute1-15")
                    for (int i = 1; i <= 15; i++) exprs.Add("CustomAttribute" + i);
                else if (v == "ExtensionCustomAttribute1-5")
                    for (int i = 1; i <= 5; i++)
                        exprs.Add("@{Name='ExtensionCustomAttribute" + i + "';Expression={($_.ExtensionCustomAttribute" + i + " | ForEach-Object { [string]$_ }) -join ','}}");
                else if (v == "ConditionalCustomAttribute1-15")
                    for (int i = 1; i <= 15; i++)
                        exprs.Add("@{Name='ConditionalCustomAttribute" + i + "';Expression={($_.ConditionalCustomAttribute" + i + " | ForEach-Object { [string]$_ }) -join ','}}");
                else exprs.Add(v);
            }
            if (exprs.Count == 0) exprs.Add("DisplayName");
            return string.Join(", ", exprs.ToArray());
        }
    }
}
