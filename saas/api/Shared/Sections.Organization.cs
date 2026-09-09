using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "Organization" category: sharing/federation/org-config singletons plus
    // email address policies. Every Get-* below exists on Exchange on-premises
    // AND Exchange Online (read-only); property surfaces differ slightly, so
    // only widely available properties are curated.
    internal static class SectionsOrganization
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildOrgSharingSection());
            AuditRegistry.Register(BuildAddressPoliciesSection());
            AuditRegistry.Register(BuildJournalRulesSection());
            AuditRegistry.Register(BuildRetentionPoliciesSection());
            AuditRegistry.Register(BuildAddressBooksSection());
            AuditRegistry.Register(BuildCertificatesSection());
            AuditRegistry.Register(BuildOwaPolicySection());
            AuditRegistry.Register(BuildRolePoliciesSection());
        }

        // Properties that are collections -> joined with ',' so a ';' CSV never clashes.
        private static readonly List<string> MultiValued = new List<string>(new string[]
        {
            "Domains", "DomainNames",
            "EnabledEmailAddressTemplates", "DisabledEmailAddressTemplates",
            "RetentionPolicyTagLinks", "AddressLists", "AssignedRoles"
        });

        private static string BuildSelectList(List<string> chosen)
        {
            var exprs = new List<string>();
            foreach (string v in chosen)
            {
                if (MultiValued.Contains(v))
                    exprs.Add("@{Name='" + v + "';Expression={($_." + v + " | ForEach-Object { [string]$_ }) -join ','}}");
                else exprs.Add(v);
            }
            if (exprs.Count == 0) exprs.Add("Name");
            return string.Join(", ", exprs.ToArray());
        }

        private static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            return PsScriptHelpers.Collect(sel, groupKeys);
        }

        // ============================================================ 1. ORG SHARING
        private static AuditSection BuildOrgSharingSection()
        {
            var section = new AuditSection(
                "org-sharing",
                "Sharing & org",
                "Sharing & organization export",
                "Audit coexistence essentials (sharing, relationships, federation, org config). One object per run.",
                "shield",
                AuditScope.Both);
            section.Category = "Organization";
            section.DefaultFileName = "OrgSharing.csv";

            var obj = new AuditOptionGroup("object", "Object", GroupMode.SingleChoice); obj.Columns = 1;
            obj.Add(new AuditOption("sharing", "Sharing policies (Get-SharingPolicy)", "sharing", true));
            obj.Add(new AuditOption("relationships", "Organization relationships (Get-OrganizationRelationship)", "relationships", false));
            obj.Add(new AuditOption("federation", "Federation trust (Get-FederationTrust)", "federation", false));
            obj.Add(new AuditOption("fedorgid", "Federated organization ID (Get-FederatedOrganizationIdentifier)", "fedorgid", false));
            obj.Add(new AuditOption("orgconfig", "Organization config essentials (Get-OrganizationConfig)", "orgconfig", false));
            section.AddGroup(obj);

            var sharing = new AuditOptionGroup("sharing", "Sharing policy properties", GroupMode.MultiCheck); sharing.Columns = 2;
            sharing.AddProp("Name", true);
            sharing.AddProp("Enabled", true);
            sharing.AddProp("Default", true);
            sharing.AddProp("Domains", true);
            section.AddGroup(sharing);

            var rel = new AuditOptionGroup("relationships", "Relationship properties", GroupMode.MultiCheck); rel.Columns = 2;
            rel.AddProp("Name", false);
            rel.AddProp("Enabled", false);
            rel.AddProp("DomainNames", false);
            rel.AddProp("FreeBusyAccessEnabled", false);
            rel.AddProp("FreeBusyAccessLevel", false);
            rel.AddProp("MailboxMoveEnabled", false);
            rel.AddProp("DeliveryReportEnabled", false);
            rel.AddProp("TargetApplicationUri", false);
            rel.AddProp("TargetSharingEpr", false);
            rel.AddProp("TargetOwaURL", false);
            section.AddGroup(rel);

            var fed = new AuditOptionGroup("federation", "Federation properties", GroupMode.MultiCheck); fed.Columns = 2;
            fed.AddProp("Name", false);
            fed.AddProp("ApplicationUri", false);
            fed.AddProp("TokenIssuerUri", false);
            fed.AddProp("AccountNamespace", false);
            fed.AddProp("DelegationTrustLink", false);
            section.AddGroup(fed);

            var org = new AuditOptionGroup("orgconfig", "Org config properties", GroupMode.MultiCheck); org.Columns = 2;
            org.AddProp("Name", false);
            org.AddProp("PublicFoldersEnabled", false);
            org.AddProp("ElcProcessingDisabled", false);
            org.AddProp("MailTipsAllTipsEnabled", false);
            org.AddProp("MailTipsExternalRecipientsTipsEnabled", false);
            org.AddProp("ConnectorsEnabled", false);
            org.AddProp("DefaultPublicFolderAgeLimit", false);
            org.AddProp("PublicFolderShowClientControl", false);
            org.AddProp("OAuth2ClientProfileEnabled", false);
            section.AddGroup(org);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string which = sel.First("object", "sharing");
                string cmdlet;
                string propGroup;
                string label;
                if (which == "relationships") { cmdlet = "Get-OrganizationRelationship"; propGroup = "relationships"; label = "organization relationship(s)"; }
                else if (which == "federation") { cmdlet = "Get-FederationTrust"; propGroup = "federation"; label = "federation trust(s)"; }
                else if (which == "fedorgid") { cmdlet = "Get-FederatedOrganizationIdentifier"; propGroup = "federation"; label = "federated organization ID(s)"; }
                else if (which == "orgconfig") { cmdlet = "Get-OrganizationConfig"; propGroup = "orgconfig"; label = "organization config(s)"; }
                else { cmdlet = "Get-SharingPolicy"; propGroup = "sharing"; label = "sharing policies"; }

                var chosen = Collect(sel, propGroup);
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying " + label + "...'");
                sb.AppendLine("$items = @(" + cmdlet + " -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} " + label + ".\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 2. ADDRESS POLICIES
        private static AuditSection BuildAddressPoliciesSection()
        {
            var section = new AuditSection(
                "address-policies",
                "Address policies",
                "Address policy export",
                "Audit email address policies (Get-EmailAddressPolicy): filters and templates must exist in the target.",
                "rule",
                AuditScope.Both);
            section.Category = "Domains / Routing";
            section.DefaultFileName = "AddressPolicies.csv";

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("Name", true);
            identity.AddProp("Priority", true);
            identity.AddProp("Enabled", true);
            section.AddGroup(identity);

            var filter = new AuditOptionGroup("filter", "Recipient filter", GroupMode.MultiCheck); filter.Columns = 2;
            filter.Hint = "RecipientFilter recalculates membership live. Do NOT reuse LdapRecipientFilter across environments.";
            filter.AddProp("RecipientFilterType", false);
            filter.AddProp("RecipientContainer", false);
            filter.AddProp("IncludedRecipients", false);
            filter.AddProp("ConditionalCompany", false);
            filter.AddProp("ConditionalDepartment", false);
            filter.AddProp("ConditionalStateOrProvince", false);
            filter.AddProp("ConditionalCustomAttribute1", false);
            filter.AddProp("ConditionalCustomAttribute2", false);
            filter.AddProp("ConditionalCustomAttribute3", false);
            filter.AddProp("ConditionalCustomAttribute4", false);
            filter.AddProp("ConditionalCustomAttribute5", false);
            filter.AddProp("LdapRecipientFilter", false);
            section.AddGroup(filter);

            var templates = new AuditOptionGroup("templates", "Address templates", GroupMode.MultiCheck); templates.Columns = 1;
            templates.AddProp("EnabledEmailAddressTemplates", true);
            templates.AddProp("DisabledEmailAddressTemplates", false);
            templates.AddProp("EnabledPrimarySMTPAddressTemplate", false);
            section.AddGroup(templates);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "filter", "templates");
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying email address policies...'");
                sb.AppendLine("$items = @(Get-EmailAddressPolicy -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} address policies.\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 3. JOURNAL RULES
        private static AuditSection BuildJournalRulesSection()
        {
            var section = new AuditSection(
                "journal-rules",
                "Journal rules",
                "Journal rule export",
                "Audit journal rules (Get-JournalRule): journaling must keep working after migration.",
                "shield",
                AuditScope.Both);
            section.Category = "Organization";
            section.DefaultFileName = "JournalRules.csv";

            var props = new AuditOptionGroup("props", "Properties", GroupMode.MultiCheck); props.Columns = 2;
            props.AddProp("Name", true);
            props.AddProp("Enabled", true);
            props.AddProp("JournalEmailAddress", true);
            props.AddProp("Scope", true);
            props.AddProp("Recipient", false);
            section.AddGroup(props);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "props");
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying journal rules...'");
                sb.AppendLine("$items = @(Get-JournalRule -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} journal rule(s).\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 4. RETENTION POLICIES
        private static AuditSection BuildRetentionPoliciesSection()
        {
            var section = new AuditSection(
                "retention-policies",
                "Retention policies",
                "Retention policy export",
                "Audit retention policies and tags (Get-RetentionPolicy / Get-RetentionPolicyTag): MRM must be rebuilt in the target.",
                "shield",
                AuditScope.Both);
            section.Category = "Organization";
            section.DefaultFileName = "RetentionPolicies.csv";

            var obj = new AuditOptionGroup("object", "Object", GroupMode.SingleChoice); obj.Columns = 1;
            obj.Add(new AuditOption("policies", "Retention policies (Get-RetentionPolicy)", "policies", true));
            obj.Add(new AuditOption("tags", "Retention tags (Get-RetentionPolicyTag)", "tags", false));
            section.AddGroup(obj);

            var policy = new AuditOptionGroup("policy", "Policy properties", GroupMode.MultiCheck); policy.Columns = 2;
            policy.AddProp("Name", true);
            policy.AddProp("RetentionPolicyTagLinks", true);
            section.AddGroup(policy);

            var tag = new AuditOptionGroup("tag", "Tag properties", GroupMode.MultiCheck); tag.Columns = 2;
            tag.AddProp("Name", false);
            tag.AddProp("Type", false);
            tag.AddProp("RetentionEnabled", false);
            tag.AddProp("AgeLimitForRetention", false);
            tag.AddProp("RetentionAction", false);
            section.AddGroup(tag);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string which = sel.First("object", "policies");
                string cmdlet;
                string propGroup;
                string label;
                if (which == "tags") { cmdlet = "Get-RetentionPolicyTag"; propGroup = "tag"; label = "retention tags"; }
                else { cmdlet = "Get-RetentionPolicy"; propGroup = "policy"; label = "retention policies"; }

                var chosen = Collect(sel, propGroup);
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying " + label + "...'");
                sb.AppendLine("$items = @(" + cmdlet + " -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} " + label + ".\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 5. ADDRESS BOOKS
        private static AuditSection BuildAddressBooksSection()
        {
            var section = new AuditSection(
                "address-books",
                "Address books",
                "Address book export",
                "Audit offline address books and address lists (Get-OfflineAddressBook / Get-AddressList): OAB must be regenerated after migration.",
                "group",
                AuditScope.OnPremises);
            section.Category = "Organization";
            section.DefaultFileName = "AddressBooks.csv";

            var obj = new AuditOptionGroup("object", "Object", GroupMode.SingleChoice); obj.Columns = 1;
            obj.Add(new AuditOption("oab", "Offline address books (Get-OfflineAddressBook)", "oab", true));
            obj.Add(new AuditOption("lists", "Address lists (Get-AddressList)", "lists", false));
            section.AddGroup(obj);

            var oab = new AuditOptionGroup("oab", "OAB properties", GroupMode.MultiCheck); oab.Columns = 2;
            oab.AddProp("Name", true);
            oab.AddProp("IsDefault", true);
            oab.AddProp("AddressLists", true);
            oab.AddProp("GeneratingMailbox", false);
            section.AddGroup(oab);

            var lists = new AuditOptionGroup("lists", "Address list properties", GroupMode.MultiCheck); lists.Columns = 2;
            lists.AddProp("Name", false);
            lists.AddProp("DisplayName", false);
            lists.AddProp("RecipientFilter", false);
            lists.AddProp("RecipientContainer", false);
            lists.AddProp("IncludedRecipients", false);
            lists.AddProp("LdapRecipientFilter", false);
            section.AddGroup(lists);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string which = sel.First("object", "oab");
                string cmdlet;
                string propGroup;
                string label;
                if (which == "lists") { cmdlet = "Get-AddressList"; propGroup = "lists"; label = "address lists"; }
                else { cmdlet = "Get-OfflineAddressBook"; propGroup = "oab"; label = "offline address books"; }

                var chosen = Collect(sel, propGroup);
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying " + label + "...'");
                sb.AppendLine("$items = @(" + cmdlet + " -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} " + label + ".\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 6. CERTIFICATES
        private static AuditSection BuildCertificatesSection()
        {
            var section = new AuditSection(
                "certificates",
                "Certificates",
                "Certificate export",
                "Audit Exchange certificates (Get-ExchangeCertificate): expired federation/SMTP certs are the classic cutover-day surprise.",
                "shield",
                AuditScope.OnPremises);
            section.Category = "Organization";
            section.DefaultFileName = "Certificates.csv";

            var identity = new AuditOptionGroup("identity", "Identity & validity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("Thumbprint", true);
            identity.AddProp("Subject", true);
            identity.AddProp("Issuer", true);
            identity.AddProp("IsSelfSigned", true);
            identity.AddProp("NotAfter", true);
            section.AddGroup(identity);

            var usage = new AuditOptionGroup("usage", "Services & domains", GroupMode.MultiCheck); usage.Columns = 2;
            usage.AddProp("Services", true);
            usage.AddProp("CertificateDomains", true);
            section.AddGroup(usage);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "usage");
                if (chosen.Count == 0) chosen.Add("Thumbprint");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Exchange certificates...'");
                sb.AppendLine("$items = @(Get-ExchangeCertificate -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} certificate(s).\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 7. OWA POLICY
        private static AuditSection BuildOwaPolicySection()
        {
            var section = new AuditSection(
                "owa-policy",
                "OWA policy",
                "OWA policy export",
                "Audit Outlook on the web policies (Get-OwaMailboxPolicy): feature toggles users notice on day one.",
                "rule",
                AuditScope.Both);
            section.Category = "Organization";
            section.DefaultFileName = "OwaPolicy.csv";

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("Name", true);
            identity.AddProp("IsDefault", true);
            section.AddGroup(identity);

            var features = new AuditOptionGroup("features", "Feature toggles", GroupMode.MultiCheck); features.Columns = 2;
            features.AddProp("DirectFileAccessOnPublicComputersEnabled", true);
            features.AddProp("DirectFileAccessOnPrivateComputersEnabled", true);
            features.AddProp("WebReadyDocumentViewingOnPublicComputersEnabled", false);
            features.AddProp("WebReadyDocumentViewingOnPrivateComputersEnabled", false);
            features.AddProp("OfflineAccessEnabled", false);
            features.AddProp("InstantMessagingEnabled", false);
            features.AddProp("CalendarEnabled", false);
            features.AddProp("ContactsEnabled", false);
            features.AddProp("TasksEnabled", false);
            features.AddProp("NotesEnabled", false);
            features.AddProp("JournalEnabled", false);
            features.AddProp("RemindersAndNotificationsEnabled", false);
            features.AddProp("SearchFoldersEnabled", false);
            features.AddProp("SignaturesEnabled", false);
            features.AddProp("ThemeSelectionEnabled", false);
            features.AddProp("SetPhotoEnabled", false);
            features.AddProp("TextMessagingEnabled", false);
            features.AddProp("WssAccessOnPublicComputersEnabled", false);
            features.AddProp("WssAccessOnPrivateComputersEnabled", false);
            features.AddProp("ReportJunkEmailEnabled", false);
            features.AddProp("UserVoiceEnabled", false);
            section.AddGroup(features);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "features");
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying OWA mailbox policies...'");
                sb.AppendLine("$items = @(Get-OwaMailboxPolicy -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} OWA policies.\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 8. ROLE POLICIES
        private static AuditSection BuildRolePoliciesSection()
        {
            var section = new AuditSection(
                "role-policies",
                "Role policies",
                "Role policy export",
                "Audit RBAC end-user surface (Get-RoleAssignmentPolicy / Get-ManagementRoleAssignment).",
                "shield",
                AuditScope.Both);
            section.Category = "Organization";
            section.DefaultFileName = "RolePolicies.csv";

            var obj = new AuditOptionGroup("object", "Object", GroupMode.SingleChoice); obj.Columns = 1;
            obj.Add(new AuditOption("policies", "Assignment policies (Get-RoleAssignmentPolicy)", "policies", true));
            obj.Add(new AuditOption("assignments", "Role assignments (Get-ManagementRoleAssignment)", "assignments", false));
            section.AddGroup(obj);

            var policy = new AuditOptionGroup("policy", "Policy properties", GroupMode.MultiCheck); policy.Columns = 2;
            policy.AddProp("Name", true);
            policy.AddProp("IsDefault", true);
            policy.AddProp("Description", false);
            policy.AddProp("AssignedRoles", true);
            section.AddGroup(policy);

            var assign = new AuditOptionGroup("assignments", "Assignment properties", GroupMode.MultiCheck); assign.Columns = 2;
            assign.AddProp("Name", false);
            assign.AddProp("Role", false);
            assign.AddProp("RoleAssigneeName", false);
            assign.AddProp("RoleAssigneeType", false);
            assign.AddProp("Enabled", false);
            section.AddGroup(assign);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string which = sel.First("object", "policies");
                string cmdlet;
                string propGroup;
                string label;
                if (which == "assignments") { cmdlet = "Get-ManagementRoleAssignment"; propGroup = "assignments"; label = "role assignments"; }
                else { cmdlet = "Get-RoleAssignmentPolicy"; propGroup = "policy"; label = "role assignment policies"; }

                var chosen = Collect(sel, propGroup);
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying " + label + "...'");
                sb.AppendLine("$items = @(" + cmdlet + " -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} " + label + ".\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }
    }
}
