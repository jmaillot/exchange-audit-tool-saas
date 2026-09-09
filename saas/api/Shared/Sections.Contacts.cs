using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    internal static class SectionsContacts
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildMailUserSection());
            AuditRegistry.Register(BuildMailContactSection());
        }

        private static readonly List<string> RecipRefs = new List<string>(new string[]
        {
            "AcceptMessagesOnlyFromSendersOrMembers", "AcceptMessagesOnlyFromDLMembers",
            "RejectMessagesFromSendersOrMembers", "RejectMessagesFromDLMembers", "ModeratedBy"
        });

        private static readonly List<string> MultiValued = new List<string>(new string[]
        {
            "EmailAddresses", "AddressListMembership"
        });

        private static readonly List<string> ContactCardProps = new List<string>(new string[]
        {
            "FirstName", "LastName", "Initials", "Title", "Department", "Company", "AssistantName",
            "Phone", "MobilePhone", "Fax", "HomePhone", "OtherTelephone", "Pager",
            "StreetAddress", "City", "StateOrProvince", "PostalCode", "CountryOrRegion", "WebPage", "Notes"
        });

        private static AuditSection BuildMailUserSection()
        {
            var section = new AuditSection(
                "mail-users",
                "Mail users",
                "Mail user export",
                "Audit mail users (have a directory account: UPN, external routing, sign-in state).",
                "group",
                AuditScope.Both);
            section.Category = "Contacts";
            section.DefaultFileName = "MailUsers.csv";

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("Name", false);
            identity.AddProp("UserPrincipalName", true);
            identity.AddProp("PrimarySmtpAddress", true);
            section.AddGroup(identity);

            var routing = new AuditOptionGroup("routing", "External routing (key)", GroupMode.MultiCheck); routing.Columns = 1;
            routing.Hint = "ExternalEmailAddress is the MailUser-specific routing target, distinct from PrimarySmtpAddress.";
            routing.AddProp("ExternalEmailAddress", true);
            section.AddGroup(routing);

            var type = new AuditOptionGroup("type", "Type", GroupMode.MultiCheck); type.Columns = 2;
            type.AddProp("RecipientType", false);
            section.AddGroup(type);

            var account = new AuditOptionGroup("account", "Account / limits", GroupMode.MultiCheck); account.Columns = 2;
            account.AddProp("AccountDisabled", true);
            account.AddProp("MaxSendSize", false);
            account.AddProp("MaxReceiveSize", false);
            account.AddProp("RecipientLimits", false);
            section.AddGroup(account);

            AddCommonRecipientGroups(section);
            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                return BuildRecipientScript(sel, ctx, "Get-MailUser", "mail user", CommonMailUserGroups());
            };
            return section;
        }

        private static AuditSection BuildMailContactSection()
        {
            var section = new AuditSection(
                "mail-contacts",
                "Mail contacts",
                "Mail contact export",
                "Audit mail contacts (pure GAL entries, no account) - maps to contact-card fields.",
                "group",
                AuditScope.Both);
            section.Category = "Contacts";
            section.DefaultFileName = "MailContacts.csv";

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("Alias", true);
            identity.AddProp("Name", false);
            identity.AddProp("PrimarySmtpAddress", true);
            section.AddGroup(identity);

            var routing = new AuditOptionGroup("routing", "External routing (key)", GroupMode.MultiCheck); routing.Columns = 1;
            routing.Hint = "ExternalEmailAddress is the coexistence routing target, distinct from PrimarySmtpAddress.";
            routing.AddProp("ExternalEmailAddress", true);
            section.AddGroup(routing);

            var type = new AuditOptionGroup("type", "Type", GroupMode.MultiCheck); type.Columns = 2;
            type.AddProp("RecipientType", false);
            section.AddGroup(type);

            var nameparts = new AuditOptionGroup("nameparts", "Name parts (Get-Contact)", GroupMode.MultiCheck); nameparts.Columns = 3;
            nameparts.AddProp("FirstName", false);
            nameparts.AddProp("LastName", false);
            nameparts.AddProp("Initials", false);
            section.AddGroup(nameparts);

            var org = new AuditOptionGroup("org", "Organization (Get-Contact)", GroupMode.MultiCheck); org.Columns = 2;
            org.AddProp("Title", false);
            org.AddProp("Department", false);
            org.AddProp("Company", false);
            org.Add(new AuditOption("Manager", "Manager (resolved to name)", "Manager", false));
            org.AddProp("AssistantName", false);
            section.AddGroup(org);

            var phone = new AuditOptionGroup("phone", "Phone (Get-Contact)", GroupMode.MultiCheck); phone.Columns = 3;
            phone.AddProp("Phone", false);
            phone.AddProp("MobilePhone", false);
            phone.AddProp("Fax", false);
            phone.AddProp("HomePhone", false);
            phone.AddProp("OtherTelephone", false);
            phone.AddProp("Pager", false);
            section.AddGroup(phone);

            var postal = new AuditOptionGroup("postal", "Postal address (Get-Contact)", GroupMode.MultiCheck); postal.Columns = 2;
            postal.AddProp("StreetAddress", false);
            postal.AddProp("City", false);
            postal.AddProp("StateOrProvince", false);
            postal.AddProp("PostalCode", false);
            postal.AddProp("CountryOrRegion", false);
            postal.AddProp("WebPage", false);
            postal.AddProp("Notes", false);
            section.AddGroup(postal);

            AddCommonRecipientGroups(section);
            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                return BuildRecipientScript(sel, ctx, "Get-MailContact", "mail contact", CommonMailContactGroups());
            };
            return section;
        }

        private static void AddCommonRecipientGroups(AuditSection section)
        {
            var visibility = new AuditOptionGroup("visibility", "Visibility", GroupMode.MultiCheck); visibility.Columns = 2;
            visibility.AddProp("HiddenFromAddressListsEnabled", false);
            visibility.AddProp("AddressListMembership", false);
            section.AddGroup(visibility);

            var mailflow = new AuditOptionGroup("mailflow", "Mail flow restrictions", GroupMode.MultiCheck); mailflow.Columns = 1;
            mailflow.Hint = "RequireSenderAuthenticationEnabled=True blocks external mail - easy to miss.";
            mailflow.AddProp("RequireSenderAuthenticationEnabled", true);
            mailflow.AddProp("AcceptMessagesOnlyFromSendersOrMembers", false);
            mailflow.AddProp("AcceptMessagesOnlyFromDLMembers", false);
            mailflow.AddProp("RejectMessagesFromSendersOrMembers", false);
            mailflow.AddProp("RejectMessagesFromDLMembers", false);
            section.AddGroup(mailflow);

            var moderation = new AuditOptionGroup("moderation", "Moderation", GroupMode.MultiCheck); moderation.Columns = 2;
            moderation.AddProp("ModerationEnabled", false);
            moderation.AddProp("ModeratedBy", false);
            section.AddGroup(moderation);

            var custom = new AuditOptionGroup("custom", "Custom attributes", GroupMode.MultiCheck); custom.Columns = 2;
            custom.Add(new AuditOption("customattr", "CustomAttribute1-15", "CustomAttribute1-15", false));
            custom.Add(new AuditOption("extcustomattr", "ExtensionCustomAttribute1-5", "ExtensionCustomAttribute1-5", false));
            section.AddGroup(custom);

            var addressing = new AuditOptionGroup("addressing", "Addressing", GroupMode.MultiCheck); addressing.Columns = 2;
            addressing.Hint = "EmailAddresses filtered to smtp: entries.";
            addressing.AddProp("EmailAddresses", false);
            section.AddGroup(addressing);
        }

        private static string[] CommonMailUserGroups()
        {
            return new string[] { "identity", "routing", "type", "account", "visibility", "mailflow", "moderation", "custom", "addressing" };
        }

        private static string[] CommonMailContactGroups()
        {
            return new string[] { "identity", "routing", "type", "nameparts", "org", "phone", "postal", "visibility", "mailflow", "moderation", "custom", "addressing" };
        }

        private static string BuildRecipientScript(AuditSelection sel, ScriptContext ctx, string baseCmdlet, string label, string[] groupKeys)
        {
            string resultSize = sel.First("size", "Unlimited");

            var chosen = new List<string>();
            foreach (string gk in groupKeys)
                foreach (string v in sel.Selected(gk))
                    if (!chosen.Contains(v)) chosen.Add(v);
            if (chosen.Count == 0) chosen.Add("DisplayName");

            bool needResolver = false;
            bool needManager = false;
            var contactProps = new List<string>();
            var exprs = new List<string>();
            foreach (string v in chosen)
            {
                if (v == "Manager") { needManager = true; needResolver = true; continue; }
                if (ContactCardProps.Contains(v)) { contactProps.Add(v); continue; }
                if (v == "EmailAddresses")
                    exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                else if (RecipRefs.Contains(v)) { needResolver = true; exprs.Add("@{Name='" + v + "';Expression={ Resolve-Recip $_." + v + " }}"); }
                else if (MultiValued.Contains(v))
                    exprs.Add("@{Name='" + v + "';Expression={($_." + v + " | ForEach-Object { [string]$_ }) -join ','}}");
                else if (v == "CustomAttribute1-15")
                    for (int i = 1; i <= 15; i++) exprs.Add("CustomAttribute" + i);
                else if (v == "ExtensionCustomAttribute1-5")
                    for (int i = 1; i <= 5; i++)
                        exprs.Add("@{Name='ExtensionCustomAttribute" + i + "';Expression={($_.ExtensionCustomAttribute" + i + " | ForEach-Object { [string]$_ }) -join ','}}");
                else exprs.Add(v);
            }
            if (exprs.Count == 0) exprs.Add("DisplayName");
            string selectList = string.Join(", ", exprs.ToArray());

            bool contactData = contactProps.Count > 0;
            bool perRow = contactData || needManager;

            bool online = ConnectionSettings.IsOnline;
            string getRecip = online ? "Get-EXORecipient" : "Get-Recipient";

            var sb = new StringBuilder();

            if (needResolver) PsScriptHelpers.EmitResolver(sb, getRecip, true);

            sb.AppendLine("Write-Host 'Querying " + label + "s...'");
            sb.AppendLine("$items = @(" + baseCmdlet + " -ResultSize " + resultSize + ")");
            sb.AppendLine("Write-Host (\"Retrieved {0} " + label + "(s).\" -f $items.Count)");
            sb.AppendLine();

            if (contactData || needManager)
            {
                sb.AppendLine("Write-Host 'Indexing Get-Contact (bulk)...'");
                sb.AppendLine("$contactIndex = @{}");
                sb.AppendLine("Get-Contact -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object { $contactIndex[[string]$_.Guid] = $_ }");
                sb.AppendLine();
            }

            if (!perRow)
            {
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
            }
            else
            {
                sb.AppendLine("$rows = foreach ($m in $items) {");
                sb.AppendLine("    $obj = $m | Select-Object " + selectList);
                sb.AppendLine("    $c = $null; $ckey = [string]$m.Guid; if ($contactIndex.ContainsKey($ckey)) { $c = $contactIndex[$ckey] }");
                foreach (string cp in contactProps)
                    sb.AppendLine("    $obj | Add-Member -NotePropertyName " + cp + " -NotePropertyValue ($(if ($c) { [string]$c." + cp + " } else { '' })) -Force");
                if (needManager)
                {
                    sb.AppendLine("    $mgr = ''; if ($c -and $c.Manager) { $mgr = Resolve-Recip $c.Manager }");
                    sb.AppendLine("    $obj | Add-Member -NotePropertyName Manager -NotePropertyValue $mgr -Force");
                }
                sb.AppendLine("    $obj");
                sb.AppendLine("}");
            }

            sb.AppendLine();
            sb.Append(ctx.ExportCsv("$rows"));
            sb.AppendLine("Write-Host 'Export complete.'");
            return sb.ToString();
        }
    }
}
