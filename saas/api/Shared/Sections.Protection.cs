using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "Protection" category: Exchange Online Protection policies and rules plus
    // DLP. There is no on-premises equivalent, so this is ExchangeOnline-only
    // (the runtime scope guard blocks it elsewhere with guidance).
    internal static class SectionsProtection
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildProtectionSection());
        }

        // Properties that are collections -> joined with ',' so a ';' CSV never clashes.
        private static readonly List<string> MultiValued = new List<string>(new string[]
        {
            "AllowedSenders", "BlockedSenders", "AllowedSenderDomains", "BlockedSenderDomains",
            "FileTypes", "DoNotRewriteUrls"
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

        private static AuditSection BuildProtectionSection()
        {
            var section = new AuditSection(
                "protection-policies",
                "Protection policies",
                "Protection policy export",
                "Audit Exchange Online Protection, Safe Links/Attachments and DLP (policies + rules). One object per run.",
                "shield",
                AuditScope.ExchangeOnline);
            section.Category = "Protection";
            section.DefaultFileName = "ProtectionPolicies.csv";

            var obj = new AuditOptionGroup("object", "Object", GroupMode.SingleChoice); obj.Columns = 1;
            obj.Add(new AuditOption("antispamin", "Anti-spam inbound (Get-HostedContentFilterPolicy)", "antispamin", true));
            obj.Add(new AuditOption("antispamout", "Anti-spam outbound (Get-HostedOutboundSpamFilterPolicy)", "antispamout", false));
            obj.Add(new AuditOption("antimalware", "Anti-malware (Get-MalwareFilterPolicy)", "antimalware", false));
            obj.Add(new AuditOption("antiphish", "Anti-phishing (Get-AntiPhishPolicy)", "antiphish", false));
            obj.Add(new AuditOption("safelinks", "Safe Links (Get-SafeLinksPolicy)", "safelinks", false));
            obj.Add(new AuditOption("safeattachments", "Safe Attachments (Get-SafeAttachmentPolicy)", "safeattachments", false));
            obj.Add(new AuditOption("dlp", "DLP policies (Get-DlpPolicy)", "dlp", false));
            obj.Add(new AuditOption("ruleantispamin", "Anti-spam inbound rules (Get-HostedContentFilterRule)", "ruleantispamin", false));
            obj.Add(new AuditOption("ruleantispamout", "Anti-spam outbound rules (Get-HostedOutboundSpamFilterRule)", "ruleantispamout", false));
            obj.Add(new AuditOption("ruleantimalware", "Anti-malware rules (Get-MalwareFilterRule)", "ruleantimalware", false));
            obj.Add(new AuditOption("ruleantiphish", "Anti-phishing rules (Get-AntiPhishRule)", "ruleantiphish", false));
            obj.Add(new AuditOption("rulesafelinks", "Safe Links rules (Get-SafeLinksRule)", "rulesafelinks", false));
            obj.Add(new AuditOption("rulesafeattachments", "Safe Attachments rules (Get-SafeAttachmentRule)", "rulesafeattachments", false));
            section.AddGroup(obj);

            var antispamin = new AuditOptionGroup("antispamin", "Inbound spam properties", GroupMode.MultiCheck); antispamin.Columns = 2;
            antispamin.AddProp("Name", true);
            antispamin.AddProp("SpamAction", true);
            antispamin.AddProp("HighConfidenceSpamAction", true);
            antispamin.AddProp("BulkSpamAction", false);
            antispamin.AddProp("PhishSpamAction", false);
            antispamin.AddProp("HighConfidencePhishAction", false);
            antispamin.AddProp("BulkThreshold", false);
            antispamin.AddProp("QuarantineRetentionPeriod", false);
            antispamin.AddProp("AllowedSenders", false);
            antispamin.AddProp("BlockedSenders", false);
            antispamin.AddProp("AllowedSenderDomains", false);
            antispamin.AddProp("BlockedSenderDomains", false);
            section.AddGroup(antispamin);

            var antispamout = new AuditOptionGroup("antispamout", "Outbound spam properties", GroupMode.MultiCheck); antispamout.Columns = 2;
            antispamout.Hint = "AutoForwardingMode governs external auto-forwarding - check it before migration.";
            antispamout.AddProp("Name", false);
            antispamout.AddProp("AutoForwardingMode", false);
            antispamout.AddProp("NotifyOutboundSpam", false);
            antispamout.AddProp("NotifyOutboundSpamRecipients", false);
            antispamout.AddProp("BccSuspiciousOutboundMail", false);
            section.AddGroup(antispamout);

            var antimalware = new AuditOptionGroup("antimalware", "Malware properties", GroupMode.MultiCheck); antimalware.Columns = 2;
            antimalware.AddProp("Name", false);
            antimalware.AddProp("Action", false);
            antimalware.AddProp("EnableFileFilter", false);
            antimalware.AddProp("FileTypes", false);
            antimalware.AddProp("QuarantineRetentionPeriod", false);
            antimalware.AddProp("ZapEnabled", false);
            section.AddGroup(antimalware);

            var antiphish = new AuditOptionGroup("antiphish", "Phish properties", GroupMode.MultiCheck); antiphish.Columns = 2;
            antiphish.AddProp("Name", false);
            antiphish.AddProp("Enabled", false);
            antiphish.AddProp("AuthenticationFailAction", false);
            antiphish.AddProp("SpoofEnabled", false);
            antiphish.AddProp("EnableSpoofIntelligence", false);
            antiphish.AddProp("EnableUnauthenticatedSender", false);
            antiphish.AddProp("EnableViaTag", false);
            antiphish.AddProp("HonorDmarcPolicy", false);
            antiphish.AddProp("DmarcRejectAction", false);
            antiphish.AddProp("DmarcQuarantineAction", false);
            antiphish.AddProp("PhishThresholdLevel", false);
            section.AddGroup(antiphish);

            var safelinks = new AuditOptionGroup("safelinks", "Safe Links properties", GroupMode.MultiCheck); safelinks.Columns = 2;
            safelinks.AddProp("Name", false);
            safelinks.AddProp("Enabled", false);
            safelinks.AddProp("TrackClicks", false);
            safelinks.AddProp("AllowClickThrough", false);
            safelinks.AddProp("ScanUrls", false);
            safelinks.AddProp("EnableForInternalSenders", false);
            safelinks.AddProp("DoNotAllowClickThrough", false);
            safelinks.AddProp("DoNotRewriteUrls", false);
            section.AddGroup(safelinks);

            var safeattachments = new AuditOptionGroup("safeattachments", "Safe Attachments properties", GroupMode.MultiCheck); safeattachments.Columns = 2;
            safeattachments.AddProp("Name", false);
            safeattachments.AddProp("Enable", false);
            safeattachments.AddProp("Action", false);
            safeattachments.AddProp("Redirect", false);
            safeattachments.AddProp("RedirectAddress", false);
            safeattachments.AddProp("ActionOnError", false);
            safeattachments.AddProp("QuarantineTag", false);
            section.AddGroup(safeattachments);

            var dlp = new AuditOptionGroup("dlp", "DLP properties", GroupMode.MultiCheck); dlp.Columns = 2;
            dlp.AddProp("Name", false);
            dlp.AddProp("State", false);
            dlp.AddProp("Mode", false);
            dlp.AddProp("Priority", false);
            dlp.AddProp("Description", false);
            section.AddGroup(dlp);

            var ruleprops = new AuditOptionGroup("ruleprops", "Rule properties", GroupMode.MultiCheck); ruleprops.Columns = 2;
            ruleprops.Hint = "Rules bind policies to recipients; the linked policy column is added automatically.";
            ruleprops.AddProp("Name", false);
            ruleprops.AddProp("State", false);
            ruleprops.AddProp("Priority", false);
            ruleprops.AddProp("Comments", false);
            section.AddGroup(ruleprops);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string which = sel.First("object", "antispamin");
                string cmdlet;
                string propGroup;
                string label;
                string ruleLink = null;
                if (which == "antispamout") { cmdlet = "Get-HostedOutboundSpamFilterPolicy"; propGroup = "antispamout"; label = "outbound spam policies"; }
                else if (which == "antimalware") { cmdlet = "Get-MalwareFilterPolicy"; propGroup = "antimalware"; label = "malware policies"; }
                else if (which == "antiphish") { cmdlet = "Get-AntiPhishPolicy"; propGroup = "antiphish"; label = "phish policies"; }
                else if (which == "safelinks") { cmdlet = "Get-SafeLinksPolicy"; propGroup = "safelinks"; label = "Safe Links policies"; }
                else if (which == "safeattachments") { cmdlet = "Get-SafeAttachmentPolicy"; propGroup = "safeattachments"; label = "Safe Attachments policies"; }
                else if (which == "dlp") { cmdlet = "Get-DlpPolicy"; propGroup = "dlp"; label = "DLP policies"; }
                else if (which == "ruleantispamin") { cmdlet = "Get-HostedContentFilterRule"; propGroup = "ruleprops"; label = "inbound spam rules"; ruleLink = "HostedContentFilterPolicy"; }
                else if (which == "ruleantispamout") { cmdlet = "Get-HostedOutboundSpamFilterRule"; propGroup = "ruleprops"; label = "outbound spam rules"; ruleLink = "HostedOutboundSpamFilterPolicy"; }
                else if (which == "ruleantimalware") { cmdlet = "Get-MalwareFilterRule"; propGroup = "ruleprops"; label = "malware rules"; ruleLink = "MalwareFilterPolicy"; }
                else if (which == "ruleantiphish") { cmdlet = "Get-AntiPhishRule"; propGroup = "ruleprops"; label = "phish rules"; ruleLink = "AntiPhishPolicy"; }
                else if (which == "rulesafelinks") { cmdlet = "Get-SafeLinksRule"; propGroup = "ruleprops"; label = "Safe Links rules"; ruleLink = "SafeLinksPolicy"; }
                else if (which == "rulesafeattachments") { cmdlet = "Get-SafeAttachmentRule"; propGroup = "ruleprops"; label = "Safe Attachments rules"; ruleLink = "SafeAttachmentPolicy"; }
                else { cmdlet = "Get-HostedContentFilterPolicy"; propGroup = "antispamin"; label = "inbound spam policies"; }

                var chosen = Collect(sel, propGroup);
                if (ruleLink != null && !chosen.Contains(ruleLink)) chosen.Add(ruleLink);
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
