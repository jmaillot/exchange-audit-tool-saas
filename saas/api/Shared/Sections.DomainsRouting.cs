using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "Domains / Routing" category: Transport rules, Accepted domains, Remote domains, Connectors.
    // These are organization-level objects (no -ResultSize). Multi-valued props join with ','.
    internal static class SectionsDomainsRouting
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildTransportRulesSection());
            AuditRegistry.Register(BuildAcceptedDomainsSection());
            AuditRegistry.Register(BuildRemoteDomainsSection());
            AuditRegistry.Register(BuildConnectorsSection());
        }

        // Properties that are collections -> joined with ',' so a ';' CSV never clashes.
        private static readonly List<string> MultiValued = new List<string>(new string[]
        {
            "SenderDomains", "SenderIPAddresses", "RecipientDomains", "SmartHosts", "AddressSpaces",
            "Bindings", "RemoteIPRanges", "PermissionGroups", "SourceTransportServers",
            "FromAddressContainsWords", "SubjectContainsWords", "SentToScope"
        });

        private static string BuildSelectList(List<string> chosen)
        {
            var exprs = new List<string>();
            foreach (string v in chosen)
            {
                if (MultiValued.Contains(v))
                    exprs.Add("@{Name='" + v + "';Expression={($_." + v + " | ForEach-Object { [string]$_ }) -join ','}}");
                else if (v == "EmailAddresses")
                    exprs.Add("@{Name='EmailAddresses';Expression={($_.EmailAddresses | Where-Object {$_ -like 'smtp:*'}) -join ','}}");
                else exprs.Add(v);
            }
            if (exprs.Count == 0) exprs.Add("Name");
            return string.Join(", ", exprs.ToArray());
        }

        private static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            return PsScriptHelpers.Collect(sel, groupKeys);
        }

        // ============================================================ 1. TRANSPORT RULES
        private static AuditSection BuildTransportRulesSection()
        {
            var section = new AuditSection(
                "transport-rules",
                "Transport rules",
                "Transport rule export",
                "Audit mail flow rules (Get-TransportRule). Auto-detect keeps only populated named properties.",
                "rule",
                AuditScope.Both);
            section.Category = "Domains / Routing";
            section.DefaultFileName = "TransportRules.csv";

            // Auto mode: scan all rules and keep only the named properties that are actually
            // populated on at least one rule (most of the ~150 rule properties are empty).
            var auto = new AuditOptionGroup("auto", "Smart mode", GroupMode.MultiCheck); auto.Columns = 1;
            auto.Hint = "Auto-detect keeps only columns that have a value on at least one rule (ignores the manual groups below).";
            auto.Add(new AuditOption("autodetect", "Auto-detect populated properties only (recommended)", "autodetect", true));
            section.AddGroup(auto);

            var identity = new AuditOptionGroup("identity", "Identity & metadata", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("Name", true);
            identity.AddProp("State", true);
            identity.AddProp("Priority", true);
            identity.AddProp("Mode", true);
            identity.AddProp("Comments", false);
            identity.AddProp("CreatedBy", false);
            identity.AddProp("LastModifiedBy", false);
            identity.AddProp("WhenChanged", false);
            identity.AddProp("ActivationDate", false);
            identity.AddProp("ExpiryDate", false);
            section.AddGroup(identity);

            var conditions = new AuditOptionGroup("conditions", "Conditions (named, exploitable)", GroupMode.MultiCheck); conditions.Columns = 2;
            conditions.Hint = "The raw Conditions object shows .NET type names; these named properties hold usable values.";
            conditions.AddProp("From", false);
            conditions.AddProp("FromMemberOf", false);
            conditions.AddProp("FromScope", false);
            conditions.AddProp("SentTo", false);
            conditions.AddProp("SentToMemberOf", false);
            conditions.AddProp("SentToScope", false);
            conditions.AddProp("SenderDomainIs", false);
            conditions.AddProp("RecipientDomainIs", false);
            conditions.AddProp("SubjectContainsWords", false);
            conditions.AddProp("SubjectOrBodyContainsWords", false);
            conditions.AddProp("FromAddressContainsWords", false);
            conditions.AddProp("HeaderContainsMessageHeader", false);
            conditions.AddProp("HeaderContainsWords", false);
            conditions.AddProp("SubjectMatchesPatterns", false);
            conditions.AddProp("AttachmentExtensionMatchesWords", false);
            conditions.AddProp("AttachmentNameMatchesPatterns", false);
            conditions.AddProp("MessageSizeOver", false);
            conditions.AddProp("AttachmentSizeOver", false);
            conditions.AddProp("WithImportance", false);
            conditions.AddProp("MessageTypeMatches", false);
            conditions.AddProp("SCLOver", false);
            conditions.AddProp("HasClassification", false);
            conditions.AddProp("SenderIpRanges", false);
            conditions.AddProp("SenderManagementRelationship", false);
            section.AddGroup(conditions);

            var exceptions = new AuditOptionGroup("exceptions", "Exceptions (named, exploitable)", GroupMode.MultiCheck); exceptions.Columns = 2;
            exceptions.AddProp("ExceptIfFrom", false);
            exceptions.AddProp("ExceptIfFromMemberOf", false);
            exceptions.AddProp("ExceptIfSentTo", false);
            exceptions.AddProp("ExceptIfSentToMemberOf", false);
            exceptions.AddProp("ExceptIfSenderDomainIs", false);
            exceptions.AddProp("ExceptIfRecipientDomainIs", false);
            exceptions.AddProp("ExceptIfSubjectContainsWords", false);
            exceptions.AddProp("ExceptIfSubjectOrBodyContainsWords", false);
            exceptions.AddProp("ExceptIfHeaderContainsWords", false);
            exceptions.AddProp("ExceptIfFromAddressContainsWords", false);
            exceptions.AddProp("ExceptIfMessageTypeMatches", false);
            section.AddGroup(exceptions);

            var actions = new AuditOptionGroup("actions", "Actions (named, exploitable)", GroupMode.MultiCheck); actions.Columns = 2;
            actions.AddProp("SetAuditSeverity", false);
            actions.AddProp("ModerateMessageByUser", false);
            actions.AddProp("ModerateMessageByManager", false);
            actions.AddProp("StopRuleProcessing", false);
            actions.AddProp("PrependSubject", false);
            actions.AddProp("ApplyClassification", false);
            actions.AddProp("ApplyHtmlDisclaimerText", false);
            actions.AddProp("ApplyHtmlDisclaimerLocation", false);
            actions.AddProp("SetSCL", false);
            actions.AddProp("SetHeaderName", false);
            actions.AddProp("SetHeaderValue", false);
            actions.AddProp("RemoveHeader", false);
            actions.AddProp("AddToRecipients", false);
            actions.AddProp("CopyTo", false);
            actions.AddProp("BlindCopyTo", false);
            actions.AddProp("RedirectMessageTo", false);
            actions.AddProp("AddManagerAsRecipientType", false);
            actions.AddProp("DeleteMessage", false);
            actions.AddProp("Quarantine", false);
            actions.AddProp("RejectMessageReasonText", false);
            actions.AddProp("RejectMessageEnhancedStatusCode", false);
            actions.AddProp("RouteMessageOutboundConnector", false);
            actions.AddProp("GenerateIncidentReport", false);
            actions.AddProp("ApplyOME", false);
            section.AddGroup(actions);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                bool autoDetect = sel.IsSelected("auto", "autodetect");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying transport rules...'");
                sb.AppendLine("$items = @(Get-TransportRule -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} transport rule(s).\" -f $items.Count)");
                sb.AppendLine();

                // Shared helper: format any property value (bool as-is, collections joined with ',').
                sb.AppendLine("function Format-TR { param($v)");
                sb.AppendLine("    if ($v -is [bool]) { return [string]$v }");
                sb.AppendLine("    if ($null -eq $v) { return '' }");
                sb.AppendLine("    if (-not ($v -is [string]) -and $v.PSObject.Properties['Count']) { return (($v | ForEach-Object { [string]$_ }) -join ',') }");
                sb.AppendLine("    return [string]$v");
                sb.AppendLine("}");
                sb.AppendLine();

                if (autoDetect)
                {
                    // Candidate named properties (excludes the raw .NET Conditions/Exceptions/Actions objects).
                    var candidates = new List<string>();
                    candidates.AddRange(new string[] { "Name", "State", "Priority", "Mode", "Comments", "CreatedBy", "LastModifiedBy", "WhenChanged", "ActivationDate", "ExpiryDate" });
                    foreach (string k in new string[] { "conditions", "exceptions", "actions" })
                        foreach (AuditOption o in FindGroup(section, k).Options)
                            candidates.Add(o.Value);

                    var arr = new StringBuilder();
                    foreach (string c in candidates) { if (arr.Length > 0) arr.Append(","); arr.Append("'" + c + "'"); }
                    sb.AppendLine("$candidates = @(" + arr.ToString() + ")");
                    // Columns always kept even if empty (identity/order/description).
                    sb.AppendLine("$always = @('Name','State','Priority','Mode')");
                    sb.AppendLine("$populated = New-Object System.Collections.Generic.List[string]");
                    sb.AppendLine("foreach ($p in $candidates) {");
                    sb.AppendLine("    if ($always -contains $p) { $populated.Add($p); continue }");
                    sb.AppendLine("    $has = $false");
                    sb.AppendLine("    foreach ($r in $items) {");
                    sb.AppendLine("        $v = $r.$p");
                    sb.AppendLine("        if ($v -is [bool]) { if ($v) { $has=$true; break } }");
                    sb.AppendLine("        elseif ($v -is [string]) { if ($v.Trim().Length -gt 0) { $has=$true; break } }");
                    sb.AppendLine("        elseif ($null -ne $v -and $v.PSObject.Properties['Count']) { if ($v.Count -gt 0) { $has=$true; break } }");
                    sb.AppendLine("        elseif ($null -ne $v) { if (([string]$v).Trim().Length -gt 0) { $has=$true; break } }");
                    sb.AppendLine("    }");
                    sb.AppendLine("    if ($has) { $populated.Add($p) }");
                    sb.AppendLine("}");
                    sb.AppendLine("Write-Host ('Populated columns: ' + ($populated -join ','))");
                    sb.AppendLine("$rows = foreach ($r in $items) {");
                    sb.AppendLine("    $o = New-Object psobject");
                    sb.AppendLine("    foreach ($c in $populated) { $o | Add-Member -NotePropertyName $c -NotePropertyValue (Format-TR $r.$c) -Force }");
                    sb.AppendLine("    $o");
                    sb.AppendLine("}");
                }
                else
                {
                    var chosen = Collect(sel, "identity", "conditions", "exceptions", "actions");
                    if (chosen.Count == 0) chosen.Add("Name");
                    var arr = new StringBuilder();
                    foreach (string c in chosen) { if (arr.Length > 0) arr.Append(","); arr.Append("'" + c + "'"); }
                    sb.AppendLine("$cols = @(" + arr.ToString() + ")");
                    sb.AppendLine("$rows = foreach ($r in $items) {");
                    sb.AppendLine("    $o = New-Object psobject");
                    sb.AppendLine("    foreach ($c in $cols) { $o | Add-Member -NotePropertyName $c -NotePropertyValue (Format-TR $r.$c) -Force }");
                    sb.AppendLine("    $o");
                    sb.AppendLine("}");
                }

                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // Finds a group by key within a section (used to enumerate candidate property names).
        private static AuditOptionGroup FindGroup(AuditSection section, string key)
        {
            foreach (AuditOptionGroup g in section.Groups) if (g.Key == key) return g;
            return new AuditOptionGroup(key, key, GroupMode.MultiCheck);
        }

        // ============================================================ 2. ACCEPTED DOMAINS
        private static AuditSection BuildAcceptedDomainsSection()
        {
            var section = new AuditSection(
                "accepted-domains",
                "Accepted domains",
                "Accepted domain export",
                "Audit accepted domains (Get-AcceptedDomain): authoritative / internal relay / external relay.",
                "rule",
                AuditScope.Both);
            section.Category = "Domains / Routing";
            section.DefaultFileName = "AcceptedDomains.csv";

            var props = new AuditOptionGroup("props", "Properties", GroupMode.MultiCheck); props.Columns = 2;
            props.AddProp("DomainName", true);
            props.AddProp("DomainType", true);
            props.AddProp("Default", true);
            props.AddProp("MatchSubDomains", false);
            props.AddProp("AddressBookEnabled", false);
            props.AddProp("ExternallyManaged", false);
            props.Add(AuditOption.PropScoped("OutboundOnly", true, false));
            props.Add(AuditOption.PropScoped("SendingFromDomainDisabled", true, false));
            props.Add(AuditOption.PropScoped("SendingToDomainDisabled", true, false));
            props.AddProp("CatchAllRecipientID", false);
            section.AddGroup(props);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "props");
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying accepted domains...'");
                sb.AppendLine("$items = @(Get-AcceptedDomain -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} accepted domain(s).\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 3. REMOTE DOMAINS
        private static AuditSection BuildRemoteDomainsSection()
        {
            var section = new AuditSection(
                "remote-domains",
                "Remote domains",
                "Remote domain export",
                "Audit remote domains (Get-RemoteDomain): OOF, auto-reply/forward, formatting.",
                "rule",
                AuditScope.Both);
            section.Category = "Domains / Routing";
            section.DefaultFileName = "RemoteDomains.csv";

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("Name", true);
            identity.AddProp("DomainName", true);
            identity.AddProp("IsInternal", false);
            identity.AddProp("TargetDeliveryDomain", false);
            section.AddGroup(identity);

            var behaviour = new AuditOptionGroup("behaviour", "Reply / forward behaviour", GroupMode.MultiCheck); behaviour.Columns = 2;
            behaviour.AddProp("AllowedOOFType", true);
            behaviour.AddProp("AutoReplyEnabled", true);
            behaviour.AddProp("AutoForwardEnabled", true);
            behaviour.AddProp("DeliveryReportEnabled", false);
            behaviour.AddProp("NDREnabled", false);
            behaviour.AddProp("MeetingForwardNotificationEnabled", false);
            section.AddGroup(behaviour);

            var format = new AuditOptionGroup("format", "Formatting", GroupMode.MultiCheck); format.Columns = 2;
            format.AddProp("TNEFEnabled", false);
            format.AddProp("CharacterSet", false);
            format.AddProp("NonMimeCharacterSet", false);
            format.AddProp("ContentType", false);
            section.AddGroup(format);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "behaviour", "format");
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying remote domains...'");
                sb.AppendLine("$items = @(Get-RemoteDomain -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} remote domain(s).\" -f $items.Count)");
                sb.AppendLine("$rows = $items | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 4. CONNECTORS
        private static AuditSection BuildConnectorsSection()
        {
            var section = new AuditSection(
                "connectors",
                "Connectors",
                "Connector export",
                "Audit inbound/outbound connectors. EXO: Get-Inbound/OutboundConnector; on-prem: Get-Receive/SendConnector.",
                "rule",
                AuditScope.Both);
            section.Category = "Domains / Routing";
            section.DefaultFileName = "Connectors.csv";

            var direction = new AuditOptionGroup("direction", "Direction", GroupMode.SingleChoice); direction.Columns = 2;
            direction.Add(new AuditOption("inbound", "Inbound", "inbound", true));
            direction.Add(new AuditOption("outbound", "Outbound", "outbound", false));
            section.AddGroup(direction);

            // EXO inbound (Get-InboundConnector)
            var exoIn = new AuditOptionGroup("exoin", "EXO inbound (Get-InboundConnector)", GroupMode.MultiCheck); exoIn.Columns = 2;
            exoIn.AddProp("Name", true);
            exoIn.AddProp("Enabled", true);
            exoIn.AddProp("ConnectorType", true);
            exoIn.AddProp("ConnectorSource", false);
            exoIn.AddProp("SenderDomains", true);
            exoIn.AddProp("SenderIPAddresses", true);
            exoIn.AddProp("RequireTls", false);
            exoIn.AddProp("TlsSenderCertificateName", false);
            exoIn.AddProp("CloudServicesMailEnabled", false);
            exoIn.AddProp("RestrictDomainsToIPAddresses", false);
            exoIn.AddProp("RestrictDomainsToCertificate", false);
            section.AddGroup(exoIn);

            // EXO outbound (Get-OutboundConnector)
            var exoOut = new AuditOptionGroup("exoout", "EXO outbound (Get-OutboundConnector)", GroupMode.MultiCheck); exoOut.Columns = 2;
            exoOut.AddProp("Name", true);
            exoOut.AddProp("Enabled", true);
            exoOut.AddProp("ConnectorType", true);
            exoOut.AddProp("ConnectorSource", false);
            exoOut.AddProp("RecipientDomains", true);
            exoOut.AddProp("SmartHosts", true);
            exoOut.AddProp("TlsSettings", false);
            exoOut.AddProp("TlsDomain", false);
            exoOut.AddProp("UseMxRecord", false);
            exoOut.AddProp("IsTransportRuleScoped", false);
            exoOut.AddProp("RouteAllMessagesViaOnPremises", false);
            section.AddGroup(exoOut);

            // On-prem inbound (Get-ReceiveConnector)
            var opIn = new AuditOptionGroup("opin", "On-prem inbound (Get-ReceiveConnector)", GroupMode.MultiCheck); opIn.Columns = 2;
            opIn.AddProp("Name", true);
            opIn.AddProp("Enabled", true);
            opIn.AddProp("Server", true);
            opIn.AddProp("Bindings", true);
            opIn.AddProp("RemoteIPRanges", true);
            opIn.AddProp("AuthMechanism", false);
            opIn.AddProp("PermissionGroups", false);
            opIn.AddProp("TransportRole", false);
            opIn.AddProp("TlsCertificateName", false);
            opIn.AddProp("MaxMessageSize", false);
            section.AddGroup(opIn);

            // On-prem outbound (Get-SendConnector)
            var opOut = new AuditOptionGroup("opout", "On-prem outbound (Get-SendConnector)", GroupMode.MultiCheck); opOut.Columns = 2;
            opOut.AddProp("Name", true);
            opOut.AddProp("Enabled", true);
            opOut.AddProp("AddressSpaces", true);
            opOut.AddProp("SmartHosts", true);
            opOut.AddProp("DNSRoutingEnabled", false);
            opOut.AddProp("SourceTransportServers", true);
            opOut.AddProp("TlsAuthLevel", false);
            opOut.AddProp("TlsCertificateName", false);
            opOut.AddProp("MaxMessageSize", false);
            section.AddGroup(opOut);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                bool online = ConnectionSettings.IsOnline;
                bool inbound = sel.First("direction", "inbound") == "inbound";

                string cmdlet;
                string groupKey;
                if (online) { cmdlet = inbound ? "Get-InboundConnector" : "Get-OutboundConnector"; groupKey = inbound ? "exoin" : "exoout"; }
                else { cmdlet = inbound ? "Get-ReceiveConnector" : "Get-SendConnector"; groupKey = inbound ? "opin" : "opout"; }

                var chosen = Collect(sel, groupKey);
                if (chosen.Count == 0) chosen.Add("Name");
                string selectList = BuildSelectList(chosen);

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying connectors (" + (inbound ? "inbound" : "outbound") + ")...'");
                sb.AppendLine("$items = @(" + cmdlet + " -ErrorAction SilentlyContinue)");
                sb.AppendLine("Write-Host (\"Retrieved {0} connector(s).\" -f $items.Count)");
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
