using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "Licenses" category (Microsoft Graph, NOT Exchange Online):
    //   licenses-overview -> GET /subscribedSkus (tenant SKUs: names, units,
    //     per-service-plan products + provisioning status).
    //   licenses-users    -> GET /subscribedSkus (skuId -> SkuPartNumber map) +
    //     GET /users (per-user SKU assignments + enabled service plans).
    // Both run in the worker with the Graph delegated token
    // ($env:EAT_GRAPH_TOKEN, Invoke-RestMethod, no EXO module, no extra
    // permissions beyond the token). All emitted cells are pre-joined
    // scalars, so Select-Object takes plain property names and the ';' CSV
    // never clashes (multi-values are ','-joined in PowerShell).
    internal static class SectionsLicenses
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildLicensesOverviewSection());
            AuditRegistry.Register(BuildLicenseUsersSection());
        }

        private static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            return PsScriptHelpers.Collect(sel, groupKeys);
        }

        private const string GraphHeaders =
            "$headers = @{ Authorization = ('Bearer ' + $env:EAT_GRAPH_TOKEN); ConsistencyLevel = 'eventual' }";

        private const string FetchSkus =
            "$skus = @()\n" +
            "$uri = 'https://graph.microsoft.com/v1.0/subscribedSkus'\n" +
            "while ($uri) {\n" +
            "    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop\n" +
            "    $skus += @($r.value)\n" +
            "    $uri = $r.'@odata.nextLink'\n" +
            "}";

        // ============================================================ 1. SKU OVERVIEW
        private static AuditSection BuildLicensesOverviewSection()
        {
            var section = new AuditSection(
                "licenses-overview",
                "License overview",
                "License SKU overview export",
                "Audit tenant license SKUs via Graph (subscribedSkus): units owned vs used plus products activated per SKU.",
                "key",
                AuditScope.Graph);
            section.Category = "Licensing";
            section.DefaultFileName = "LicensesOverview.csv";

            var obj = new AuditOptionGroup("object", "View", GroupMode.SingleChoice); obj.Columns = 1;
            obj.Add(new AuditOption("skus", "Per SKU (one row per license)", "skus", true));
            obj.Add(new AuditOption("products", "Per product (one row per SKU x service plan)", "products", false));
            section.AddGroup(obj);

            var sku = new AuditOptionGroup("sku", "SKU properties", GroupMode.MultiCheck); sku.Columns = 2;
            sku.AddProp("SkuPartNumber", true);
            sku.AddProp("TotalUnits", true);
            sku.AddProp("ConsumedUnits", true);
            sku.AddProp("AvailableUnits", true);
            sku.AddProp("SkuId", false);
            sku.AddProp("CapabilityStatus", false);
            section.AddGroup(sku);

            var plan = new AuditOptionGroup("plan", "Product (service plan) properties", GroupMode.MultiCheck); plan.Columns = 2;
            plan.AddProp("ServicePlanName", true);
            plan.AddProp("ProvisioningStatus", true);
            plan.AddProp("ServicePlanId", false);
            plan.AddProp("AppliesTo", false);
            section.AddGroup(plan);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                string which = sel.First("object", "skus");
                var chosen = Collect(sel, "sku", "plan");
                if (chosen.Count == 0) chosen.Add("SkuPartNumber");
                string selectList = string.Join(", ", chosen.ToArray());

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying subscribed SKUs (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(FetchSkus);
                sb.AppendLine("Write-Host (\"Retrieved {0} subscribed SKU(s).\" -f $skus.Count)");
                if (which == "products")
                {
                    sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                    sb.AppendLine("foreach ($s in $skus) {");
                    sb.AppendLine("    $total = 0; try { $total = [int]$s.prepaidUnits.enabled } catch { }");
                    sb.AppendLine("    $consumed = 0; try { $consumed = [int]$s.consumedUnits } catch { }");
                    sb.AppendLine("    foreach ($p in @($s.servicePlans)) {");
                    sb.AppendLine("        $rows.Add([pscustomobject]@{");
                    sb.AppendLine("            SkuPartNumber = $s.skuPartNumber; SkuId = $s.skuId;");
                    sb.AppendLine("            TotalUnits = $total; ConsumedUnits = $consumed; AvailableUnits = ($total - $consumed);");
                    sb.AppendLine("            CapabilityStatus = $s.capabilityStatus;");
                    sb.AppendLine("            ServicePlanName = $p.servicePlanName; ServicePlanId = $p.servicePlanId;");
                    sb.AppendLine("            ProvisioningStatus = $p.provisioningStatus; AppliesTo = $p.appliesTo");
                    sb.AppendLine("        })");
                    sb.AppendLine("    }");
                    sb.AppendLine("}");
                }
                else
                {
                    sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                    sb.AppendLine("foreach ($s in $skus) {");
                    sb.AppendLine("    $total = 0; try { $total = [int]$s.prepaidUnits.enabled } catch { }");
                    sb.AppendLine("    $consumed = 0; try { $consumed = [int]$s.consumedUnits } catch { }");
                    sb.AppendLine("    $rows.Add([pscustomobject]@{");
                    sb.AppendLine("        SkuPartNumber = $s.skuPartNumber; SkuId = $s.skuId;");
                    sb.AppendLine("        TotalUnits = $total; ConsumedUnits = $consumed; AvailableUnits = ($total - $consumed);");
                    sb.AppendLine("        CapabilityStatus = $s.capabilityStatus;");
                    sb.AppendLine("        ServicePlanName = ((@($s.servicePlans) | ForEach-Object { $_.servicePlanName }) -join ',');");
                    sb.AppendLine("        ServicePlanId = ((@($s.servicePlans) | ForEach-Object { $_.servicePlanId }) -join ',');");
                    sb.AppendLine("        ProvisioningStatus = ((@($s.servicePlans) | ForEach-Object { \"$($_.servicePlanName):$($_.provisioningStatus)\" }) -join ',');");
                    sb.AppendLine("        AppliesTo = ((@($s.servicePlans) | ForEach-Object { $_.appliesTo }) -join ',')");
                    sb.AppendLine("    })");
                    sb.AppendLine("}");
                }
                sb.AppendLine("$rows = $rows | Select-Object " + selectList);
                sb.AppendLine();
                sb.Append(ctx.ExportCsv("$rows"));
                sb.AppendLine("Write-Host 'Export complete.'");
                return sb.ToString();
            };

            return section;
        }

        // ============================================================ 2. USERS + LICENSES
        private static AuditSection BuildLicenseUsersSection()
        {
            var section = new AuditSection(
                "licenses-users",
                "Users & licenses",
                "Licensed users export",
                "Audit per-user license assignments via Graph: SKUs per user plus enabled products (service plans).",
                "user",
                AuditScope.Graph);
            section.Category = "Licensing";
            section.DefaultFileName = "LicenseUsers.csv";

            var identity = new AuditOptionGroup("identity", "User properties", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("UserPrincipalName", true);
            identity.AddProp("DisplayName", true);
            identity.AddProp("UsageLocation", true);
            identity.AddProp("AccountEnabled", false);
            section.AddGroup(identity);

            var licenses = new AuditOptionGroup("licenses", "License assignment", GroupMode.MultiCheck); licenses.Columns = 2;
            licenses.AddProp("AssignedSkus", true);
            licenses.AddProp("SkuIds", false);
            licenses.AddProp("DisabledPlans", false);
            section.AddGroup(licenses);

            var plans = new AuditOptionGroup("plans", "Activated products (service plans)", GroupMode.MultiCheck); plans.Columns = 2;
            plans.AddProp("EnabledServicePlans", true);
            plans.AddProp("DisabledServicePlans", false);
            section.AddGroup(plans);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "licenses", "plans");
                if (chosen.Count == 0) chosen.Add("UserPrincipalName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                string topParam = size == "100" ? "&$top=100" : "&$top=999";
                string maxUsers = size == "100" ? "100" : (size == "1000" ? "1000" : "-1");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying subscribed SKUs for license names (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(FetchSkus);
                sb.AppendLine("$skuMap = @{}");
                sb.AppendLine("foreach ($s in $skus) { $skuMap[[string]$s.skuId] = [string]$s.skuPartNumber }");
                sb.AppendLine("Write-Host 'Querying licensed users (Graph)...'");
                sb.AppendLine("$users = @()");
                sb.AppendLine("$uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,usageLocation,accountEnabled,assignedLicenses,assignedPlans" + topParam + "'");
                sb.AppendLine("$maxUsers = " + maxUsers);
                sb.AppendLine("while ($uri -and ($maxUsers -lt 0 -or $users.Count -lt $maxUsers)) {");
                sb.AppendLine("    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop");
                sb.AppendLine("    $users += @($r.value)");
                sb.AppendLine("    $uri = $r.'@odata.nextLink'");
                sb.AppendLine("}");
                sb.AppendLine("if ($maxUsers -ge 0 -and $users.Count -gt $maxUsers) { $users = $users[0..($maxUsers - 1)] }");
                sb.AppendLine("Write-Host (\"Retrieved {0} user(s).\" -f $users.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("foreach ($u in $users) {");
                sb.AppendLine("    $names = New-Object System.Collections.Generic.List[string]");
                sb.AppendLine("    $ids = New-Object System.Collections.Generic.List[string]");
                sb.AppendLine("    $dis = New-Object System.Collections.Generic.List[string]");
                sb.AppendLine("    foreach ($a in @($u.assignedLicenses)) {");
                sb.AppendLine("        $sid = [string]$a.skuId; $ids.Add($sid)");
                sb.AppendLine("        if ($skuMap.ContainsKey($sid)) { $names.Add($skuMap[$sid]) } else { $names.Add($sid) }");
                sb.AppendLine("        foreach ($d in @($a.disabledPlans)) { if (-not $dis.Contains([string]$d)) { $dis.Add([string]$d) } }");
                sb.AppendLine("    }");
                sb.AppendLine("    $en = ((@($u.assignedPlans) | Where-Object { $_.capabilityStatus -eq 'Enabled' } | ForEach-Object { $_.service }) -join ',')");
                sb.AppendLine("    $di = ((@($u.assignedPlans) | Where-Object { $_.capabilityStatus -eq 'Disabled' } | ForEach-Object { $_.service }) -join ',')");
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        UserPrincipalName = $u.userPrincipalName; DisplayName = $u.displayName;");
                sb.AppendLine("        UsageLocation = $u.usageLocation; AccountEnabled = $u.accountEnabled;");
                sb.AppendLine("        AssignedSkus = ($names -join ','); SkuIds = ($ids -join ','); DisabledPlans = ($dis -join ',');");
                sb.AppendLine("        EnabledServicePlans = $en; DisabledServicePlans = $di");
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
