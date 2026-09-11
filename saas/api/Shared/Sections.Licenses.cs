using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "Licenses" category (Microsoft Graph, NOT Exchange Online):
    //   licenses-overview -> GET /subscribedSkus (tenant SKUs: display names,
    //     units, capability status, per-service-plan products + provisioning).
    //   licenses-users    -> GET /subscribedSkus (skuId -> SkuPartNumber map) +
    //     GET /users (per-user SKU assignments + enabled/disabled plans, all
    //     resolved to friendly display names, unique per user).
    //   licenses-unlicensed -> GET /users $filter=assignedLicenses/$count eq 0
    //     (users with no assignment; optional last sign-in needs AuditLog.Read.All).
    //   licenses-groups     -> GET /subscribedSkus + GET /groups (group-assigned
    //     licenses; needs Group.Read.All, skipped unless group names are checked)
    //     + GET /users licenseAssignmentStates (direct vs group-based counts).
    // "LicenceName" (product display name) and "ServicePlanFriendlyName" come
    // from Microsoft's official Entra licensing CSV (GUID -> Product_Display_Name,
    // Service_Plan_Name -> friendly name). Lookup chain inside each audit, best
    // first: 1) live download, 2) baked-in worker copy (/worker/licensing-names.csv,
    // see worker/Dockerfile), 3) raw IDs. All emitted cells are pre-joined
    // scalars, so Select-Object takes plain property names and the ';' CSV
    // never clashes (multi-values are ','-joined in PowerShell).
    internal static class SectionsLicenses
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildLicensesOverviewSection());
            AuditRegistry.Register(BuildLicenseUsersSection());
            AuditRegistry.Register(BuildUnlicensedUsersSection());
            AuditRegistry.Register(BuildGroupsLicensesSection());
        }

        private static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            return PsScriptHelpers.Collect(sel, groupKeys);
        }

        private const string LicenseNamesUrl =
            "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv";

        // Emits the $LicenceNameOf (SKU GUID -> product display name),
        // $PlanFriendlyOf (service plan code -> friendly name) and
        // $PlanFriendlyByIdOf (service plan GUID -> friendly name) lookups.
        // Unknown SKUs/plans simply miss the table and callers fall back to raw IDs.
        private static void AppendLicenseNameLookup(StringBuilder sb)
        {
            sb.AppendLine("$LicenseRows = @()");
            sb.AppendLine("try {");
            sb.AppendLine("    $dl = Invoke-WebRequest -Uri '" + LicenseNamesUrl + "' -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop");
            sb.AppendLine("    $LicenseRows = @($dl.Content.TrimStart([char]0xFEFF) | ConvertFrom-Csv)");
            sb.AppendLine("    Write-Host (\"Loaded license name database (live, {0} rows).\" -f $LicenseRows.Count)");
            sb.AppendLine("} catch {");
            sb.AppendLine("    Write-Host 'License name database download failed, trying baked-in copy...'");
            sb.AppendLine("    if (Test-Path '/worker/licensing-names.csv') {");
            sb.AppendLine("        $LicenseRows = @((Get-Content '/worker/licensing-names.csv' -Raw -Encoding UTF8).TrimStart([char]0xFEFF) | ConvertFrom-Csv)");
            sb.AppendLine("        Write-Host (\"Loaded license name database (baked-in, {0} rows).\" -f $LicenseRows.Count)");
            sb.AppendLine("    } else {");
            sb.AppendLine("        Write-Host 'No license name database available, using raw IDs.'");
            sb.AppendLine("    }");
            sb.AppendLine("}");
            sb.AppendLine("$LicenceNameOf = @{}");
            sb.AppendLine("$PlanFriendlyCand = @{}");
            sb.AppendLine("$PlanFriendlyIdCand = @{}");
            sb.AppendLine("foreach ($r in $LicenseRows) {");
            sb.AppendLine("    $g = [string]$r.GUID");
            sb.AppendLine("    if ($g -and -not $LicenceNameOf.ContainsKey($g) -and [string]$r.Product_Display_Name) { $LicenceNameOf[$g] = [string]$r.Product_Display_Name }");
            sb.AppendLine("    $fn = [string]$r.Service_Plans_Included_Friendly_Names");
            sb.AppendLine("    $pn = [string]$r.Service_Plan_Name");
            sb.AppendLine("    if ($pn -and $fn) {");
            sb.AppendLine("        if (-not $PlanFriendlyCand.ContainsKey($pn)) { $PlanFriendlyCand[$pn] = New-Object System.Collections.Generic.List[string] }");
            sb.AppendLine("        if ($PlanFriendlyCand[$pn] -notcontains $fn) { $PlanFriendlyCand[$pn].Add($fn) }");
            sb.AppendLine("    }");
            sb.AppendLine("    $pi = [string]$r.Service_Plan_Id");
            sb.AppendLine("    if ($pi -and $fn) {");
            sb.AppendLine("        if (-not $PlanFriendlyIdCand.ContainsKey($pi)) { $PlanFriendlyIdCand[$pi] = New-Object System.Collections.Generic.List[string] }");
            sb.AppendLine("        if ($PlanFriendlyIdCand[$pi] -notcontains $fn) { $PlanFriendlyIdCand[$pi].Add($fn) }");
            sb.AppendLine("    }");
            sb.AppendLine("}");
            sb.AppendLine("$PlanFriendlyOf = @{}");
            sb.AppendLine("foreach ($k in $PlanFriendlyCand.Keys) {");
            sb.AppendLine("    $best = @($PlanFriendlyCand[$k] | Where-Object { $_ -cmatch '[a-z]' })[0]");
            sb.AppendLine("    if (-not $best) { $best = $PlanFriendlyCand[$k][0] }");
            sb.AppendLine("    $PlanFriendlyOf[$k] = $best");
            sb.AppendLine("}");
            sb.AppendLine("$PlanFriendlyByIdOf = @{}");
            sb.AppendLine("foreach ($k in $PlanFriendlyIdCand.Keys) {");
            sb.AppendLine("    $best = @($PlanFriendlyIdCand[$k] | Where-Object { $_ -cmatch '[a-z]' })[0]");
            sb.AppendLine("    if (-not $best) { $best = $PlanFriendlyIdCand[$k][0] }");
            sb.AppendLine("    $PlanFriendlyByIdOf[$k] = $best");
            sb.AppendLine("}");
        }

        // SKU GUID -> product display name, falling back to the SkuPartNumber
        // so the column always carries signal.
        private const string LicenceLookupSnip =
            "$lic = [string]$s.skuPartNumber; if ($LicenceNameOf.ContainsKey([string]$s.skuId)) { $lic = $LicenceNameOf[[string]$s.skuId] }";

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
            sku.AddProp("LicenceName", true);
            sku.AddProp("SkuPartNumber", true);
            sku.AddProp("TotalUnits", true);
            sku.AddProp("ConsumedUnits", true);
            sku.AddProp("AvailableUnits", true);
            sku.AddProp("CapabilityStatus", true);
            sku.AddProp("SkuId", true);
            section.AddGroup(sku);

            var plan = new AuditOptionGroup("plan", "Product (service plan) properties", GroupMode.MultiCheck); plan.Columns = 2;
            plan.AddProp("ServicePlanFriendlyName", true);
            plan.AddProp("ServicePlanName", true);
            plan.AddProp("ProvisioningStatus", true);
            plan.AddProp("ServicePlanId", false);
            plan.AddProp("AppliesTo", false);
            section.AddGroup(plan);

            // Per-view defaults (the web resets checkboxes when the View radio
            // changes): Per SKU checks everything except the plan-friendly name;
            // Per product checks the licence identity plus every plan column.
            section.ViewDefaults["skus"] = new List<string>
            {
                "sku::LicenceName", "sku::SkuPartNumber", "sku::TotalUnits",
                "sku::ConsumedUnits", "sku::AvailableUnits", "sku::CapabilityStatus",
                "sku::SkuId", "plan::ServicePlanName"
            };
            section.ViewDefaults["products"] = new List<string>
            {
                "sku::LicenceName", "sku::SkuPartNumber", "sku::SkuId",
                "plan::ServicePlanFriendlyName", "plan::ServicePlanName",
                "plan::ServicePlanId", "plan::ProvisioningStatus", "plan::AppliesTo"
            };

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
                AppendLicenseNameLookup(sb);
                if (which == "products")
                {
                    sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                    sb.AppendLine("foreach ($s in $skus) {");
                    sb.AppendLine("    $total = 0; try { $total = [int]$s.prepaidUnits.enabled } catch { }");
                    sb.AppendLine("    $consumed = 0; try { $consumed = [int]$s.consumedUnits } catch { }");
                    sb.AppendLine("    " + LicenceLookupSnip);
                    sb.AppendLine("    foreach ($p in @($s.servicePlans)) {");
                    sb.AppendLine("        $fr = [string]$p.servicePlanName; if ($PlanFriendlyOf.ContainsKey($fr)) { $fr = $PlanFriendlyOf[$fr] }");
                    sb.AppendLine("        $rows.Add([pscustomobject]@{");
                    sb.AppendLine("            LicenceName = $lic; SkuPartNumber = $s.skuPartNumber; SkuId = $s.skuId;");
                    sb.AppendLine("            TotalUnits = $total; ConsumedUnits = $consumed; AvailableUnits = ($total - $consumed);");
                    sb.AppendLine("            CapabilityStatus = $s.capabilityStatus;");
                    sb.AppendLine("            ServicePlanFriendlyName = $fr; ServicePlanName = $p.servicePlanName; ServicePlanId = $p.servicePlanId;");
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
                    sb.AppendLine("    " + LicenceLookupSnip);
                    sb.AppendLine("    $rows.Add([pscustomobject]@{");
                    sb.AppendLine("        LicenceName = $lic; SkuPartNumber = $s.skuPartNumber; SkuId = $s.skuId;");
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
            identity.AddProp("DisplayName", true);
            identity.AddProp("UserPrincipalName", true);
            identity.AddProp("AccountEnabled", true);
            identity.AddProp("UsageLocation", true);
            section.AddGroup(identity);

            var licenses = new AuditOptionGroup("licenses", "License assignment", GroupMode.MultiCheck); licenses.Columns = 2;
            licenses.AddProp("AssignedSkus", true);
            licenses.AddProp("DisabledPlans", true);
            licenses.AddProp("EnabledPlans", true);
            section.AddGroup(licenses);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "licenses");
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
                AppendLicenseNameLookup(sb);
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
                sb.AppendLine("    $dis = New-Object System.Collections.Generic.List[string]");
                sb.AppendLine("    $enb = New-Object System.Collections.Generic.List[string]");
                sb.AppendLine("    foreach ($a in @($u.assignedLicenses)) {");
                sb.AppendLine("        $sid = [string]$a.skuId");
                sb.AppendLine("        if ($skuMap.ContainsKey($sid)) { $names.Add($skuMap[$sid]) } else { $names.Add($sid) }");
                sb.AppendLine("        foreach ($d in @($a.disabledPlans)) {");
                sb.AppendLine("            $dfn = [string]$d; if ($PlanFriendlyByIdOf.ContainsKey($dfn)) { $dfn = $PlanFriendlyByIdOf[$dfn] }");
                sb.AppendLine("            if ($dis -notcontains $dfn) { $dis.Add($dfn) }");
                sb.AppendLine("        }");
                sb.AppendLine("    }");
                sb.AppendLine("    foreach ($p in @($u.assignedPlans)) {");
                sb.AppendLine("        if ([string]$p.capabilityStatus -eq 'Enabled') {");
                sb.AppendLine("            $pfn = [string]$p.servicePlanId; if ($PlanFriendlyByIdOf.ContainsKey($pfn)) { $pfn = $PlanFriendlyByIdOf[$pfn] }");
                sb.AppendLine("            if ($enb -notcontains $pfn) { $enb.Add($pfn) }");
                sb.AppendLine("        }");
                sb.AppendLine("    }");
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        DisplayName = $u.displayName; UserPrincipalName = $u.userPrincipalName;");
                sb.AppendLine("        AccountEnabled = $u.accountEnabled; UsageLocation = $u.usageLocation;");
                sb.AppendLine("        AssignedSkus = ($names -join ','); DisabledPlans = ($dis -join ','); EnabledPlans = ($enb -join ',')");
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

        // ============================================================ 3. UNLICENSED USERS
        private static AuditSection BuildUnlicensedUsersSection()
        {
            var section = new AuditSection(
                "licenses-unlicensed",
                "Unlicensed users",
                "Unlicensed users export",
                "Audit users with no license assignment via Graph (assignedLicenses/$count eq 0).",
                "user",
                AuditScope.Graph);
            section.Category = "Licensing";
            section.DefaultFileName = "UnlicensedUsers.csv";

            var identity = new AuditOptionGroup("identity", "User properties", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("DisplayName", true);
            identity.AddProp("UserPrincipalName", true);
            identity.AddProp("AccountEnabled", true);
            identity.AddProp("UserType", true);
            identity.AddProp("UsageLocation", true);
            identity.AddProp("CreatedDateTime", true);
            identity.AddProp("Mail", true);
            section.AddGroup(identity);

            var signin = new AuditOptionGroup("signin", "Sign-in activity", GroupMode.MultiCheck); signin.Columns = 1;
            signin.AddProp("LastSignInDateTime", true);
            section.AddGroup(signin);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "signin");
                if (chosen.Count == 0) chosen.Add("UserPrincipalName");
                string selectList = string.Join(", ", chosen.ToArray());
                // signInActivity is only requested from Graph when its column is
                // checked: without AuditLog.Read.All the whole query would 403.
                bool wantSignIn = chosen.Contains("LastSignInDateTime");
                string size = sel.First("size", "Unlimited");
                string topParam = size == "100" ? "&$top=100" : "&$top=999";
                string maxUsers = size == "100" ? "100" : (size == "1000" ? "1000" : "-1");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying unlicensed users (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine("$users = @()");
                sb.AppendLine("$uri = 'https://graph.microsoft.com/v1.0/users?$count=true&$filter=assignedLicenses/$count%20eq%200&$select=displayName,userPrincipalName,accountEnabled,userType,usageLocation,createdDateTime,mail" + (wantSignIn ? ",signInActivity" : "") + topParam + "'");
                sb.AppendLine("$maxUsers = " + maxUsers);
                sb.AppendLine("while ($uri -and ($maxUsers -lt 0 -or $users.Count -lt $maxUsers)) {");
                sb.AppendLine("    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop");
                sb.AppendLine("    $users += @($r.value)");
                sb.AppendLine("    $uri = $r.'@odata.nextLink'");
                sb.AppendLine("}");
                sb.AppendLine("if ($maxUsers -ge 0 -and $users.Count -gt $maxUsers) { $users = $users[0..($maxUsers - 1)] }");
                sb.AppendLine("Write-Host (\"Retrieved {0} unlicensed user(s).\" -f $users.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("foreach ($u in $users) {");
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        DisplayName = $u.displayName; UserPrincipalName = $u.userPrincipalName;");
                sb.AppendLine("        AccountEnabled = $u.accountEnabled; UserType = $u.userType;");
                sb.AppendLine("        UsageLocation = $u.usageLocation; CreatedDateTime = $u.createdDateTime; Mail = $u.mail;");
                sb.AppendLine("        LastSignInDateTime = $u.signInActivity.lastSignInDateTime");
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

        // ============================================================ 4. GROUPS & LICENSES
        private static AuditSection BuildGroupsLicensesSection()
        {
            var section = new AuditSection(
                "licenses-groups",
                "Groups & licenses",
                "Group vs direct license assignments export",
                "Audit per-SKU license assignments via Graph: group-based vs direct user assignments.",
                "key",
                AuditScope.Graph);
            section.Category = "Licensing";
            section.DefaultFileName = "LicenseGroups.csv";

            var sku = new AuditOptionGroup("sku", "SKU properties", GroupMode.MultiCheck); sku.Columns = 2;
            sku.AddProp("LicenceName", true);
            sku.AddProp("SkuPartNumber", true);
            sku.AddProp("SkuId", true);
            sku.AddProp("TotalUnits", true);
            sku.AddProp("ConsumedUnits", true);
            section.AddGroup(sku);

            var assign = new AuditOptionGroup("assign", "Assignment (group vs direct)", GroupMode.MultiCheck); assign.Columns = 2;
            assign.AddProp("AssignmentType", true);
            assign.AddProp("AssignedGroups", true);
            assign.AddProp("DirectUsers", true);
            assign.AddProp("GroupUsers", true);
            section.AddGroup(assign);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "sku", "assign");
                if (chosen.Count == 0) chosen.Add("SkuPartNumber");
                string selectList = string.Join(", ", chosen.ToArray());
                // Group display names are only fetched when the column is checked:
                // without Group.Read.All, uncheck AssignedGroups and the rest works.
                bool wantGroups = chosen.Contains("AssignedGroups");
                string size = sel.First("size", "Unlimited");
                string topParam = size == "100" ? "&$top=100" : "&$top=999";
                string maxUsers = size == "100" ? "100" : (size == "1000" ? "1000" : "-1");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying subscribed SKUs (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(FetchSkus);
                sb.AppendLine("Write-Host (\"Retrieved {0} subscribed SKU(s).\" -f $skus.Count)");
                AppendLicenseNameLookup(sb);
                sb.AppendLine("$licGroups = @{}");
                if (wantGroups)
                {
                    sb.AppendLine("try {");
                    sb.AppendLine("Write-Host 'Querying groups with licenses (Graph)...'");
                    sb.AppendLine("$guri = 'https://graph.microsoft.com/v1.0/groups?$count=true&$filter=assignedLicenses/$count%20ne%200&$select=id,displayName,assignedLicenses&$top=999'");
                    sb.AppendLine("while ($guri) {");
                    sb.AppendLine("    $gr = Invoke-RestMethod -Uri $guri -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("    foreach ($gg in @($gr.value)) {");
                    sb.AppendLine("        foreach ($al in @($gg.assignedLicenses)) {");
                    sb.AppendLine("            $gk = [string]$al.skuId");
                    sb.AppendLine("            if (-not $licGroups.ContainsKey($gk)) { $licGroups[$gk] = New-Object System.Collections.Generic.List[string] }");
                    sb.AppendLine("            if ($licGroups[$gk] -notcontains [string]$gg.displayName) { $licGroups[$gk].Add([string]$gg.displayName) }");
                    sb.AppendLine("        }");
                    sb.AppendLine("    }");
                    sb.AppendLine("    $guri = $gr.'@odata.nextLink'");
                    sb.AppendLine("}");
                    sb.AppendLine("} catch {");
                    sb.AppendLine("    Write-Host (\"WARNING: groups query failed (Group.Read.All consent?), continuing without group names: \" + $_.Exception.Message)");
                    sb.AppendLine("}");
                }
                sb.AppendLine("Write-Host 'Querying users license states (Graph)...'");
                sb.AppendLine("$users = @()");
                sb.AppendLine("$uri = 'https://graph.microsoft.com/v1.0/users?$select=id,licenseAssignmentStates" + topParam + "'");
                sb.AppendLine("$maxUsers = " + maxUsers);
                sb.AppendLine("while ($uri -and ($maxUsers -lt 0 -or $users.Count -lt $maxUsers)) {");
                sb.AppendLine("    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop");
                sb.AppendLine("    $users += @($r.value)");
                sb.AppendLine("    $uri = $r.'@odata.nextLink'");
                sb.AppendLine("}");
                sb.AppendLine("if ($maxUsers -ge 0 -and $users.Count -gt $maxUsers) { $users = $users[0..($maxUsers - 1)] }");
                sb.AppendLine("Write-Host (\"Retrieved {0} user(s).\" -f $users.Count)");
                sb.AppendLine("$directN = @{}; $groupN = @{}");
                sb.AppendLine("foreach ($u in $users) {");
                sb.AppendLine("    foreach ($st in @($u.licenseAssignmentStates)) {");
                sb.AppendLine("        $sk = [string]$st.skuId");
                sb.AppendLine("        if ([string]$st.assignedByGroup) {");
                sb.AppendLine("            if (-not $groupN.ContainsKey($sk)) { $groupN[$sk] = 0 }");
                sb.AppendLine("            $groupN[$sk]++");
                sb.AppendLine("        } else {");
                sb.AppendLine("            if (-not $directN.ContainsKey($sk)) { $directN[$sk] = 0 }");
                sb.AppendLine("            $directN[$sk]++");
                sb.AppendLine("        }");
                sb.AppendLine("    }");
                sb.AppendLine("}");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("foreach ($s in $skus) {");
                sb.AppendLine("    $total = 0; try { $total = [int]$s.prepaidUnits.enabled } catch { }");
                sb.AppendLine("    $consumed = 0; try { $consumed = [int]$s.consumedUnits } catch { }");
                sb.AppendLine("    " + LicenceLookupSnip);
                sb.AppendLine("    $k = [string]$s.skuId");
                sb.AppendLine("    $dc = 0; if ($directN.ContainsKey($k)) { $dc = $directN[$k] }");
                sb.AppendLine("    $gc = 0; if ($groupN.ContainsKey($k)) { $gc = $groupN[$k] }");
                sb.AppendLine("    $atype = 'Unassigned'");
                sb.AppendLine("    if ($dc -gt 0 -and $gc -gt 0) { $atype = 'Both' }");
                sb.AppendLine("    elseif ($gc -gt 0) { $atype = 'Group' }");
                sb.AppendLine("    elseif ($dc -gt 0) { $atype = 'Direct' }");
                sb.AppendLine("    $gnames = ''; if ($licGroups.ContainsKey($k)) { $gnames = ($licGroups[$k] -join ',') }");
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        LicenceName = $lic; SkuPartNumber = $s.skuPartNumber; SkuId = $s.skuId;");
                sb.AppendLine("        TotalUnits = $total; ConsumedUnits = $consumed;");
                sb.AppendLine("        AssignmentType = $atype; AssignedGroups = $gnames; DirectUsers = $dc; GroupUsers = $gc");
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
