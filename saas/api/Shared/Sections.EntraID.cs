using System;
using System.Collections.Generic;
using System.Text;

namespace ExchangeAuditTool
{
    // "Entra ID" category (Microsoft Graph, NOT Exchange Online):
    //   entra-users -> GET /users (one row per user: identity, account,
    //     organization, contact, address, on-premises sync) + optional
    //     per-user lookups, one query each, only when the column is checked:
    //     signInActivity.lastSignInDateTime (in the base query, needs
    //     AuditLog.Read.All or the whole query 403s),
    //     /users/{id}/authentication/methods (MFAEnabled = Yes when a
    //     non-password method is registered, needs
    //     UserAuthenticationMethod.Read.All),
    //     /users/{id}/transitiveMemberOf (AssignedRoles = directoryRole
    //     display names + GroupCount = group count, needs
    //     RoleManagement.Read.Directory),
    //     /users/{id}/manager (ManagerUPN, Get-MgUserManager equivalent).
    // All emitted cells are pre-joined scalars, so Select-Object takes plain
    // property names and the ';' CSV never clashes (multi-values are ','-joined).
    internal static class SectionsEntraID
    {
        public static void Register()
        {
            AuditRegistry.Register(BuildUsersSection());
        }

        private static List<string> Collect(AuditSelection sel, params string[] groupKeys)
        {
            return PsScriptHelpers.Collect(sel, groupKeys);
        }

        private const string GraphHeaders =
            "$headers = @{ Authorization = ('Bearer ' + $env:EAT_GRAPH_TOKEN); ConsistencyLevel = 'eventual' }";

        private static string SizePrelude(string size)
        {
            // NOTE: plain concatenation, NOT string.Format: the PowerShell body
            // contains literal { } blocks that string.Format would try to parse.
            return "$size = '" + size + "'\n$maxUsers = -1\nif ($size -eq '100') { $maxUsers = 100 } elseif ($size -eq '1000') { $maxUsers = 1000 }";
        }

        // ============================================================ 1. USERS
        private static AuditSection BuildUsersSection()
        {
            var section = new AuditSection(
                "entra-users",
                "Users",
                "Entra ID users export",
                "Audit Entra ID users via Graph: identity, organization and contact plus MFA state, roles, group count and manager (slow per-user lookups).",
                "user",
                AuditScope.Graph);
            section.Category = "Entra ID";
            section.Product = "Entra ID";
            section.DefaultFileName = "EntraUsers.csv";

            var identity = new AuditOptionGroup("identity", "Identity", GroupMode.MultiCheck); identity.Columns = 2;
            identity.AddProp("Id", true);
            identity.AddProp("DisplayName", true);
            identity.AddProp("GivenName", true);
            identity.AddProp("Surname", true);
            identity.AddProp("UserPrincipalName", true);
            identity.AddProp("Mail", true);
            identity.AddProp("MailNickname", true);
            section.AddGroup(identity);

            var account = new AuditOptionGroup("account", "Account", GroupMode.MultiCheck); account.Columns = 2;
            account.AddProp("AccountEnabled", true);
            account.AddProp("UserType", true);
            account.AddProp("CreatedDateTime", true);
            account.AddProp("EmployeeId", true);
            section.AddGroup(account);

            var org = new AuditOptionGroup("organization", "Organization", GroupMode.MultiCheck); org.Columns = 2;
            org.AddProp("Department", true);
            org.AddProp("CompanyName", true);
            org.AddProp("JobTitle", true);
            org.AddProp("OfficeLocation", true);
            section.AddGroup(org);

            var contact = new AuditOptionGroup("contact", "Contact & locale", GroupMode.MultiCheck); contact.Columns = 2;
            contact.AddProp("MobilePhone", true);
            contact.AddProp("BusinessPhones", true);
            contact.AddProp("PreferredLanguage", true);
            contact.AddProp("UsageLocation", true);
            section.AddGroup(contact);

            var address = new AuditOptionGroup("address", "Address", GroupMode.MultiCheck); address.Columns = 2;
            address.AddProp("City", true);
            address.AddProp("Country", true);
            address.AddProp("StreetAddress", true);
            address.AddProp("PostalCode", true);
            address.AddProp("State", true);
            section.AddGroup(address);

            var onprem = new AuditOptionGroup("onprem", "Hybrid sync", GroupMode.MultiCheck); onprem.Columns = 2;
            onprem.AddProp("OnPremisesSyncEnabled", true);
            onprem.AddProp("OnPremisesImmutableId", true);
            section.AddGroup(onprem);

            var signin = new AuditOptionGroup("signin", "Sign-in activity", GroupMode.MultiCheck); signin.Columns = 1;
            signin.Hint = "Needs the AuditLog.Read.All consent; without it the whole query is rejected - uncheck the column instead.";
            signin.AddProp("LastSignInDateTime", true);
            section.AddGroup(signin);

            var security = new AuditOptionGroup("security", "Authentication (slow per-user lookup)", GroupMode.MultiCheck); security.Columns = 1;
            security.Hint = "One authentication-methods query per user - slower on large tenants. Needs the UserAuthenticationMethod.Read.All consent.";
            security.Add(AuditOption.Prop("MFAEnabled", true).MarkSlow());
            section.AddGroup(security);

            var memberships = new AuditOptionGroup("memberships", "Roles & groups (slow per-user lookup)", GroupMode.MultiCheck); memberships.Columns = 2;
            memberships.Hint = "One transitive-membership query per user - slower on large tenants. Needs the RoleManagement.Read.Directory consent.";
            memberships.Add(AuditOption.Prop("AssignedRoles", true).MarkSlow());
            memberships.Add(AuditOption.Prop("GroupCount", true).MarkSlow());
            section.AddGroup(memberships);

            var manager = new AuditOptionGroup("manager", "Manager (slow per-user lookup)", GroupMode.MultiCheck); manager.Columns = 1;
            manager.Hint = "One manager query per user (Get-MgUserManager equivalent) - slower on large tenants. Empty when no manager is assigned.";
            manager.Add(AuditOption.Prop("ManagerUPN", true).MarkSlow());
            section.AddGroup(manager);

            PsScriptHelpers.AddSizeGroup(section);

            section.BuildScript = delegate (AuditSelection sel, ScriptContext ctx)
            {
                var chosen = Collect(sel, "identity", "account", "organization", "contact", "address", "onprem", "signin", "security", "memberships", "manager");
                if (chosen.Count == 0) chosen.Add("UserPrincipalName");
                string selectList = string.Join(", ", chosen.ToArray());
                string size = sel.First("size", "Unlimited");
                // signInActivity is only requested from Graph when its column
                // is checked: without AuditLog.Read.All the whole query 403s.
                bool wantSignIn = chosen.Contains("LastSignInDateTime");
                bool wantMfa = chosen.Contains("MFAEnabled");
                bool wantRoles = chosen.Contains("AssignedRoles");
                bool wantGroupCount = chosen.Contains("GroupCount");
                bool wantMembers = wantRoles || wantGroupCount;
                bool wantManager = chosen.Contains("ManagerUPN");

                var sb = new StringBuilder();
                sb.AppendLine("Write-Host 'Querying Entra ID users (Graph)...'");
                sb.AppendLine(GraphHeaders);
                sb.AppendLine(SizePrelude(size));
                sb.AppendLine("$select = 'id,displayName,givenName,surname,userPrincipalName,mail,mailNickname,accountEnabled,userType,createdDateTime,employeeId,department,companyName,jobTitle,officeLocation,mobilePhone,businessPhones,preferredLanguage,usageLocation,city,country,streetAddress,postalCode,state,onPremisesSyncEnabled,onPremisesImmutableId'");
                if (wantSignIn)
                    sb.AppendLine("$select += ',signInActivity'");
                sb.AppendLine("$users = @()");
                sb.AppendLine("$uri = 'https://graph.microsoft.com/v1.0/users?$select=' + $select + '&$top=999'");
                sb.AppendLine("while ($uri -and ($maxUsers -lt 0 -or $users.Count -lt $maxUsers)) {");
                sb.AppendLine("    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop");
                sb.AppendLine("    $users += @($r.value)");
                sb.AppendLine("    $uri = $r.'@odata.nextLink'");
                sb.AppendLine("}");
                sb.AppendLine("if ($maxUsers -ge 0 -and $users.Count -gt $maxUsers) { $users = $users[0..($maxUsers - 1)] }");
                sb.AppendLine("Write-Host (\"Retrieved {0} user(s).\" -f $users.Count)");
                sb.AppendLine("$rows = New-Object System.Collections.Generic.List[object]");
                sb.AppendLine("$i = 0");
                if (wantMfa) sb.AppendLine("$mfaFail = 0");
                if (wantMembers) sb.AppendLine("$memFail = 0");
                sb.AppendLine("foreach ($u in $users) {");
                sb.AppendLine("    $i++; if (($i % 20) -eq 0) { Write-Host (\"  ... {0}/{1}\" -f $i, $users.Count) }");
                sb.AppendLine("    $uid = [string]$u.id");
                sb.AppendLine("    $bp = ((@($u.businessPhones) | ForEach-Object { [string]$_ }) -join ',')");
                if (wantSignIn)
                    sb.AppendLine("    $ls = ''; try { if ($u.signInActivity) { $ls = [string]$u.signInActivity.lastSignInDateTime } } catch { }");
                if (wantMfa)
                {
                    sb.AppendLine("    $mf = ''");
                    sb.AppendLine("    try {");
                    sb.AppendLine("        $mm = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/users/\" + $uid + \"/authentication/methods\") -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("        $strong = @(@($mm.value) | Where-Object { [string]$_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' })");
                    sb.AppendLine("        if ($strong.Count -gt 0) { $mf = 'Yes' } else { $mf = 'No' }");
                    sb.AppendLine("    } catch { $mfaFail++ }");
                }
                if (wantMembers)
                {
                    sb.AppendLine("    $ar = ''; $gc2 = ''");
                    sb.AppendLine("    try {");
                    sb.AppendLine("        $tmo = @()");
                    sb.AppendLine("        $turi = \"https://graph.microsoft.com/v1.0/users/\" + $uid + \"/transitiveMemberOf?`$select=id,displayName\"");
                    sb.AppendLine("        while ($turi) { $tr = Invoke-RestMethod -Uri $turi -Headers $headers -Method Get -ErrorAction Stop; $tmo += @($tr.value); $turi = $tr.'@odata.nextLink' }");
                    sb.AppendLine("        $rn = @($tmo | Where-Object { [string]$_.'@odata.type' -eq '#microsoft.graph.directoryRole' } | ForEach-Object { [string]$_.displayName } | Where-Object { $_ -ne '' } | Sort-Object -Unique)");
                    sb.AppendLine("        $ar = ($rn -join ',')");
                    sb.AppendLine("        $gc2 = @($tmo | Where-Object { [string]$_.'@odata.type' -eq '#microsoft.graph.group' }).Count");
                    sb.AppendLine("    } catch { $memFail++ }");
                }
                if (wantManager)
                {
                    sb.AppendLine("    $mu = ''");
                    sb.AppendLine("    try {");
                    sb.AppendLine("        $mg = Invoke-RestMethod -Uri (\"https://graph.microsoft.com/v1.0/users/\" + $uid + \"/manager?`$select=userPrincipalName,mail\") -Headers $headers -Method Get -ErrorAction Stop");
                    sb.AppendLine("        $mu = [string]$mg.userPrincipalName");
                    sb.AppendLine("        if (-not $mu) { $mu = [string]$mg.mail }");
                    sb.AppendLine("    } catch { }");
                }
                sb.AppendLine("    $rows.Add([pscustomobject]@{");
                sb.AppendLine("        Id = $uid; DisplayName = $u.displayName; GivenName = $u.givenName; Surname = $u.surname;");
                sb.AppendLine("        UserPrincipalName = $u.userPrincipalName; MFAEnabled = " + (wantMfa ? "$mf" : "''") + "; Mail = $u.mail; MailNickname = $u.mailNickname;");
                sb.AppendLine("        AccountEnabled = $u.accountEnabled; UserType = $u.userType; CreatedDateTime = $u.createdDateTime;");
                sb.AppendLine("        LastSignInDateTime = " + (wantSignIn ? "$ls" : "''") + "; EmployeeId = $u.employeeId;");
                sb.AppendLine("        Department = $u.department; CompanyName = $u.companyName; JobTitle = $u.jobTitle; OfficeLocation = $u.officeLocation;");
                sb.AppendLine("        MobilePhone = $u.mobilePhone; BusinessPhones = $bp; PreferredLanguage = $u.preferredLanguage; UsageLocation = $u.usageLocation;");
                sb.AppendLine("        City = $u.city; Country = $u.country; StreetAddress = $u.streetAddress; PostalCode = $u.postalCode; State = $u.state;");
                sb.AppendLine("        OnPremisesSyncEnabled = $u.onPremisesSyncEnabled; OnPremisesImmutableId = $u.onPremisesImmutableId;");
                sb.AppendLine("        AssignedRoles = " + (wantRoles ? "$ar" : "''") + "; GroupCount = " + (wantGroupCount ? "$gc2" : "''") + "; ManagerUPN = " + (wantManager ? "$mu" : "''"));
                sb.AppendLine("    })");
                sb.AppendLine("}");
                if (wantMfa) sb.AppendLine("if ($mfaFail -gt 0) { Write-Host (\"WARNING: {0} MFA querie(s) failed (consent for UserAuthenticationMethod.Read.All?).\" -f $mfaFail) }");
                if (wantMembers) sb.AppendLine("if ($memFail -gt 0) { Write-Host (\"WARNING: {0} membership querie(s) failed (consent for RoleManagement.Read.Directory?).\" -f $memFail) }");
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
