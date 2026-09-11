# Microsoft 365 Audit Tool (Exchange Online + Licensing + Teams + SharePoint + Entra ID)

Exchange Online audit exports as a Dockerized SaaS: 42 audit sections (25 Exchange Online with selectable properties + 4 license sections + 9 Teams sections + 3 SharePoint sections + 1 Entra ID section, via Microsoft Graph), `;`-delimited UTF-8 CSV + formatted XLSX export, Azure Portal-style web UI, single multi-tenant Entra app (tokens in tab session only: F5-safe, dropped on disconnect or tab close).

Section definitions live in `saas/api/Shared/`, so the checkboxes and generated PowerShell always match the API that serves them.

## Architecture

```
web (nginx, Portal UI) -> api (.NET 8, :8080) -> worker (pwsh 7.4 + ExchangeOnlineManagement, :8081)
                                     | shared volume audit-data (/data/*.csv, *.xlsx, *.log)
```

* `GET /api/sections` — section catalog (Online defaults, `slow` flags). Drives all checkboxes.
* `POST /api/jobs` — validates `selection` against the allow-list (no raw PS accepted), builds the script with the shared `BuildScript()`, forwards to worker with the Bearer token(s) live from the tab session only. Graph sections (`Licensing` + `Teams` + `SharePoint` + `Entra ID` categories, `scope: Graph`) additionally need the Graph token via `X-Graph-Token`.
* `GET /api/jobs/{id}` — status + 200-row preview + log tail.
* `GET /api/jobs/{id}/download?format=csv|xlsx` — export files.
* Worker `POST /run` — EXO sections: `Connect-ExchangeOnline -AccessToken …` (or `Connect-IPPSSession` for protection), runs the script, `Export-Csv -Delimiter ';'`, disconnects, drops the token. License sections: `Invoke-RestMethod` against Microsoft Graph with `$env:EAT_GRAPH_TOKEN` (subscribedSkus + users), no EXO connection. Outputs stay on the worker until the API pulls them (`GET /file/{job}?kind=csv|log`, then `DELETE`); storages are never shared, so nothing persists past the containers.

## 1. One-time multi-tenant app registration (you, the publisher)

1. Entra admin center (`entra.microsoft.com`) → **Identity** → **Applications** → **App registrations** → **New registration**: name `M365 Audit Tool`, supported account types **Accounts in any organizational directory (multitenant)**. Note the **Application (client) ID** → put it in `saas/.env` as `EAT_CLIENT_ID=`.
2. API permissions (exact clicks) — still on your app page, left menu **Manage → API permissions** → **Add a permission** → tab **APIs my organization uses** → search `Office 365 Exchange Online` → select it → **Delegated permissions** → tick **`Exchange.Manage`** → **Add permissions**. Then **Add a permission** → **Microsoft Graph** → **Delegated permissions** → `User.Read` is already there by default, add **`User.Read.All`** + **`Organization.Read.All`** (Licensing: per-user assignments + subscribed SKUs) + **`Group.Read.All`** + **`Team.ReadBasic.All`** + **`Channel.ReadBasic.All`** + **`TeamMember.Read.All`** + **`ChannelMember.Read.All`** + **`TeamsAppInstallation.ReadForTeam`** + **`TeamsTab.Read.All`** + **`Reports.Read.All`** + **`ChannelMessage.Read.All`** + **`Sites.Read.All`** + **`SharePointTenantSettings.Read.All`** (Teams + SharePoint categories) + **`AuditLog.Read.All`** (Last sign-in columns) + **`UserAuthenticationMethod.Read.All`** (MFA column) + **`RoleManagement.Read.Directory`** (roles/group counts) — also add each to `EAT_GRAPH_SCOPES`. Finish with **Grant admin consent for [your org]** (green checkmarks). No application permissions, no cert needed for delegated flow.
3. Authentication — left menu **Manage → Authentication** → **Add a platform** → **Single-page application**, redirect URI `https://<your-web>/`, **Save**; then tick **Allow public client flows** → **Save**. The web app signs users in directly (MSAL + PKCE, no secret).
4. Login library — Microsoft deprecated the MSAL CDN, so vendor the file once (any machine with Docker):
  ```bash
  docker run --rm -v /volume1/docker/exchange-audit-tool-saas/saas/web:/out node:20-alpine sh -c "cd /tmp && npm pack @azure/msal-browser@3 --silent && tar -xzf azure-msal-browser-*.tgz package/lib/msal-browser.min.js && cp package/lib/msal-browser.min.js /out/"
  cd saas && docker compose -f docker-compose.traefik.yml up -d --build web
  ```
  This takes the LTS 3.x UMD build (`msal-browser.min.js`, global `msal`) which matches the login code. Then verify: `docker exec exchange-audit-web ls -la /usr/share/nginx/html/msal-browser.min.js`.
5. Auditors sign in with an Exchange read role (e.g. **View-Only Organization Management**). Delegated calls inherit their RBAC — no `New-ServicePrincipal` step. First user in a tenant clicks **Register this tenant** in the web UI (admin, one click, pre-filled from the UPN domain); afterwards everyone just signs in.

## 2. Docker Compose setup

Prereqs: Docker + Compose plugin.

```bash
cp saas/.env.example saas/.env
# UI-only testing without Exchange:
echo 'EAT_DEMO_MODE=true' >> saas/.env
cd saas && docker compose up --build
```

* Web: http://localhost:3000 (Portal UI, `WEB_PORT` in `.env`)
* API: http://localhost:8080 (`/api/health`, `/api/sections`)
* Data volume `audit-data` holds `<jobId>.csv/.xlsx/.log/.script.ps1`.

Real Exchange run: set `EAT_DEMO_MODE=false`, restart worker. The worker image pre-installs `ExchangeOnlineManagement` (`saas/worker/Dockerfile`).

## 3. Usage

1. Open the web app → Home → enter your **work email (UPN)** → **Connect with Microsoft** and sign in. Tenant org is pre-filled from your email domain (editable). Tokens live in the tab session only (MSAL `sessionStorage`: F5-safe, never on disk); Disconnect or closing the tab drops them. No token copy-paste — the **Advanced** section keeps manual paste as fallback.
2. Pick a section in the left nav (grouped by category, e.g. Mailboxes, Groups, Protection, Licensing, Teams, SharePoint, Entra ID). Each section offers property groups with Online defaults, `Filter`, `Select all`, `Deselect slow` (only in sections with slow options, for a fast first pass) and an amber `Slow options` warning. The **Licensing**, **Teams**, **SharePoint** and **Entra ID** categories need the Graph consent from setup step 2 — without it, sign-in succeeds but runs fail with a Graph 403 (see troubleshooting).
3. Press **RUN AUDIT**. **Smart mode** (first block, on by default — native on transport-rules) drops columns empty on every row — applied to all sections (dropped columns are listed in the job log). Poll `GET /api/jobs/{id}` every 3s; preview shows the first 200 rows.
4. **Download CSV / XLSX**. CSV is `;`-delimited UTF-8 (multi-values `,`-joined); XLSX has bold header, filter, frozen top row.
5. **Activity log** view shows the executed (redacted: `-AccessToken ***`, never the token) commands and job output.

API example:

```bash
TOKEN=<exo-access-token>; ORG=contoso.onmicrosoft.com
curl -s localhost:8080/api/sections | head -c 300
curl -s -X POST localhost:8080/api/jobs \
 -H "Authorization: Bearer $TOKEN" -H "X-Tenant-Id: $ORG" -H 'Content-Type: application/json' \
 -d '{"sectionId":"mailbox-list","organization":"'"$ORG"'","includeXlsx":true,
      "selection":{"identity":["DisplayName","PrimarySmtpAddress"],"size":["Unlimited"]}}'
# -> {"jobId":"...","status":"queued"}
curl -s localhost:8080/api/jobs/<jobId> -H "X-Tenant-Id: $ORG" | head -c 500
curl -OJ localhost:8080/api/jobs/<jobId>/download?format=csv
```

## 4. Coverage (Online + Graph licenses + Graph Teams + Graph SharePoint + Graph Entra ID, 42 sections)

On-premises-only sections (`address-books`, `certificates`) are hidden in SaaS.

| IDs | Cmdlets (worker) |
|---|---|
| `mailbox-list`, `mobile-devices`, `shared-mailboxes`, `room-mailboxes`, `equipment-mailboxes` | `Get-EXOMailbox`, `Get-EXOMailboxPermission/Statistics`, `Get-EXORecipientPermission`, `Get-MobileDevice`, `Get-CalendarProcessing`, `Get-Place`, `Get-User`, `Get-MailboxRegionalConfiguration/FolderPermission` |
| `distribution-groups`, `security-groups`, `dynamic-groups`, `m365-groups` | `Get-DistributionGroup/Member`, `Get-DynamicDistributionGroup`, `Get-UnifiedGroup/Links`, `Get-Recipient`, `Get-EXORecipientPermission` |
| `mail-users`, `mail-contacts` | `Get-MailUser`, `Get-MailContact`, `Get-Contact` |
| `transport-rules`, `accepted-domains`, `remote-domains`, `connectors` | `Get-TransportRule/AcceptedDomain/RemoteDomain/InboundConnector/OutboundConnector` |
| `org-sharing`, `address-policies`, `journal-rules`, `retention-policies`, `owa-policy`, `role-policies` | `Get-SharingPolicy/OrganizationRelationship/FederationTrust/OrganizationConfig/EmailAddressPolicy/JournalRule/RetentionPolicy(Tag)/OwaMailboxPolicy/RoleAssignmentPolicy/ManagementRoleAssignment` |
| `protection-policies` | `Get-HostedContentFilter(Policy/Rule)`, `Get-HostedOutboundSpamFilter*`, `Get-MalwareFilter*`, `Get-AntiPhish*`, `Get-SafeLinks*`, `Get-SafeAttachment*`, `Get-DlpPolicy` via `Connect-IPPSSession` |
| `pf-mailboxes`, `pf-hierarchy`, `mail-pf` | `Get-Mailbox -PublicFolder`, `Get-PublicFolder(-Statistics/ClientPermission)`, `Get-MailPublicFolder` |
| `licenses-overview`, `licenses-users`, `licenses-unlicensed`, `licenses-groups` (category `Licensing`, Graph — not EXO) | `GET /subscribedSkus` (SKU display names via Microsoft licensing CSV with baked-in fallback, total/consumed units, capability status, service plans + provisioning status), `GET /users` (per-user SKU assignments, enabled/disabled service plans; unlicensed users via `assignedLicenses/$count eq 0`, optional last sign-in needs `AuditLog.Read.All`; group vs direct counts via `licenseAssignmentStates`), `GET /groups` (group-assigned licenses, needs `Group.Read.All`) |
| `teams-inventory`, `teams-channels`, `teams-channel-memberships`, `teams-members`, `teams-apps`, `teams-tabs`, `teams-sites`, `teams-activity`, `teams-compliance` (category `Teams`, Graph — not EXO) | `GET /groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')` (team identity, visibility, dynamic membership, created date), `GET /teams/{id}` (archived flag + guest/member/messaging/fun settings), `GET /teams/{id}/channels` + `/tabs` (channel/private/shared/tabs counts), `GET /teams/{id}/installedApps/$count` (apps count), `GET /groups/{id}/owners|members/$count` (owner/member/guest counts), `GET /groups/{id}/sites/root` (SharePoint URL), `GET /teams/{id}/channels` + `/filesFolder` + `/members` + `?$select=summary` (channel type, SharePoint folder, owner/member/guest counts; private/shared channel memberships with cached `/users/{id}` lookup for UPN/UserType), `GET /teams/{id}/installedApps?$expand=teamsApp,teamsAppDefinition` (app name, publisher, app ID, version, distribution + Microsoft-publisher filters), `GET /teams/{id}/channels/{id}/tabs?$expand=teamsApp` (tab name, app type, URL), `GET /groups/{id}/sites/root` (SharePoint URL, last modified) + `GET /groups/{id}/drive` (live storage quota) + usage report (site template + file count, ≤48h stale), `GET /teams/{id}/channels/{id}/messages?$top=1&$expand=replies` (latest message per team) |
| `sharepoint-sites`, `sharepoint-subsites`, `sharepoint-sharing` (category `SharePoint`, Graph — not EXO) | Group-connected sites enumerated through their M365 groups (live title/URL/dates/quota/owners/labels/Teams flag) + private/shared channel sites via `filesFolder` + usage report as best-effort enrichment (template, owner, status) matched by site id/URL; non-group sites need `sp-migration/`. Recursive `/sites/{id}/sites` walk for subsites (IDs/URL/title/dates). `GET /admin/sharepoint/settings` + `/organization` (tenant sharing policy, one row — needs `SharePointTenantSettings.Read.All` + SharePoint Admin/Global Reader role; full `Get-SPOTenant` surface in `sp-migration/`) |
| `entra-users` (category `Entra ID`, Graph — not EXO) | `GET /users` (identity, account, organization, contact, address, hybrid sync; last sign-in via `signInActivity` needs `AuditLog.Read.All`) + slow per-user lookups, only when checked: `/authentication/methods` (MFA flag, needs `UserAuthenticationMethod.Read.All`), `/transitiveMemberOf` (directory roles + group count, needs `RoleManagement.Read.Directory`), `/manager` (manager UPN, `Get-MgUserManager` equivalent) |
| `spo-admin`, `pnp-deep-dive` (folder `sp-migration/`, admin workstation — not SaaS) | `Get-SPOSite` (lock, sharing, hub, quotas, deleted sites) + PnP per-web deep dive (webs, lists, fields, content types, pages, webparts, navigation, groups, unique permissions, features, regional settings) |

## 5. Security notes

* Tokens in tab session only (MSAL `sessionStorage` → `Authorization` / `X-Graph-Token` headers → worker env for the `pwsh` child). Survives F5, dies with the tab or Disconnect. Never written to disk (`localStorage`), `/data`, logs, or images. `Disconnect-ExchangeOnline` after each EXO job; Graph jobs use no persistent connection.
* `POST /api/jobs` allow-list validates every `group:value` against `AuditSection.Groups`; unknown groups/values → 422.
* Timeouts: connect ~5 min, audit 30 min (`EAT_AUDIT_TIMEOUT_SEC`). One session per job; queue client-side to respect EXO throttling. `Unlimited/1000/100` size group preserved.

## 6. Troubleshooting

| Symptom | Fix |
|---|---|
| `401 missing Bearer` | Paste a fresh EXO token (audience `https://outlook.office365.com`), check `X-Tenant-Id`. |
| `401 Missing X-Graph-Token` / Graph `403` on a Licensing, Teams, SharePoint or Entra ID run | Graph consent missing: grant `User.Read.All` + `Organization.Read.All` + `Group.Read.All` + `Team.ReadBasic.All` + `Channel.ReadBasic.All` + `TeamMember.Read.All` + `ChannelMember.Read.All` + `TeamsAppInstallation.ReadForTeam` + `TeamsTab.Read.All` + `Reports.Read.All` + `ChannelMessage.Read.All` + `Sites.Read.All` + `SharePointTenantSettings.Read.All` + `AuditLog.Read.All` + `UserAuthenticationMethod.Read.All` + `RoleManagement.Read.Directory` (README §1 step 2), then **Disconnect + Connect** so both tokens are re-issued, and relaunch. `403` on Unlicensed users / Entra ID Users with Last sign-in checked = `AuditLog.Read.All` not consented (or uncheck the column). `403` with MFA / roles-group warnings in the Entra ID log = `UserAuthenticationMethod.Read.All` / `RoleManagement.Read.Directory` not consented (or uncheck those columns). `403` on Sharing Policy = auditor lacks the SharePoint Administrator or Global Reader role. |
| `Unauthorized / access denied` | Auditor lacks Exchange read role; admin must consent once per tenant. |
| Worker `pwsh exit 1` | See `Activity log` / `/data/<job>.log` tail in job status; usually throttling or a `Slow` per-object lookup — rerun with fewer options. |
| No CSV produced | Check `.log`; in demo mode only 3 sample rows are written. |
| XLSX missing | API converts after worker success; check API logs. Limits 1048575 rows / 16384 cols. |
| Refresh the baked licensing CSV | Audits always try the live Microsoft CSV first; the worker image only carries a fallback copy. Plain rebuilds reuse the cached layer (no re-check). To force an online freshness check (conditional download — skipped with 304 if unchanged): `CSV_REFRESH_TOKEN=$(date +%Y%m%d) docker compose up -d --build worker`. |
