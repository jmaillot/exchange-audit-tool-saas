# Exchange Audit SaaS (Exchange Online)

Exchange Online audit exports as a Dockerized SaaS: 25 audit sections with selectable properties, `;`-delimited UTF-8 CSV + formatted XLSX export, Azure Portal-style web UI, single multi-tenant Entra app (token in memory only, dropped on disconnect).

Section definitions live in `saas/api/Shared/`, so the checkboxes and generated PowerShell always match the API that serves them.

## Architecture

```
web (nginx, Portal UI) -> api (.NET 8, :8080) -> worker (pwsh 7.4 + ExchangeOnlineManagement, :8081)
                                     | shared volume audit-data (/data/*.csv, *.xlsx, *.log)
```

* `GET /api/sections` — section catalog (Online defaults, `slow` flags). Drives all checkboxes.
* `POST /api/jobs` — validates `selection` against the allow-list (no raw PS accepted), builds the script with the shared `BuildScript()`, forwards to worker with the Bearer token in memory only.
* `GET /api/jobs/{id}` — status + 200-row preview + log tail.
* `GET /api/jobs/{id}/download?format=csv|xlsx` — export files.
* Worker `POST /run` — `Connect-ExchangeOnline -AccessToken …` (or `Connect-IPPSSession` for protection), runs the script, `Export-Csv -Delimiter ';'`, disconnects, drops the token.

## 1. One-time multi-tenant app registration (you, the publisher)

1. Entra admin center (`entra.microsoft.com`) → **Identity** → **Applications** → **App registrations** → **New registration**: name `Exchange Audit SaaS`, supported account types **Accounts in any organizational directory (multitenant)**. Note the **Application (client) ID** → put it in `saas/.env` as `EAT_CLIENT_ID=`.
2. API permissions (exact clicks) — still on your app page, left menu **Manage → API permissions** → **Add a permission** → tab **APIs my organization uses** → search `Office 365 Exchange Online` → select it → **Delegated permissions** → tick **`Exchange.Manage`** → **Add permissions**. Then **Add a permission** → **Microsoft Graph** → **Delegated permissions** → `User.Read` is already there by default. Finish with **Grant admin consent for [your org]** (green checkmarks). No application permissions, no cert needed for delegated flow.
3. Authentication — left menu **Manage → Authentication** → **Add a platform** → **Single-page application**, redirect URI `https://<your-web>/`, **Save**; then tick **Allow public client flows** → **Save**. The web app signs users in directly (MSAL + PKCE, no secret).
4. Login library — nothing to do if the server has internet: the web app loads Microsoft's login library from their CDN automatically. Only for offline/air-gapped servers: download `msal-browser.min.js` (UMD build, MSAL.js releases) into `saas/web/` and rebuild `web`.
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

1. Open the web app → Home → enter your **work email (UPN)** → **Connect with Microsoft** and sign in. Tenant org is pre-filled from your email domain (editable). Token stays in browser memory (MSAL memory cache); Disconnect drops it. No token copy-paste — the **Advanced** section keeps manual paste as fallback.
2. Pick a section in the left nav (grouped by category, e.g. Mailboxes, Groups, Protection). Each section offers property groups with Online defaults, `Filter`, `Select all` and an amber `Slow options` warning.
3. Set output filename, keep `XLSX` checked if wanted, press **RUN AUDIT**. Poll `GET /api/jobs/{id}` every 3s; preview shows the first 200 rows.
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

## 4. Coverage (Online only, 25 sections)

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

## 5. Security notes

* Tokens in memory only (browser variable → `Authorization` header → worker env for the `pwsh` child). Never written to `/data`, logs, or images. `Disconnect-ExchangeOnline` after each job.
* `POST /api/jobs` allow-list validates every `group:value` against `AuditSection.Groups`; unknown groups/values → 422.
* Timeouts: connect ~5 min, audit 30 min (`EAT_AUDIT_TIMEOUT_SEC`). One session per job; queue client-side to respect EXO throttling. `Unlimited/1000/100` size group preserved.

## 6. Troubleshooting

| Symptom | Fix |
|---|---|
| `401 missing Bearer` | Paste a fresh EXO token (audience `https://outlook.office365.com`), check `X-Tenant-Id`. |
| `Unauthorized / access denied` | Auditor lacks Exchange read role; admin must consent once per tenant. |
| Worker `pwsh exit 1` | See `Activity log` / `/data/<job>.log` tail in job status; usually throttling or a `Slow` per-object lookup — rerun with fewer options. |
| No CSV produced | Check `.log`; in demo mode only 3 sample rows are written. |
| XLSX missing | API converts after worker success; check API logs. Limits 1048575 rows / 16384 cols. |
