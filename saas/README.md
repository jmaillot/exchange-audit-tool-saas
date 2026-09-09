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

1. Entra admin center → App registrations → New: name `Exchange Audit SaaS`, supported account types **Accounts in any organizational directory (multitenant)**, redirect URI `https://<your-web>/signin-oidc`.
2. API permissions → Add **Exchange** delegated `ManageAsUser` (run EXO PowerShell as the signed-in reader) + **Microsoft Graph** delegated `User.Read`. No application permissions, no cert needed for delegated flow.
3. Expose no scopes. Note the **Application (client) ID**.
4. Customers consent once per tenant (admin): `https://login.microsoftonline.com/organizations/adminconsent?client_id=<YOUR-APP-ID>`.
5. Auditors need an Exchange read role (e.g. **View-Only Organization Management**). Delegated calls inherit their RBAC — no `New-ServicePrincipal` step.
6. Token acquisition is outside this compose stack (your identity layer / `az account get-access-token --resource https://outlook.office365.com`). Paste the token + tenant org into the web Home page. Optional later: add a cert to the same app for app-only scheduled runs (`Exchange.ManageAsApp` + role assignment).

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

1. Open the web app → Home → enter **Tenant organization** (`contoso.onmicrosoft.com`) + **Access token** → Connect. Token stays in browser memory; Disconnect drops it. API/worker never persist it.
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
