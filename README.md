# Microsoft 365 Audit Tool (Exchange Online + Licensing + Teams + SharePoint + Entra ID)

Dockerized SaaS port of the Exchange Audit Tool: 42 audit sections (25 Exchange Online + 4 license sections + 9 Teams sections + 3 SharePoint sections + 1 Entra ID section, via Microsoft Graph) with the same selectable properties and `;`-delimited CSV + XLSX export, Azure Portal-style web UI, single multi-tenant Entra app (tokens in tab session only: F5-safe, dropped on disconnect or tab close).

Start here: [`saas/README.md`](saas/README.md) — multi-tenant app setup, Docker Compose setup, usage, API reference, troubleshooting.

```
saas/
  api/      → .NET 8 API (section catalog, job dispatch, XLSX export)
  api/Shared/ → audit DSL + section builders (Exchange Online + Graph)
  worker/   → pwsh 7.4 + ExchangeOnlineManagement job runner
  web/      → nginx + Portal-style UI (no build step)
  docker-compose.yml
```
