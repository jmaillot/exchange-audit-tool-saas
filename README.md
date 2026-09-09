# Exchange Audit SaaS (Exchange Online)

Dockerized SaaS port of the Exchange Audit Tool: 25 Exchange Online audit sections with the same selectable properties and `;`-delimited CSV + XLSX export, Azure Portal-style web UI, single multi-tenant Entra app (token in memory only, dropped on disconnect).

Start here: [`saas/README.md`](saas/README.md) — multi-tenant app setup, Docker Compose setup, usage, API reference, troubleshooting.

```
saas/
  api/      → .NET 8 API (section catalog, job dispatch, XLSX export)
  api/Shared/ → audit DSL + section builders (Exchange Online only)
  worker/   → pwsh 7.4 + ExchangeOnlineManagement job runner
  web/      → nginx + Portal-style UI (no build step)
  docker-compose.yml
```
