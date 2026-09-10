/* Exchange Audit SaaS - fallback web settings.
 * Preferred way: set these in saas/.env (EAT_CLIENT_ID, EAT_MSAL_SRC,
 * EAT_EXO_SCOPES) — the API serves them at /api/config and they win.
 * Anything set here is only used when /api/config is unreachable. */
window.EAT_CONFIG = {
  clientId: "",
  msalSources: ["./msal-browser.min.js"],
  exoScopes: ["https://outlook.office365.com/.default"],
  graphScopes: ["User.Read.All", "Organization.Read.All"]
};
