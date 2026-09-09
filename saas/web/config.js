/* Exchange Audit SaaS - deployment settings. Edit clientId after registering
 * the single multi-tenant Entra app (see saas/README.md section 1). */
window.EAT_CONFIG = {
  // Application (client) ID of your multi-tenant Entra app.
  clientId: "PASTE-YOUR-APP-CLIENT-ID-HERE",
  // Microsoft login library (UMD build of @azure/msal-browser, global `msal`).
  // Download msal-browser.min.js from the MSAL.js releases and place it next
  // to index.html, or point here at your own hosted copy.
  msalSrc: "./msal-browser.min.js",
  // Delegated scopes requested at sign-in (EXO resource).
  exoScopes: ["https://outlook.office365.com/.default"]
};
