/* Exchange Audit SaaS - Azure Portal style blade. No build step. */
const API = "";
const state = { sections: [], current: null, activeProduct: "", checks: {}, smartMode: true, token: "", tokenExp: 0, graphToken: "", graphTokenExp: 0, restoring: false, org: "", upn: "", jobId: null, jobSection: null, jobStart: 0, poll: null, msal: null, msalAccount: null, connLabel: null };
// UI chrome dictionary. Anything served by /api/sections (section/category/
// group/option names, CSV columns) is NEVER translated on purpose.
const I18N = {
en: {
  searchSections: "Search audit sections", menuAria: "Menu", accountAria: "Account", navAria: "Navigation",
  navConnection: "Connection", navMonitor: "Monitor", navActivity: "Activity log",
  groupSys: "Connection & Monitor",
  crumbConnectionHome: "Connection &gt; M365 Audit",
  noticeHtml: `<strong>First time here?</strong> Enter your work email below, then click <a id="registerLink" href="#" target="_blank" rel="noopener">Register this tenant</a> (admin, once per tenant) before connecting.`,
  connectTitle: "Connect with Microsoft",
  connectDesc: "Enter your work email (UPN) and sign in at Microsoft with an Exchange reader account (e.g. View-Only Organization Management). Tokens live in this tab only and are dropped on disconnect or tab close.",
  labelUpn: "Work email (UPN)", labelOrg: "Tenant organization",
  connectBtn: "Connect with Microsoft", disconnectBtn: "Disconnect & drop token",
  coverageTitle: "Coverage ({n} sections)",
  covSection: "Section", covCategory: "Category", covProduct: "Product",
  productLicensing: "Licensing", productExchangeOnline: "Exchange Online",
  crumbConnection: "Connection",
  slowBadge: "Slow options selected - this run may take longer",
  filterOptions: "Filter options",
  selectAll: "Select all", deselectAll: "Deselect all", selectAllCount: "Select all ({on}/{total})", deselectSlow: "Deselect slow",
  dlCsv: "Download CSV", dlXlsx: "Download XLSX",
  runAudit: "RUN AUDIT", cancelBtn: "Cancel",
  resultsPreview: "Results preview", resultsTable: "Results", noResults: "No results yet.",
  crumbActivity: "Connection &gt; Activity log", activityTitle: "Activity log",
  connectedAs: "Connected: {label}", notConnected: "Not connected",
  statusConnected: "Connected (tab-session token).", statusDisconnected: "Disconnected, token dropped.",
  enterUpn: "Enter your work email (UPN).",
  serverNotConfigured: "Server not configured: missing clientId (see .env EAT_CLIENT_ID).",
  msalUnavailable: "Microsoft login unavailable — please contact your administrator.",
  openingSignin: "Opening Microsoft sign-in...",
  signedInLog: "Signed in as {upn} (tenant {org}). Tab-session token.",
  signinFailed: "Sign-in failed: {err}", signinErrorLog: "Sign-in error: {err}", msalLoadLog: "MSAL load failed: {src}",
  tokenRefreshedLog: "Token refreshed silently.", disconnectedLog: "Disconnected, token dropped.",
  connectFirstAlert: "Connect first (Connection page).",
  sessionExpired: "Session expired, please reconnect.", tokenRefreshFailedLog: "Token refresh failed: {err}",
  queued: "Queued...", postError: "Error: {err}", postFailedLog: "FAILED: {body}", jobStartedLog: "Job {id} started.",
  jobGone: "Job {id} no longer known by the API (restarted?).",
  runningNote: "Running... {elapsed} (job {id}){err}", finishedNote: "Finished in {elapsed} (job {id}){err}",
  previewLabel: "{cols} columns x {rows} rows (preview of first 200). ",
  jobDoneLog: "Job {id} {status} in {elapsed}{cols}.", jobDoneCols: " ({n} columns)",
  openFailedLog: "Failed to open section: {err}",
  cancelLog: "Polling stopped (worker job continues to timeout).",
  apiDownLog: "API unreachable: {err}", apiDownCoverage: "API unreachable.",
  slowTag: "slow",
  smartTitle: "Smart mode", smartHint: "Keeps only columns that have a value on at least one row.",
  smartLabel: "Auto-detect populated properties only (recommended)",
  sectionsLoadedLog: "Loaded {n} sections from API."
},
fr: {
  searchSections: "Rechercher des sections", menuAria: "Menu", accountAria: "Compte", navAria: "Navigation",
  navConnection: "Connexion", navMonitor: "Supervision", navActivity: "Journal d'activité",
  groupSys: "Connexion & Monitor",
  crumbConnectionHome: "Connexion &gt; M365 Audit",
  noticeHtml: `<strong>Première visite ?</strong> Saisissez votre e-mail professionnel ci-dessous, puis cliquez <a id="registerLink" href="#" target="_blank" rel="noopener">Enregistrer ce tenant</a> (admin, une seule fois par tenant) avant de vous connecter.`,
  connectTitle: "Se connecter avec Microsoft",
  connectDesc: "Saisissez votre e-mail professionnel (UPN) et connectez-vous chez Microsoft avec un compte lecteur Exchange (ex. View-Only Organization Management). Les jetons restent dans cet onglet uniquement et sont supprimés à la déconnexion ou à la fermeture de l'onglet.",
  labelUpn: "E-mail professionnel (UPN)", labelOrg: "Organisation du tenant",
  connectBtn: "Se connecter avec Microsoft", disconnectBtn: "Se déconnecter",
  coverageTitle: "Couverture ({n} sections)",
  covSection: "Section", covCategory: "Catégorie", covProduct: "Produit",
  productLicensing: "Licences", productExchangeOnline: "Exchange Online",
  crumbConnection: "Connexion",
  slowBadge: "Options lentes sélectionnées — l'exécution sera bien plus longue",
  filterOptions: "Filtrer les options",
  selectAll: "Tout sélectionner", deselectAll: "Tout désélectionner", selectAllCount: "Tout sélectionner ({on}/{total})", deselectSlow: "Décocher lentes",
  dlCsv: "Télécharger CSV", dlXlsx: "Télécharger XLSX",
  runAudit: "LANCER L'AUDIT", cancelBtn: "Annuler",
  resultsPreview: "Aperçu des résultats", resultsTable: "Résultats", noResults: "Aucun résultat pour l'instant.",
  crumbActivity: "Connexion &gt; Journal d'activité", activityTitle: "Journal d'activité",
  connectedAs: "Connecté : {label}", notConnected: "Non connecté",
  statusConnected: "Connecté (jeton de session).", statusDisconnected: "Déconnecté, jeton supprimé.",
  enterUpn: "Saisissez votre e-mail professionnel (UPN).",
  serverNotConfigured: "Serveur non configuré : clientId manquant (voir .env EAT_CLIENT_ID).",
  msalUnavailable: "Connexion Microsoft indisponible — contactez votre administrateur.",
  openingSignin: "Ouverture de la connexion Microsoft…",
  signedInLog: "Connecté en tant que {upn} (tenant {org}). Jeton de session.",
  signinFailed: "Échec de connexion : {err}", signinErrorLog: "Erreur de connexion : {err}", msalLoadLog: "Chargement MSAL impossible : {src}",
  tokenRefreshedLog: "Jeton actualisé silencieusement.", disconnectedLog: "Déconnecté, jeton supprimé.",
  connectFirstAlert: "Connectez-vous d'abord (page Connexion).",
  sessionExpired: "Session expirée, reconnectez-vous.", tokenRefreshFailedLog: "Échec d'actualisation du jeton : {err}",
  queued: "En file d'attente…", postError: "Erreur : {err}", postFailedLog: "ÉCHEC : {body}", jobStartedLog: "Job {id} démarré.",
  jobGone: "Job {id} inconnu de l'API (redémarrée ?).",
  runningNote: "En cours… {elapsed} (job {id}){err}", finishedNote: "Terminé en {elapsed} (job {id}){err}",
  previewLabel: "{cols} colonnes x {rows} lignes (aperçu des 200 premières). ",
  jobDoneLog: "Job {id} {status} en {elapsed}{cols}.", jobDoneCols: " ({n} colonnes)",
  openFailedLog: "Ouverture de la section impossible : {err}",
  cancelLog: "Suivi arrêté (le job continue côté serveur jusqu'au timeout).",
  apiDownLog: "API injoignable : {err}", apiDownCoverage: "API injoignable.",
  slowTag: "lent",
  smartTitle: "Mode intelligent", smartHint: "Ne garde que les colonnes ayant une valeur sur au moins une ligne.",
  smartLabel: "Détection auto des propriétés remplies uniquement (recommandé)",
  sectionsLoadedLog: "{n} sections chargées depuis l'API."
}};
let lang = "en";
try { lang = localStorage.getItem("eat.lang") || ((navigator.language || "en").toLowerCase().startsWith("fr") ? "fr" : "en"); } catch (e) {}
if (!I18N[lang]) lang = "en";
function t(key, vars) {
  let s = (I18N[lang] && I18N[lang][key]) ?? I18N.en[key] ?? key;
  if (vars) for (const k in vars) s = s.split("{" + k + "}").join(vars[k]);
  return s;
}
function maskMid(s) {
  // Privacy: keep first 2 + last char, hide the middle
  // ("hswtfrance59820" -> "hs•••0", "adm-pdw-jmaillot" -> "ad•••t").
  s = String(s || "");
  if (s.length <= 4) return "•••";
  return s.slice(0, 2) + "•••" + s.slice(-1);
}
function maskHost(host) {
  // Mask only the tenant label, keep the public suffix readable
  // ("hswtfrance59820.onmicrosoft.com" -> "hs•••0.onmicrosoft.com").
  const parts = String(host || "").split(".");
  parts[0] = maskMid(parts[0]);
  return parts.join(".");
}
function maskId(s) {
  // Mask a UPN or domain for on-screen display. Full values stay in the
  // hover tooltip and out of screenshots.
  const i = String(s || "").indexOf("@");
  if (i < 0) return maskHost(s);
  return maskMid(s.slice(0, i)) + "@" + maskHost(s.slice(i + 1));
}
function refreshConnText() {
  const label = state.connLabel ? t("connectedAs", { label: state.connLabel }) : t("notConnected");
  $("connText").textContent = state.connLabel && state.upn
    ? t("connectedAs", { label: state.upn })
    : t("notConnected");
  $("connDot").title = label;
  $("connDot").classList.toggle("on", !!state.connLabel);
}
function applyI18n() {
  document.documentElement.lang = lang;
  document.querySelectorAll("[data-i18n]").forEach(el => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll("[data-i18n-ph]").forEach(el => { el.placeholder = t(el.dataset.i18nPh); el.setAttribute("aria-label", t(el.dataset.i18nPh)); });
  document.querySelectorAll("[data-i18n-aria]").forEach(el => { el.setAttribute("aria-label", t(el.dataset.i18nAria)); });
  document.querySelectorAll("[data-i18n-html]").forEach(el => { el.innerHTML = t(el.dataset.i18nHtml); });
  const lb = $("langBtn");
  if (lb) { lb.textContent = lang.toUpperCase(); lb.title = lang === "fr" ? "Switch to English" : "Passer en français"; }
  refreshConnText();
  refreshResultInfoLang();
  updateRegisterLink();
}
function refreshResultInfoLang() {
  // The results-preview line is set imperatively (status notes, empty text),
  // so data-i18n can't own it. Retranslate only when it shows the "no results"
  // placeholder in any language; live status text is left for the next poll.
  const el = $("resultInfo");
  if (!el) return;
  const known = Object.values(I18N).map(d => d.noResults).filter(Boolean);
  if (known.includes(el.textContent)) el.textContent = t("noResults");
}
function setLang(l) {
  lang = I18N[l] ? l : "en";
  try { localStorage.setItem("eat.lang", lang); } catch (e) {}
  applyI18n();
  renderProductTabs();
  renderNav();
  renderCoverage();
  const vis = ["home", "section", "activity"].find(x => !$("view-" + x).classList.contains("hidden"));
  if (vis === "section" && state.current) { markNav(state.current.id, state.current.category); renderGroups($("filter").value); updateSlow(); }
  else showView(vis || "home");
}
// Defaults; /api/config (backed by saas/.env) overrides, web/config.js is the fallback.
const EAT_CFG = Object.assign(
  { clientId: "", msalSources: ["./msal-browser.min.js"], exoScopes: ["https://outlook.office365.com/.default"], graphScopes: ["User.Read.All", "Organization.Read.All", "Group.Read.All", "Team.ReadBasic.All", "Channel.ReadBasic.All", "TeamMember.Read.All", "ChannelMember.Read.All", "TeamsAppInstallation.ReadForTeam", "TeamsTab.Read.All", "Reports.Read.All", "ChannelMessage.Read.All", "Sites.Read.All", "SharePointTenantSettings.Read.All", "AuditLog.Read.All", "UserAuthenticationMethod.Read.All", "RoleManagement.Read.Directory"] },
  window.EAT_CONFIG || {});
fetch("api/config").then(r => r.json()).then(c => {
  if (c.clientId) EAT_CFG.clientId = c.clientId;
  if (c.msalSources && c.msalSources.length) EAT_CFG.msalSources = c.msalSources;
  if (c.exoScopes && c.exoScopes.length) EAT_CFG.exoScopes = c.exoScopes;
  if (c.graphScopes && c.graphScopes.length) EAT_CFG.graphScopes = c.graphScopes;
  updateRegisterLink();
  // Config may arrive after the first restore attempt (clientId was empty):
  // retry the silent session restore once it is known.
  if (!state.msalAccount) restoreSession();
}).catch(() => {});
const $ = id => document.getElementById(id);
const logEl = () => $("activityLog");

function log(msg) {
  const t = new Date().toISOString().slice(11, 19);
  logEl().textContent += `[${t}] ${msg}\n`;
  logEl().scrollTop = 1e9;
}
function authHeaders() {
  const h = { "Content-Type": "application/json", "Authorization": "Bearer " + state.token, "X-Tenant-Id": state.org };
  if (state.graphToken) h["X-Graph-Token"] = "Bearer " + state.graphToken;
  return h;
}

async function loadSections() {
  const r = await fetch(API + "/api/sections");
  if (!r.ok) throw new Error("API " + r.status);
  state.sections = await r.json();
  initActiveProduct();
  renderProductTabs(); renderNav(); renderCoverage(); showView("home");
  log(t("sectionsLoadedLog", { n: state.sections.length }));
}

function setCat(block, open) {
  block.querySelector(".nav-items").style.display = open ? "" : "none";
  block.querySelector(".nav-cat").setAttribute("aria-expanded", open ? "true" : "false");
}
function markNav(id, cat) {
  document.querySelectorAll("#navGroups .nav-item").forEach(b => b.classList.toggle("active", b.dataset.section === id));
  document.querySelectorAll("#navGroups .nav-block").forEach(bl => bl.classList.toggle("open", !!cat && bl.dataset.cat === cat));
}
function renderNav() {
  const host = $("navGroups"); host.innerHTML = "";
  // Left nav shows the active product tab only (plus the system entries).
  const visible = state.sections.filter(s => productOf(s) === state.activeProduct);
  const cats = {};
  visible.forEach(s => { (cats[s.category || "Other"] ||= []).push(s); });
  // System entries first, then one collapsible block per category of the
  // active product tab. No product header needed: the tab above says it.
  const groups = [{ cat: t("groupSys"), items: [
    { id: "__home", title: t("navConnection"), view: "home" },
    { id: "__activity", title: t("navActivity"), view: "activity" }
  ] }];
  Object.keys(cats).sort().forEach(cat => {
    const items = cats[cat].map(s => ({ id: s.id, title: s.navTitle }));
    if (items.length) groups.push({ cat, items });
  });
  groups.forEach(g => {
    const block = document.createElement("div");
    block.className = "nav-block"; block.dataset.cat = g.cat;
    const head = document.createElement("button");
    head.className = "nav-cat";
    head.appendChild(document.createTextNode(g.cat + " "));
    const n = document.createElement("span");
    n.className = "nav-count"; n.textContent = g.items.length;
    head.appendChild(n);
    const wrap = document.createElement("div");
    wrap.className = "nav-items";
    g.items.forEach(it => {
      const b = document.createElement("button");
      b.className = "nav-item"; b.textContent = it.title; b.dataset.section = it.id;
      if (it.view) {
        b.dataset.view = it.view;
        b.onclick = () => showView(it.view);
      } else {
        b.onclick = () => openSection(it.id);
      }
      wrap.appendChild(b);
    });
    head.onclick = () => setCat(block, wrap.style.display === "none");
    block.appendChild(head); block.appendChild(wrap); host.appendChild(block);
    setCat(block, true);
  });
}

function productOf(s) {
  // Product comes from the API (section.Product). Fallback keeps old payloads
  // working: Graph scope => Licensing, everything else => Exchange Online.
  // New products (Teams, SharePoint…) need no frontend change: they appear as
  // their own tab automatically.
  if (s && s.product) return s.product;
  if (s && (s.scope === "Graph" || s.category === "Licensing")) return "Licensing";
  return "Exchange Online";
}
function productLabel(p) {
  if (p === "Exchange Online") return t("productExchangeOnline");
  if (p === "Licensing") return t("productLicensing");
  return p || "";
}
function listProducts() {
  // Stable order: Exchange Online first, Licensing last, anything else
  // (Teams, future products…) alphabetical in between.
  const seen = [...new Set(state.sections.map(productOf))];
  const rank = p => p === "Exchange Online" ? 0 : p === "Licensing" ? 2 : 1;
  return seen.sort((a, b) => rank(a) - rank(b) || a.localeCompare(b));
}
function initActiveProduct() {
  let saved = "";
  try { saved = localStorage.getItem("eat.product") || ""; } catch (e) {}
  const products = listProducts();
  state.activeProduct = products.includes(saved) ? saved : (products[0] || "Exchange Online");
}
function setActiveProduct(p) {
  if (!p || p === state.activeProduct) return;
  state.activeProduct = p;
  try { localStorage.setItem("eat.product", p); } catch (e) {}
  renderProductTabs(); renderNav(); renderCoverage(); applyTopSearch();
  if (state.current && productOf(state.current) !== p) {
    // Stay on the open section (its options/results are untouched); only the
    // left nav refilters. Going home is one click away.
  }
}
function renderProductTabs() {
  const host = $("productTabs");
  if (!host) return;
  host.innerHTML = "";
  listProducts().forEach(p => {
    const b = document.createElement("button");
    b.className = "product-tab" + (p === state.activeProduct ? " active" : "");
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", p === state.activeProduct ? "true" : "false");
    b.textContent = productLabel(p);
    b.onclick = () => setActiveProduct(p);
    host.appendChild(b);
  });
  host.style.display = listProducts().length ? "" : "none";
}
function renderCoverage() {
  // Coverage lists EVERY product (Exchange Online, Licensing, future Teams…
  // need no change here), one collapsed <details> per product. Collapsed by
  // default; click the header to expand.
  $("coverageTitle").textContent = t("coverageTitle", { n: state.sections.length });
  const host = $("coverage");
  host.innerHTML = "";
  listProducts().forEach(p => {
    const list = state.sections
      .filter(s => productOf(s) === p)
      .sort((a, b) => (a.category || "").localeCompare(b.category || "") || a.navTitle.localeCompare(b.navTitle));
    const det = document.createElement("details");
    det.className = "cov-block";
    const sum = document.createElement("summary");
    sum.className = "cov-head";
    sum.textContent = productLabel(p) + " — " + t("coverageTitle", { n: list.length });
    det.appendChild(sum);
    const rows = list.map(s =>
      `<tr><td>${esc(s.navTitle)}</td><td>${esc(s.category || "Other")}</td></tr>`);
    const tbl = document.createElement("table");
    tbl.className = "cov";
    tbl.innerHTML =
      `<thead><tr><th>${esc(t("covSection"))}</th><th>${esc(t("covCategory"))}</th></tr></thead>` +
      `<tbody>${rows.join("")}</tbody>`;
    det.appendChild(tbl);
    host.appendChild(det);
  });
}

function showView(v) {
  ["home", "section", "activity"].forEach(x => $("view-" + x).classList.toggle("hidden", x !== v));
  if (v === "home") markNav("__home", t("groupSys"));
  else if (v === "activity") markNav("__activity", t("groupSys"));
}

function openSection(id) {
  try {
  const s = state.sections.find(x => x.id === id); if (!s) return;
  // If the section lives under another product tab (e.g. restoring a
  // Licensing job while the Exchange tab is active), switch tabs first so
  // the left nav stays consistent with the open section.
  if (productOf(s) !== state.activeProduct) {
    state.activeProduct = productOf(s);
    try { localStorage.setItem("eat.product", state.activeProduct); } catch (e) {}
    renderProductTabs(); renderNav(); renderCoverage();
  }
  // Reopening the section of a known job resumes its live view instead of
  // wiping it (the job keeps running server-side either way).
  const resume = state.jobId && state.jobSection === id;
  state.current = s;
  if (!resume) state.checks = {};
  showView("section");
  markNav(id, s.category);
  $("crumbSection").textContent = s.navTitle;
  $("secTitle").textContent = s.title;
  $("secSub").textContent = s.subtitle || "";
  if (s.tipHtml) { $("secTip").innerHTML = s.tipHtml; $("secTip").classList.remove("hidden"); }
  else { $("secTip").innerHTML = ""; $("secTip").classList.add("hidden"); }
  $("filter").value = "";
  renderGroups("");
  // Fresh open (not a job resume): sections with per-view defaults start on
  // the selected view's set instead of the static option defaults.
  if (!resume && applyViewDefaults()) renderGroups("");
  updateSlow();
  if (resume) { if (!state.poll) pollStart(); else fetchJobStatus(); }
  else { pollStop(); resetResults(); state.jobSection = null; state.jobStart = 0; }
  } catch (e) { log(t("openFailedLog", { err: e.message || e })); }
}

function renderGroups(filter) {
  const host = $("groups"); host.innerHTML = "";
  const f = (filter || "").toLowerCase();
  // Smart mode block (same look as transport-rules' native one). Sections that
  // already define their own (key "auto") keep it; the flag is read in runAudit.
  const nativeSmart = state.current.groups.some(g => g.key === "auto" || (g.title || "").toLowerCase().includes("smart"));
  if (!nativeSmart) {
    const label = t("smartLabel");
    const matchTests = ["smart mode", "mode intelligent", I18N.en.smartLabel.toLowerCase(), I18N.fr.smartLabel.toLowerCase()];
    if (!f || matchTests.some(x => x.includes(f))) {
      const card = document.createElement("div"); card.className = "card grp";
      const h = document.createElement("h3"); h.textContent = t("smartTitle"); card.appendChild(h);
      const hint = document.createElement("div"); hint.className = "hint"; hint.textContent = t("smartHint"); card.appendChild(hint);
      const opts = document.createElement("div"); opts.className = "opts";
      const lbl = document.createElement("label"); lbl.className = "opt";
      const inp = document.createElement("input");
      inp.type = "checkbox"; inp.checked = state.smartMode;
      inp.onchange = () => { state.smartMode = inp.checked; };
      lbl.appendChild(inp);
      lbl.appendChild(document.createTextNode(label + " "));
      opts.appendChild(lbl); card.appendChild(opts); host.appendChild(card);
    }
  }
  state.current.groups.forEach(g => {
    const options = g.options || [];
    const visible = options.filter(o => o.label.toLowerCase().includes(f));
    if (f && visible.length === 0) return;
    const card = document.createElement("div"); card.className = "card grp";
    card.innerHTML = `<h3>${esc(g.title)}</h3>${g.hint ? `<div class="hint">${esc(g.hint)}</div>` : ""}`;
    const opts = document.createElement("div"); opts.className = "opts";
    options.forEach(o => {
      const key = g.key + "::" + o.value;
      if (!(key in state.checks)) state.checks[key] = !!o.defaultChecked;
      const lbl = document.createElement("label"); lbl.className = "opt" + (f && !o.label.toLowerCase().includes(f) ? " dim" : "");
      const inp = document.createElement("input");
      inp.type = g.mode === "SingleChoice" ? "radio" : "checkbox";
      inp.name = "g_" + g.key; inp.checked = state.checks[key];
      inp.onchange = () => {
        if (g.mode === "SingleChoice") {
          options.forEach(x => state.checks[g.key + "::" + x.value] = (x.value === o.value));
          applyViewDefaults();
          renderGroups($("filter").value);
        } else state.checks[key] = inp.checked;
        updateSlow(); updateSelectAll();
      };
      lbl.appendChild(inp);
      lbl.appendChild(document.createTextNode(o.label + " "));
      if (o.slow || g.slow) { const em = document.createElement("span"); em.className = "slow"; em.textContent = t("slowTag"); lbl.appendChild(em); }
      opts.appendChild(lbl);
    });
    card.appendChild(opts); host.appendChild(card);
  });
  updateSelectAll();
}

function viewDefaultSet() {
  // Returns the "group::value" list for the currently selected SingleChoice
  // view (licenses-overview Per SKU vs Per product), or null when the section
  // defines no per-view defaults.
  const vd = state.current && state.current.viewDefaults;
  if (!vd) return null;
  for (const g of state.current.groups) {
    if (g.mode !== "SingleChoice") continue;
    const sel = (g.options || []).find(x => state.checks[g.key + "::" + x.value]);
    if (sel && vd[sel.value]) return vd[sel.value];
  }
  return null;
}
function applyViewDefaults() {
  const set = viewDefaultSet();
  if (!set) return false;
  const want = new Set(set);
  state.current.groups.forEach(g => {
    if (g.mode === "SingleChoice") return;
    (g.options || []).forEach(o => { state.checks[g.key + "::" + o.value] = want.has(g.key + "::" + o.value); });
  });
  return true;
}
function sectionHasSlow() {
  return state.current && state.current.groups.some(g => g.slow || (g.options || []).some(o => o.slow));
}
function updateSlow() {
  const slow = state.current.groups.some(g => (g.options || []).some(o => (o.slow || g.slow) && state.checks[g.key + "::" + o.value]));
  $("slowBadge").classList.toggle("hidden", !slow);
  // One-click fast pass: visible only in sections that define slow options.
  $("deselectSlow").textContent = t("deselectSlow");
  $("deselectSlow").classList.toggle("hidden", !sectionHasSlow());
}
function updateSelectAll() {
  const keys = Object.keys(state.checks).filter(k => {
    const [gk, v] = k.split("::");
    const g = state.current.groups.find(x => x.key === gk);
    const o = g && ((g.options || []).find(x => x.value === v));
    return o && o.label.toLowerCase().includes(($("filter").value || "").toLowerCase()) && g.mode !== "SingleChoice";
  });
  const on = keys.filter(k => state.checks[k]).length;
  $("selectAll").textContent = keys.length && on === keys.length ? t("deselectAll") : t("selectAllCount", { on, total: keys.length });
}
function selection() {
  const sel = {};
  state.current.groups.forEach(g => {
    sel[g.key] = (g.options || []).filter(o => state.checks[g.key + "::" + o.value]).map(o => o.value);
  });
  return sel;
}
function resetResults() {
  $("grid").innerHTML = ""; $("resultInfo").textContent = t("noResults");
  $("jobLog").textContent = ""; $("jobLog").classList.add("hidden");
  $("dlCsv").disabled = $("dlXlsx").disabled = true;
}

async function runAudit() {
  if (!state.token || !state.org) { alert(t("connectFirstAlert")); showView("home"); return; }
  try { await ensureToken(); } catch (e) { $("resultInfo").textContent = t("sessionExpired"); log(t("tokenRefreshFailedLog", { err: e.message || e })); return; }
  if (state.current && state.current.scope === "Graph" && !state.graphToken) {
    try { await ensureGraphToken(false); }
    catch (e) { $("resultInfo").textContent = t("sessionExpired"); log(t("tokenRefreshFailedLog", { err: e.message || e })); return; }
    if (!state.graphToken) { alert(t("connectFirstAlert")); showView("home"); return; }
  }
  const sel = selection();
  const nativeAuto = state.current.groups.some(g => g.key === "auto");
  const smart = nativeAuto ? (sel.auto || []).includes("autodetect") : state.smartMode;
  const body = { sectionId: state.current.id, selection: sel, organization: state.org, includeXlsx: true, smartMode: smart };
  log(`RUN ${body.sectionId} selection=${JSON.stringify(body.selection)}`);
  $("runBtn").disabled = true; $("cancelBtn").disabled = false;
  $("resultInfo").textContent = t("queued");
  const r = await fetch(API + "/api/jobs", { method: "POST", headers: authHeaders(), body: JSON.stringify(body) });
  const j = await r.json();
  if (!r.ok) { $("resultInfo").textContent = t("postError", { err: j.error || r.status }); log(t("postFailedLog", { body: JSON.stringify(j) })); $("runBtn").disabled = false; $("cancelBtn").disabled = true; return; }
  state.jobId = j.jobId;
  state.jobSection = state.current.id;
  state.jobStart = Date.now();
  persistJob();
  log(t("jobStartedLog", { id: j.jobId }));
  pollStart();
}
function persistJob() {
  // Job tracking only (no token): survives a page reload, dies with the tab.
  try { sessionStorage.setItem("eat.lastJob", JSON.stringify({ jobId: state.jobId, section: state.jobSection, start: state.jobStart, org: state.org })); } catch (e) {}
}
async function restoreJob() {
  let saved = null;
  try { saved = JSON.parse(sessionStorage.getItem("eat.lastJob") || "null"); } catch (e) {}
  if (!saved || !saved.jobId || !saved.section) return;
  state.jobId = saved.jobId; state.jobSection = saved.section; state.jobStart = saved.start || Date.now();
  state.org = saved.org || ""; if (state.org && !$("org").value) $("org").value = state.org;
  try {
    const r = await fetch(API + "/api/jobs/" + state.jobId, { headers: { "X-Tenant-Id": state.org } });
    if (!r.ok) throw new Error("gone");
    const j = await r.json();
    if (j.status === "running" || j.status === "queued") openSection(state.jobSection);
    // Finished jobs: stay on Connection; reopening their section restores the view.
  } catch (e) {
    state.jobId = null; state.jobSection = null;
    try { sessionStorage.removeItem("eat.lastJob"); } catch (e2) {}
  }
}
function fmtElapsed(ms) {
  const s = Math.max(0, Math.floor(ms / 1000));
  const p = n => String(n).padStart(2, "0");
  return p(Math.floor(s / 3600)) + ":" + p(Math.floor(s % 3600 / 60)) + ":" + p(s % 60);
}
async function fetchJobStatus() {
  const r = await fetch(API + "/api/jobs/" + state.jobId, { headers: { "X-Tenant-Id": state.org } });
  if (!r.ok) {
    pollStop(); $("runBtn").disabled = false; $("cancelBtn").disabled = true;
    $("resultInfo").textContent = t("jobGone", { id: state.jobId });
    state.jobId = null; state.jobSection = null;
    try { sessionStorage.removeItem("eat.lastJob"); } catch (e) {}
    return;
  }
  const j = await r.json();
  const elapsed = state.jobStart ? fmtElapsed(Date.now() - state.jobStart) : "--:--:--";
  const active = j.status === "running" || j.status === "queued";
  const errSuffix = j.error ? " - " + j.error : "";
  const note = active ? t("runningNote", { elapsed, id: state.jobId, err: errSuffix })
                      : t("finishedNote", { elapsed, id: state.jobId, err: errSuffix });
  if (j.header) renderGrid(j.header, j.preview || [], note);
  else $("resultInfo").textContent = note;
  if (j.log) { $("jobLog").textContent = j.log.slice(-8000); $("jobLog").classList.remove("hidden"); }
  if (j.status === "succeeded" || j.status === "failed") {
    pollStop(); $("runBtn").disabled = false; $("cancelBtn").disabled = true;
    $("dlCsv").disabled = !j.hasCsv; $("dlXlsx").disabled = !j.hasXlsx;
    log(t("jobDoneLog", { id: state.jobId, status: j.status, elapsed, cols: j.header ? t("jobDoneCols", { n: j.header.length }) : "" }));
  }
}
function pollStart() {
  pollStop();
  fetchJobStatus();
  state.poll = setInterval(fetchJobStatus, 3000);
}
function pollStop() { if (state.poll) clearInterval(state.poll); state.poll = null; }
function renderGrid(header, rows, note) {
  const tbl = $("grid"); tbl.innerHTML = "";
  const trh = document.createElement("tr");
  header.forEach(h => { const th = document.createElement("th"); th.textContent = h; th.title = h; trh.appendChild(th); });
  tbl.appendChild(trh);
  rows.forEach(r => {
    const tr = document.createElement("tr");
    r.forEach(c => { const td = document.createElement("td"); td.textContent = c ?? ""; td.title = c ?? ""; tr.appendChild(td); });
    tbl.appendChild(tr);
  });
  $("resultInfo").textContent = t("previewLabel", { cols: header.length, rows: rows.length }) + (note || "");
}
function esc(s) { return String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

$("deselectSlow").onclick = () => {
  if (!state.current) return;
  state.current.groups.forEach(g => {
    if (g.mode === "SingleChoice") return;
    (g.options || []).forEach(o => { if (o.slow || g.slow) state.checks[g.key + "::" + o.value] = false; });
  });
  renderGroups($("filter").value); updateSlow();
};
$("selectAll").onclick = () => {
  const f = ($("filter").value || "").toLowerCase();
  const keys = [];
  state.current.groups.forEach(g => {
    if (g.mode === "SingleChoice") return;
    (g.options || []).forEach(o => { if (o.label.toLowerCase().includes(f)) keys.push(g.key + "::" + o.value); });
  });
  const target = keys.some(k => !state.checks[k]);
  keys.forEach(k => state.checks[k] = target);
  renderGroups($("filter").value); updateSlow();
};
$("filter").oninput = e => renderGroups(e.target.value);
$("runBtn").onclick = runAudit;
$("cancelBtn").onclick = async () => {
  // "Cancel" only stops watching: the server job keeps running. Take one
  // final refresh so the screen shows the truth instead of a frozen
  // "queued" label with disabled downloads; reopening the section resumes
  // live polling for a still-running job.
  pollStop(); $("runBtn").disabled = false; $("cancelBtn").disabled = true;
  log(t("cancelLog"));
  try { if (state.jobId) await fetchJobStatus(); } catch (e) { log(t("tokenRefreshFailedLog", { err: e.message || e })); }
};
$("crumbHome").onclick = e => { e.preventDefault(); showView("home"); };
$("hamburger").onclick = () => $("sidenav").classList.toggle("hidden");
$("langBtn").onclick = () => setLang(lang === "fr" ? "en" : "fr");
function setConnected(label) {
  state.connLabel = label;
  $("homeStatus").textContent = t("statusConnected");
  refreshConnText();
}
function setDisconnected(msg) {
  state.token = ""; state.tokenExp = 0; state.graphToken = ""; state.graphTokenExp = 0; state.msalAccount = null; state.jobId = null; state.connLabel = null; pollStop();
  $("homeStatus").textContent = msg || t("statusDisconnected");
  refreshConnText();
}
function upnDomain(upn) {
  const i = (upn || "").lastIndexOf("@");
  return i > 0 ? upn.slice(i + 1).trim().toLowerCase() : "";
}
function updateRegisterLink() {
  const d = upnDomain($("upn").value) || $("org").value.trim();
  const a = $("registerLink");
  if (d && EAT_CFG.clientId && EAT_CFG.clientId.indexOf("PASTE") !== 0) {
    a.href = "https://login.microsoftonline.com/" + encodeURIComponent(d) +
      "/adminconsent?client_id=" + encodeURIComponent(EAT_CFG.clientId);
    a.style.display = "";
  } else a.style.display = "none";
}
function loadScript(src) {
  return new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = src; s.onload = () => res(true);
    s.onerror = () => rej(new Error("load " + src));
    document.head.appendChild(s);
  });
}
$("upn").oninput = () => {
  const d = upnDomain($("upn").value);
  if (d && !$("org").value) $("org").value = d;
  updateRegisterLink();
};
$("org").oninput = updateRegisterLink;
async function loadMsal() {
  if (!window.msal) {
    for (const src of (EAT_CFG.msalSources || ["./msal-browser.min.js"])) {
      try { await loadScript(src); if (window.msal) break; }
      catch (e) { log(t("msalLoadLog", { src })); }
    }
  }
  return !!window.msal;
}
async function buildMsal(domain) {
  // Tab-lifetime session: tokens survive F5 (sessionStorage), die with the
  // tab, never touch disk (no localStorage), never reach our servers.
  const app = new window.msal.PublicClientApplication({
    auth: { clientId: EAT_CFG.clientId, authority: "https://login.microsoftonline.com/" + domain },
    cache: { cacheLocation: "sessionStorage" }
  });
  if (app.initialize) await app.initialize(); // MSAL v3+: required before any API call
  return app;
}
function persistSession(domain) {
  // Non-sensitive routing hints only (org/domain/UPN). Tokens stay inside
  // MSAL's own sessionStorage entries and are never handled here.
  try {
    sessionStorage.setItem("eat.org", state.org);
    sessionStorage.setItem("eat.domain", domain);
    sessionStorage.setItem("eat.upn", state.upn);
  } catch (e) {}
}
function clearPersistedSession() {
  // Local sign-out: drop MSAL's tab cache (access + refresh tokens) and our
  // session keys. Nothing ever left the tab, so no server call is needed.
  try {
    const drop = [];
    for (let i = 0; i < sessionStorage.length; i++) {
      const k = sessionStorage.key(i);
      if (k && (k.indexOf("msal.") === 0 || k.indexOf("eat.") === 0)) drop.push(k);
    }
    drop.forEach(k => sessionStorage.removeItem(k));
  } catch (e) {}
}
async function restoreSession() {
  // F5 survival: rebuild the session silently from the tab cache. Any failure
  // just means "stay disconnected" — never a popup, never a thrown error.
  // Guarded: startup + config-arrival can trigger two overlapping attempts.
  if (state.restoring || state.msalAccount) return;
  state.restoring = true;
  try {
    const domain = sessionStorage.getItem("eat.domain") || "";
    const upn = sessionStorage.getItem("eat.upn") || "";
    if (!domain || !EAT_CFG.clientId) return;
    if (!(await loadMsal())) return;
    const app = await buildMsal(domain);
    const accounts = app.getAllAccounts() || [];
    const account = accounts.find(a => a.username === upn) || accounts[0];
    if (!account) return;
    app.setActiveAccount(account);
    const tok = await app.acquireTokenSilent({
      scopes: EAT_CFG.exoScopes || ["https://outlook.office365.com/.default"],
      account
    });
    state.msal = app; state.msalAccount = account;
    state.token = tok.accessToken;
    state.tokenExp = (tok.expiresOn ? tok.expiresOn.getTime() : Date.now() + 50 * 60 * 1000);
    state.upn = account.username || upn;
    state.org = sessionStorage.getItem("eat.org") || domain;
    if ($("upn") && !$("upn").value) $("upn").value = state.upn;
    if ($("org")) $("org").value = state.org;
    try { await ensureGraphToken(true); } catch (e) { log(t("tokenRefreshFailedLog", { err: "Graph: " + (e.message || e) })); }
    setConnected(state.org + " (" + state.upn + ")");
    log(t("signedInLog", { upn: maskId(state.upn), org: maskHost(state.org) }));
  } catch (e) { log(t("signinErrorLog", { err: e.message || e })); }
  finally { state.restoring = false; }
}
$("connectBtn").onclick = async () => {
  const upn = $("upn").value.trim();
  const domain = upnDomain(upn);
  if (!upn || !domain) { $("homeStatus").textContent = t("enterUpn"); return; }
  if (!EAT_CFG.clientId || EAT_CFG.clientId.indexOf("PASTE") === 0) {
    $("homeStatus").textContent = t("serverNotConfigured");
    return;
  }
  if (!(await loadMsal())) {
    $("homeStatus").textContent = t("msalUnavailable");
    return;
  }
  try {
    $("homeStatus").textContent = t("openingSignin");
    const app = await buildMsal(domain);
    const scopes = EAT_CFG.exoScopes || ["https://outlook.office365.com/.default"];
    const login = await app.loginPopup({ scopes, loginHint: upn });
    app.setActiveAccount(login.account);
    state.msal = app; state.msalAccount = login.account;
    state.token = login.accessToken;
    state.tokenExp = (login.expiresOn ? login.expiresOn.getTime() : Date.now() + 50 * 60 * 1000);
    state.upn = login.account.username || upn;
    state.org = $("org").value.trim() || domain;
    $("org").value = state.org;
    persistSession(domain);
    // Second token for the Microsoft Graph license sections (different
    // resource: silent when consented, popup otherwise). Failure only blocks
    // the Licensing category, not the Exchange sections.
    try {
      state.graphToken = ""; state.graphTokenExp = 0;
      await ensureGraphToken();
    } catch (e) { log(t("signinErrorLog", { err: "Graph token: " + (e.message || e) })); }
    setConnected(state.org + " (" + state.upn + ")");
    log(t("signedInLog", { upn: maskId(state.upn), org: maskHost(state.org) }));
  } catch (e) {
    $("homeStatus").textContent = t("signinFailed", { err: e.message || e });
    log(t("signinErrorLog", { err: e.message || e }));
  }
};
async function ensureToken() {
  if (state.msal && state.msalAccount && Date.now() > state.tokenExp - 5 * 60 * 1000) {
    const tok = await state.msal.acquireTokenSilent({
      scopes: EAT_CFG.exoScopes || ["https://outlook.office365.com/.default"],
      account: state.msalAccount
    });
    state.token = tok.accessToken;
    state.tokenExp = (tok.expiresOn ? tok.expiresOn.getTime() : Date.now() + 50 * 60 * 1000);
    log(t("tokenRefreshedLog"));
  }
  // Graph refresh stays best-effort here: a missing/expired Graph token must
  // never block the Exchange sections. License runs re-acquire interactively.
  try { await ensureGraphToken(true); } catch (e) { log(t("tokenRefreshFailedLog", { err: "Graph: " + (e.message || e) })); }
}
async function ensureGraphToken(silentOnly) {
  // Graph license token: refresh when missing/expiring. Interactive popup only
  // during an explicit Connect (silentOnly falsy); background refreshes stay
  // silent so Exchange polling never pops a window.
  if (!state.msal || !state.msalAccount) return;
  if (state.graphToken && Date.now() < state.graphTokenExp - 5 * 60 * 1000) return;
  const scopes = EAT_CFG.graphScopes || ["User.Read.All", "Organization.Read.All", "Group.Read.All", "Team.ReadBasic.All", "Channel.ReadBasic.All", "TeamMember.Read.All", "ChannelMember.Read.All", "TeamsAppInstallation.ReadForTeam", "TeamsTab.Read.All", "Reports.Read.All", "ChannelMessage.Read.All", "Sites.Read.All", "SharePointTenantSettings.Read.All", "AuditLog.Read.All", "UserAuthenticationMethod.Read.All", "RoleManagement.Read.Directory"];
  try {
    const tok = await state.msal.acquireTokenSilent({ scopes, account: state.msalAccount });
    state.graphToken = tok.accessToken;
    state.graphTokenExp = (tok.expiresOn ? tok.expiresOn.getTime() : Date.now() + 50 * 60 * 1000);
  } catch (e) {
    if (silentOnly) throw e;
    const tok = await state.msal.loginPopup({ scopes });
    state.graphToken = tok.accessToken;
    state.graphTokenExp = (tok.expiresOn ? tok.expiresOn.getTime() : Date.now() + 50 * 60 * 1000);
    try { state.msal.setActiveAccount(tok.account || state.msalAccount); } catch (e2) {}
  }
}
$("disconnectBtn").onclick = () => {
  state.msal = null;
  clearPersistedSession();
  setDisconnected();
  log(t("disconnectedLog"));
};
$("dlCsv").onclick = () => window.open(API + "/api/jobs/" + state.jobId + "/download?format=csv", "_blank");
$("dlXlsx").onclick = () => window.open(API + "/api/jobs/" + state.jobId + "/download?format=xlsx", "_blank");
document.querySelectorAll(".nav-item[data-view]").forEach(b => b.onclick = () => showView(b.dataset.view));

// Top search filters the audit sections in the left nav (Enter opens the first match).
function applyTopSearch() {
  const q = ($("topSearch").value || "").trim().toLowerCase();
  document.querySelectorAll("#navGroups .nav-block").forEach(block => {
    const items = [...block.querySelectorAll(".nav-item")];
    let vis = 0;
    items.forEach(b => {
      const show = !q || b.textContent.toLowerCase().includes(q);
      b.style.display = show ? "" : "none";
      if (show) vis++;
    });
    block.style.display = (vis || !q) ? "" : "none";
    setCat(block, q ? vis > 0 : true);
  });
}
$("topSearch").oninput = () => applyTopSearch();
$("topSearch").onkeydown = e => {
  if (e.key === "Enter") {
    const first = [...document.querySelectorAll("#navGroups .nav-item")].find(b => b.style.display !== "none");
    if (first) first.click();
  }
};

applyI18n();
loadSections().then(restoreSession).then(restoreJob).catch(e => { log(t("apiDownLog", { err: e.message })); $("coverage").textContent = t("apiDownCoverage"); });
updateRegisterLink();
