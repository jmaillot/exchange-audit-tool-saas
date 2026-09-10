/* Exchange Audit SaaS - Azure Portal style blade. No build step. */
const API = "";
const state = { sections: [], current: null, checks: {}, smartMode: true, token: "", tokenExp: 0, org: "", upn: "", jobId: null, jobSection: null, jobStart: 0, poll: null, msal: null, msalAccount: null };
// Defaults; /api/config (backed by saas/.env) overrides, web/config.js is the fallback.
const EAT_CFG = Object.assign(
  { clientId: "", msalSources: ["./msal-browser.min.js"], exoScopes: ["https://outlook.office365.com/.default"] },
  window.EAT_CONFIG || {});
fetch("api/config").then(r => r.json()).then(c => {
  if (c.clientId) EAT_CFG.clientId = c.clientId;
  if (c.msalSources && c.msalSources.length) EAT_CFG.msalSources = c.msalSources;
  if (c.exoScopes && c.exoScopes.length) EAT_CFG.exoScopes = c.exoScopes;
  updateRegisterLink();
}).catch(() => {});
const $ = id => document.getElementById(id);
const logEl = () => $("activityLog");

function log(msg) {
  const t = new Date().toISOString().slice(11, 19);
  logEl().textContent += `[${t}] ${msg}\n`;
  logEl().scrollTop = 1e9;
}
function authHeaders() {
  return { "Content-Type": "application/json", "Authorization": "Bearer " + state.token, "X-Tenant-Id": state.org };
}

async function loadSections() {
  const r = await fetch(API + "/api/sections");
  if (!r.ok) throw new Error("API " + r.status);
  state.sections = await r.json();
  renderNav(); renderCoverage();
  log(`Loaded ${state.sections.length} Exchange Online sections from API.`);
}

function setCat(block, open) {
  block.querySelector(".nav-items").style.display = open ? "" : "none";
  const c = block.querySelector(".nav-cat");
  c.setAttribute("aria-expanded", open ? "true" : "false");
  c.querySelector(".caret").textContent = open ? "\u25BE" : "\u25B8";
}
function renderNav() {
  const host = $("navGroups"); host.innerHTML = "";
  const cats = {};
  state.sections.forEach(s => { (cats[s.category || "Other"] ||= []).push(s); });
  Object.keys(cats).sort().forEach(cat => {
    const block = document.createElement("div");
    block.className = "nav-block"; block.dataset.cat = cat;
    const t = document.createElement("button");
    t.className = "nav-cat";
    t.innerHTML = `<span class="caret"></span>`;
    t.appendChild(document.createTextNode(cat + " "));
    const n = document.createElement("span");
    n.className = "nav-count"; n.textContent = cats[cat].length;
    t.appendChild(n);
    const wrap = document.createElement("div");
    wrap.className = "nav-items";
    cats[cat].forEach(s => {
      const b = document.createElement("button");
      b.className = "nav-item"; b.textContent = s.navTitle; b.dataset.section = s.id;
      b.onclick = () => openSection(s.id);
      wrap.appendChild(b);
    });
    t.onclick = () => setCat(block, wrap.style.display === "none");
    block.appendChild(t); block.appendChild(wrap); host.appendChild(block);
    setCat(block, false);
  });
  document.querySelectorAll(".nav-item[data-view]").forEach(b => b.onclick = () => showView(b.dataset.view));
}

function renderCoverage() {
  $("coverage").innerHTML = state.sections.map(s => `<div>${esc(s.navTitle)}</div>`).join("");
}

function showView(v) {
  ["home", "section", "activity"].forEach(x => $("view-" + x).classList.toggle("hidden", x !== v));
  document.querySelectorAll(".nav-item").forEach(b => b.classList.toggle("active", b.dataset.view === v));
  if (v !== "section") document.querySelectorAll("#navGroups .nav-item").forEach(b => b.classList.remove("active"));
}

function openSection(id) {
  try {
  const s = state.sections.find(x => x.id === id); if (!s) return;
  // Reopening the section of a known job resumes its live view instead of
  // wiping it (the job keeps running server-side either way).
  const resume = state.jobId && state.jobSection === id;
  state.current = s;
  if (!resume) state.checks = {};
  showView("section");
  document.querySelectorAll("#navGroups .nav-item").forEach(b => b.classList.toggle("active", b.dataset.section === id));
  $("crumbSection").textContent = s.navTitle;
  $("secTitle").textContent = s.title;
  $("secSub").textContent = s.subtitle || "";
  $("filter").value = "";
  renderGroups("");
  updateSlow();
  if (resume) { if (!state.poll) pollStart(); else fetchJobStatus(); }
  else { pollStop(); resetResults(); state.jobSection = null; state.jobStart = 0; }
  } catch (e) { log("Failed to open section: " + (e.message || e)); }
}

function renderGroups(filter) {
  const host = $("groups"); host.innerHTML = "";
  const f = (filter || "").toLowerCase();
  // Smart mode block (same look as transport-rules' native one). Sections that
  // already define their own (key "auto") keep it; the flag is read in runAudit.
  const nativeSmart = state.current.groups.some(g => g.key === "auto" || (g.title || "").toLowerCase().includes("smart"));
  if (!nativeSmart) {
    const label = "Auto-detect populated properties only (recommended)";
    if (!f || "smart mode".includes(f) || label.toLowerCase().includes(f)) {
      const card = document.createElement("div"); card.className = "card grp";
      card.innerHTML = `<h3>Smart mode</h3><div class="hint">Keeps only columns that have a value on at least one row.</div>`;
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
          renderGroups($("filter").value);
        } else state.checks[key] = inp.checked;
        updateSlow(); updateSelectAll();
      };
      lbl.appendChild(inp);
      lbl.appendChild(document.createTextNode(o.label + " "));
      if (o.slow || g.slow) { const em = document.createElement("span"); em.className = "slow"; em.textContent = "slow"; lbl.appendChild(em); }
      opts.appendChild(lbl);
    });
    card.appendChild(opts); host.appendChild(card);
  });
  updateSelectAll();
}

function updateSlow() {
  const slow = state.current.groups.some(g => (g.options || []).some(o => (o.slow || g.slow) && state.checks[g.key + "::" + o.value]));
  $("slowBadge").classList.toggle("hidden", !slow);
}
function updateSelectAll() {
  const keys = Object.keys(state.checks).filter(k => {
    const [gk, v] = k.split("::");
    const g = state.current.groups.find(x => x.key === gk);
    const o = g && ((g.options || []).find(x => x.value === v));
    return o && o.label.toLowerCase().includes(($("filter").value || "").toLowerCase()) && g.mode !== "SingleChoice";
  });
  const on = keys.filter(k => state.checks[k]).length;
  $("selectAll").textContent = keys.length && on === keys.length ? "Deselect all" : `Select all (${on}/${keys.length})`;
}
function selection() {
  const sel = {};
  state.current.groups.forEach(g => {
    sel[g.key] = (g.options || []).filter(o => state.checks[g.key + "::" + o.value]).map(o => o.value);
  });
  return sel;
}
function resetResults() {
  $("grid").innerHTML = ""; $("resultInfo").textContent = "No results yet.";
  $("jobLog").textContent = ""; $("jobLog").classList.add("hidden");
  $("dlCsv").disabled = $("dlXlsx").disabled = true;
}

async function runAudit() {
  if (!state.token || !state.org) { alert("Connect first (token + tenant organization)."); showView("home"); return; }
  try { await ensureToken(); } catch (e) { $("resultInfo").textContent = "Session expired, please reconnect."; log("Token refresh failed: " + (e.message || e)); return; }
  const sel = selection();
  const nativeAuto = state.current.groups.some(g => g.key === "auto");
  const smart = nativeAuto ? (sel.auto || []).includes("autodetect") : state.smartMode;
  const body = { sectionId: state.current.id, selection: sel, organization: state.org, includeXlsx: true, smartMode: smart };
  log(`RUN ${body.sectionId} selection=${JSON.stringify(body.selection)}`);
  $("runBtn").disabled = true; $("cancelBtn").disabled = false;
  $("resultInfo").textContent = "Queued...";
  const r = await fetch(API + "/api/jobs", { method: "POST", headers: authHeaders(), body: JSON.stringify(body) });
  const j = await r.json();
  if (!r.ok) { $("resultInfo").textContent = "Error: " + (j.error || r.status); log("FAILED: " + JSON.stringify(j)); $("runBtn").disabled = false; $("cancelBtn").disabled = true; return; }
  state.jobId = j.jobId;
  state.jobSection = state.current.id;
  state.jobStart = Date.now();
  log("Job " + j.jobId + " started.");
  pollStart();
}
function fmtElapsed(ms) {
  const s = Math.max(0, Math.floor(ms / 1000));
  const p = n => String(n).padStart(2, "0");
  return p(Math.floor(s / 3600)) + ":" + p(Math.floor(s % 3600 / 60)) + ":" + p(s % 60);
}
async function fetchJobStatus() {
  const r = await fetch(API + "/api/jobs/" + state.jobId, { headers: { "X-Tenant-Id": state.org } });
  const j = await r.json();
  const elapsed = state.jobStart ? fmtElapsed(Date.now() - state.jobStart) : "--:--:--";
  const active = j.status === "running" || j.status === "queued";
  const note = `${active ? "Running... " + elapsed : "Finished in " + elapsed} (job ${state.jobId})${j.error ? " - " + j.error : ""}`;
  if (j.header) renderGrid(j.header, j.preview || [], note);
  else $("resultInfo").textContent = note;
  if (j.log) { $("jobLog").textContent = j.log.slice(-8000); $("jobLog").classList.remove("hidden"); }
  if (j.status === "succeeded" || j.status === "failed") {
    pollStop(); $("runBtn").disabled = false; $("cancelBtn").disabled = true;
    $("dlCsv").disabled = !j.hasCsv; $("dlXlsx").disabled = !j.hasXlsx;
    log(`Job ${state.jobId} ${j.status} in ${elapsed}${j.header ? " (" + j.header.length + " columns)" : ""}.`);
  }
}
function pollStart() {
  pollStop();
  fetchJobStatus();
  state.poll = setInterval(fetchJobStatus, 3000);
}
function pollStop() { if (state.poll) clearInterval(state.poll); state.poll = null; }
function renderGrid(header, rows, note) {
  const t = $("grid"); t.innerHTML = "";
  const trh = document.createElement("tr");
  header.forEach(h => { const th = document.createElement("th"); th.textContent = h; th.title = h; trh.appendChild(th); });
  t.appendChild(trh);
  rows.forEach(r => {
    const tr = document.createElement("tr");
    r.forEach(c => { const td = document.createElement("td"); td.textContent = c ?? ""; td.title = c ?? ""; tr.appendChild(td); });
    t.appendChild(tr);
  });
  $("resultInfo").textContent = `${header.length} columns x ${rows.length} rows (preview of first 200). ` + (note || "");
}
function esc(s) { return String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

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
$("cancelBtn").onclick = () => { pollStop(); $("runBtn").disabled = false; $("cancelBtn").disabled = true; log("Polling stopped (worker job continues to timeout)."); };
$("crumbHome").onclick = e => { e.preventDefault(); showView("home"); };
$("hamburger").onclick = () => $("sidenav").classList.toggle("hidden");
function setConnected(label) {
  $("connDot").classList.add("on"); $("connText").textContent = "Connected: " + label;
  $("homeStatus").textContent = "Connected (token in memory only).";
}
function setDisconnected(msg) {
  state.token = ""; state.tokenExp = 0; state.msalAccount = null; state.jobId = null; pollStop();
  $("connDot").classList.remove("on"); $("connText").textContent = "Not connected";
  $("homeStatus").textContent = msg || "Disconnected, token dropped.";
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
$("connectBtn").onclick = async () => {
  const upn = $("upn").value.trim();
  const domain = upnDomain(upn);
  if (!upn || !domain) { $("homeStatus").textContent = "Enter your work email (UPN)."; return; }
  if (!EAT_CFG.clientId || EAT_CFG.clientId.indexOf("PASTE") === 0) {
    $("homeStatus").textContent = "Server not configured: set clientId in web/config.js.";
    return;
  }
  if (!window.msal) {
    for (const src of (EAT_CFG.msalSources || ["./msal-browser.min.js"])) {
      try { await loadScript(src); if (window.msal) break; }
      catch (e) { log("MSAL load failed: " + src); }
    }
  }
  if (!window.msal) {
    $("homeStatus").textContent = "Microsoft login unavailable — please contact your administrator.";
    return;
  }
  try {
    $("homeStatus").textContent = "Opening Microsoft sign-in...";
    // Tokens cached in memory only (never localStorage/sessionStorage).
    const app = new window.msal.PublicClientApplication({
      auth: { clientId: EAT_CFG.clientId, authority: "https://login.microsoftonline.com/" + domain },
      cache: { cacheLocation: "memory" }
    });
    const scopes = EAT_CFG.exoScopes || ["https://outlook.office365.com/.default"];
    const login = await app.loginPopup({ scopes, loginHint: upn });
    app.setActiveAccount(login.account);
    state.msal = app; state.msalAccount = login.account;
    state.token = login.accessToken;
    state.tokenExp = (login.expiresOn ? login.expiresOn.getTime() : Date.now() + 50 * 60 * 1000);
    state.upn = login.account.username || upn;
    state.org = $("org").value.trim() || domain;
    $("org").value = state.org;
    setConnected(state.org + " (" + state.upn + ")");
    log("Signed in as " + state.upn + " (tenant " + state.org + "). Token held in memory.");
  } catch (e) {
    $("homeStatus").textContent = "Sign-in failed: " + (e.message || e);
    log("Sign-in error: " + (e.message || e));
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
    log("Token refreshed silently.");
  }
}
$("disconnectBtn").onclick = () => {
  state.msal = null;
  setDisconnected();
  log("Disconnected, token dropped.");
};
$("dlCsv").onclick = () => window.open(API + "/api/jobs/" + state.jobId + "/download?format=csv", "_blank");
$("dlXlsx").onclick = () => window.open(API + "/api/jobs/" + state.jobId + "/download?format=xlsx", "_blank");
document.querySelectorAll(".nav-item[data-view]").forEach(b => b.onclick = () => showView(b.dataset.view));

// Top search filters the audit sections in the left nav (Enter opens the first match).
$("topSearch").oninput = e => {
  const q = e.target.value.trim().toLowerCase();
  document.querySelectorAll("#navGroups .nav-block").forEach(block => {
    const items = [...block.querySelectorAll(".nav-item")];
    let vis = 0;
    items.forEach(b => {
      const show = !q || b.textContent.toLowerCase().includes(q);
      b.style.display = show ? "" : "none";
      if (show) vis++;
    });
    block.style.display = (vis || !q) ? "" : "none";
    setCat(block, q ? vis > 0 : false);
  });
};
$("topSearch").onkeydown = e => {
  if (e.key === "Enter") {
    const first = [...document.querySelectorAll("#navGroups .nav-item")].find(b => b.style.display !== "none");
    if (first) first.click();
  }
};

loadSections().catch(e => { log("API unreachable: " + e.message); $("coverage").textContent = "API unreachable."; });
updateRegisterLink();
