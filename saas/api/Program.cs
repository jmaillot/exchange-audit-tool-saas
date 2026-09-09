using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using ExchangeAuditTool;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddCors(o => o.AddDefaultPolicy(p => p.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod()));
var app = builder.Build();
app.UseCors();

string dataDir = Environment.GetEnvironmentVariable("EAT_DATA_DIR") ?? "/data";
string workerUrl = Environment.GetEnvironmentVariable("EAT_WORKER_URL") ?? "http://worker:8081";
Directory.CreateDirectory(dataDir);

AuditRegistry.BuildAll();
// SaaS = Exchange Online only. OnPremises-only sections are hidden.
var onlineSections = AuditRegistry.Sections.Where(s => s.Scope != AuditScope.OnPremises).ToList();

var jobs = new ConcurrentDictionary<string, JobRecord>();

string Redact(string s)
{
    if (string.IsNullOrEmpty(s)) return s;
    return System.Text.RegularExpressions.Regex.Replace(s, "-AccessToken\\s+\\S+", "-AccessToken ***",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
}

app.MapGet("/api/health", () => Results.Json(new { ok = true, sections = onlineSections.Count }));

// Public web config (no secrets): clientId + login library sources + scopes,
// read from environment so the NAS admin only edits saas/.env.
app.MapGet("/api/config", () =>
{
    string scopes = Environment.GetEnvironmentVariable("EAT_EXO_SCOPES") ?? "https://outlook.office365.com/.default";
    var sources = new List<string>();
    string envSrc = Environment.GetEnvironmentVariable("EAT_MSAL_SRC") ?? "";
    if (!string.IsNullOrWhiteSpace(envSrc)) sources.Add(envSrc.Trim());
    sources.Add("./msal-browser.min.js"); // optional vendored copy next to index.html
    sources.Add("https://alcdn.msauth.net/browser/2.30.0/js/msal-browser.min.js"); // Microsoft CDN
    return Results.Json(new
    {
        clientId = (Environment.GetEnvironmentVariable("EAT_CLIENT_ID") ?? "").Trim(),
        msalSources = sources,
        exoScopes = scopes.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
    });
});

// Section catalog: drives the web checkboxes. Mirrors the AuditOptionGroup model.
app.MapGet("/api/sections", () =>
{
    var list = onlineSections.Select(s => new
    {
        id = s.Id,
        navTitle = s.NavTitle,
        title = s.Title,
        subtitle = s.Subtitle,
        category = s.Category,
        scope = s.Scope.ToString(),
        defaultFileName = s.DefaultFileName,
        groups = s.Groups.Select(g => new
        {
            key = g.Key,
            title = g.Title,
            hint = g.Hint,
            mode = g.Mode.ToString(),
            columns = g.Columns,
            slow = g.Slow,
            options = g.Options.Select(o => new
            {
                key = o.Key,
                label = o.Label,
                value = o.Value,
                // SaaS always uses Online defaults.
                defaultChecked = o.DefOnline,
                slow = o.Slow
            })
        })
    });
    return Results.Json(list);
});

app.MapPost("/api/jobs", async (HttpRequest req) =>
{
    string body;
    using (var sr = new StreamReader(req.Body)) body = await sr.ReadToEndAsync();
    using var doc = JsonDocument.Parse(body);
    var root = doc.RootElement;
    string sectionId = root.TryGetProperty("sectionId", out var sid) ? sid.GetString() ?? "" : "";
    var section = onlineSections.FirstOrDefault(s => s.Id == sectionId);
    if (section == null) return Results.Json(new { error = "Unknown or on-premises-only section: " + sectionId }, statusCode: 422);

    // Bearer token = EXO delegated access token for https://outlook.office365.com.
    // Never persisted, never logged. Forwarded to worker in-memory only.
    string auth = req.Headers.Authorization.ToString();
    string token = auth.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) ? auth.Substring(7).Trim() : "";
    string org = req.Headers.TryGetValue("X-Tenant-Id", out var tv) ? tv.ToString() : "";
    if (root.TryGetProperty("organization", out var orgProp)) org = orgProp.GetString() ?? org;
    if (string.IsNullOrEmpty(token)) return Results.Json(new { error = "Missing Authorization: Bearer <EXO access token>" }, statusCode: 401);
    if (string.IsNullOrEmpty(org)) return Results.Json(new { error = "Missing tenant organization (X-Tenant-Id header or organization field)" }, statusCode: 422);

    // Validate selection against allow-list. No raw PowerShell accepted from client.
    var sel = new AuditSelection();
    var allowed = section.Groups.ToDictionary(g => g.Key, g => g);
    if (root.TryGetProperty("selection", out var selProp) && selProp.ValueKind == JsonValueKind.Object)
    {
        foreach (var g in selProp.EnumerateObject())
        {
            if (!allowed.TryGetValue(g.Name, out var grp))
                return Results.Json(new { error = "Unknown group: " + g.Name }, statusCode: 422);
            var validValues = new HashSet<string>(grp.Options.Select(o => o.Value));
            var vals = new List<string>();
            foreach (var v in g.Value.EnumerateArray())
            {
                string s = v.GetString() ?? "";
                if (!validValues.Contains(s))
                    return Results.Json(new { error = $"Invalid value '{s}' for group '{g.Name}'" }, statusCode: 422);
                if (!vals.Contains(s)) vals.Add(s);
            }
            sel.Set(g.Name, vals);
        }
    }
    // Default single-choice groups to first option when client sends nothing.
    foreach (var g in section.Groups.Where(g => g.Mode == GroupMode.SingleChoice))
        if (sel.Selected(g.Key).Count == 0 && g.Options.Count > 0)
            sel.Set(g.Key, new List<string> { g.Options[0].Value });

    bool includeXlsx = true;
    if (root.TryGetProperty("includeXlsx", out var xlsx)) includeXlsx = xlsx.GetBoolean();

    bool anySelected = section.Groups.Any(g => sel.Selected(g.Key).Count > 0);
    if (!anySelected)
        return Results.Json(new { error = "Select at least one property before running the audit." }, statusCode: 422);

    bool smartMode = true;
    if (root.TryGetProperty("smartMode", out var sm)) smartMode = sm.GetBoolean();

    string jobId = Guid.NewGuid().ToString("N");
    string csvPath = Path.Combine(dataDir, jobId + ".csv");
    // Generate the audit body with the shared audit-DSL logic. CsvPath here is the
    // worker-side path (/data/<job>.csv, same shared volume).
    string auditBody;
    try
    {
        auditBody = section.BuildScript(sel, new ScriptContext(csvPath));
    }
    catch (Exception ex)
    {
        return Results.Json(new { error = "Script build failed: " + ex.Message }, statusCode: 422);
    }

    var job = new JobRecord { Id = jobId, SectionId = sectionId, Organization = org, Status = "queued", CreatedUtc = DateTime.UtcNow, IncludeXlsx = includeXlsx };
    jobs[jobId] = job;
    File.WriteAllText(Path.Combine(dataDir, jobId + ".script.ps1"), auditBody, new UTF8Encoding(true));

    // Dispatch to worker (fire and forget; status polled via GET).
    // Storages are NOT shared (tmpfs per container in production), so the
    // API pulls the outputs back over HTTP after a successful run.
    _ = Task.Run(async () =>
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromMinutes(32) };
        string baseUrl = workerUrl.TrimEnd('/');
        try
        {
            job.Status = "running";
            var payload = JsonSerializer.Serialize(new { jobId, organization = org, script = auditBody, includeXlsx });
            using var fwd = new HttpRequestMessage(HttpMethod.Post, baseUrl + "/run");
            fwd.Content = new StringContent(payload, Encoding.UTF8, "application/json");
            fwd.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            var resp = await http.SendAsync(fwd);
            string respText = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode)
            {
                string err = respText, tail = "";
                try
                {
                    using var d = JsonDocument.Parse(respText);
                    if (d.RootElement.TryGetProperty("error", out var e)) err = e.GetString() ?? err;
                    if (d.RootElement.TryGetProperty("log", out var l)) tail = l.GetString() ?? "";
                }
                catch { }
                job.Status = "failed";
                job.Error = "worker " + (int)resp.StatusCode + ": " + err;
                job.Log = tail;
                return;
            }
            try
            {
                using var csvResp = await http.GetAsync(baseUrl + "/file/" + jobId + "?kind=csv",
                    HttpCompletionOption.ResponseHeadersRead);
                csvResp.EnsureSuccessStatusCode();
                using var fs = File.Create(csvPath);
                await (await csvResp.Content.ReadAsStreamAsync()).CopyToAsync(fs);
            }
            catch (Exception ex)
            {
                job.Status = "failed";
                job.Error = "fetch CSV from worker: " + ex.Message;
                return;
            }
            try { job.Log = await http.GetStringAsync(baseUrl + "/file/" + jobId + "?kind=log"); }
            catch { }
            if (smartMode)
            {
                var dropped = PruneEmptyColumns(csvPath);
                if (dropped.Count > 0)
                    job.Log += "\n[smart-mode] dropped " + dropped.Count + " empty column(s): " + string.Join(",", dropped);
            }
            try { await http.DeleteAsync(baseUrl + "/file/" + jobId); }
            catch { }
            job.Status = "succeeded";
            if (includeXlsx) TryConvertToXlsx(csvPath);
        }
        catch (Exception ex)
        {
            job.Status = "failed";
            job.Error = ex.Message;
        }
        finally { token = ""; } // drop reference ASAP
    });

    return Results.Json(new { jobId, status = job.Status });
});

app.MapGet("/api/jobs/{id}", (string id) =>
{
    if (!jobs.TryGetValue(id, out var job)) return Results.NotFound();
    string csv = Path.Combine(dataDir, id + ".csv");
    var preview = new List<string[]>();
    string[] header = Array.Empty<string>();
    long rows = 0;
    if (File.Exists(csv))
    {
        try
        {
            using var sr = new StreamReader(csv, Encoding.UTF8);
            string hl = sr.ReadLine() ?? "";
            header = ParseDelimited(hl, ';').ToArray();
            string line;
            while ((line = sr.ReadLine()) != null && preview.Count < 200)
            {
                rows++;
                preview.Add(ParseDelimited(line, ';').ToArray());
            }
            // Count remainder for the label.
            while ((line = sr.ReadLine()) != null) rows++;
        }
        catch { }
    }
    // Worker-side log: attached by the dispatcher, else shared-volume tail.
    string log = job.Log ?? "";
    string logPath = Path.Combine(dataDir, id + ".log");
    if (string.IsNullOrEmpty(log) && File.Exists(logPath)) { try { var t = File.ReadAllText(logPath); log = t.Length > 8000 ? t.Substring(t.Length - 8000) : t; } catch { } }
    return Results.Json(new
    {
        jobId = id,
        sectionId = job.SectionId,
        status = job.Status,
        error = job.Error,
        header,
        preview,
        previewTruncatedAt = 200,
        hasCsv = File.Exists(csv),
        hasXlsx = File.Exists(Path.Combine(dataDir, id + ".xlsx"))
    });
});

app.MapGet("/api/jobs/{id}/download", (string id, string format) =>
{
    if (!jobs.TryGetValue(id, out var job)) return Results.NotFound();
    string ext = (format ?? "csv").ToLowerInvariant() == "xlsx" ? ".xlsx" : ".csv";
    string path = Path.Combine(dataDir, id + ext);
    if (!File.Exists(path)) return Results.NotFound();
    string ct = ext == ".xlsx"
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        : "text/csv";
    string baseName = "audit";
    var section = onlineSections.FirstOrDefault(s => s.Id == job.SectionId);
    if (section != null && !string.IsNullOrEmpty(section.DefaultFileName))
        baseName = Path.GetFileNameWithoutExtension(section.DefaultFileName);
    string stamp = File.GetLastWriteTimeUtc(path).ToString("yyyyMMdd-HHmmss");
    return Results.File(path, ct, baseName + "-" + stamp + ext);
});

app.Run();

// Smart mode (every section): drop CSV columns empty on all rows, mirroring
// the transport-rules auto-detect UX. Returns the dropped column names.
static List<string> PruneEmptyColumns(string csvPath)
{
    var dropped = new List<string>();
    string[] lines;
    try { lines = File.ReadAllLines(csvPath, Encoding.UTF8); }
    catch { return dropped; }
    if (lines.Length < 2) return dropped;
    var header = ParseDelimited(lines[0], ';');
    if (header.Count > 0 && header[0].Length > 0 && header[0][0] == '\uFEFF')
        header[0] = header[0].Substring(1); // keep BOM out of the first column name
    int cols = header.Count;
    if (cols == 0) return dropped;
    var rows = new List<List<string>>();
    foreach (string ln in lines.Skip(1))
    {
        var r = ParseDelimited(ln, ';');
        if (r.Count != cols) return new List<string>(); // unexpected shape: leave file untouched
        rows.Add(r);
    }
    var keep = new bool[cols];
    for (int c = 0; c < cols; c++)
    {
        foreach (var r in rows)
            if (!string.IsNullOrWhiteSpace(r[c])) { keep[c] = true; break; }
        if (!keep[c]) dropped.Add(header[c]);
    }
    if (dropped.Count == 0 || dropped.Count == cols) return new List<string>();
    var keptIdx = Enumerable.Range(0, cols).Where(c => keep[c]).ToList();
    var sb = new StringBuilder();
    sb.AppendLine(string.Join(";", keptIdx.Select(c => CsvCell(header[c]))));
    foreach (var r in rows)
        sb.AppendLine(string.Join(";", keptIdx.Select(c => CsvCell(r[c]))));
    try { File.WriteAllText(csvPath, sb.ToString(), new UTF8Encoding(true)); }
    catch { return new List<string>(); }
    return dropped;
}

static string CsvCell(string v)
{
    v ??= "";
    if (v.Contains(';') || v.Contains('"') || v.Contains('\r') || v.Contains('\n'))
        return "\"" + v.Replace("\"", "\"\"") + "\"";
    return v;
}

// Minimal ;-delimited parser (quotes per Export-Csv).
static List<string> ParseDelimited(string line, char delim)
{
    var cells = new List<string>();
    if (line == null) return cells;
    var cur = new StringBuilder();
    bool inQ = false;
    for (int i = 0; i < line.Length; i++)
    {
        char c = line[i];
        if (inQ)
        {
            if (c == '"')
            {
                if (i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; }
                else inQ = false;
            }
            else cur.Append(c);
        }
        else
        {
            if (c == '"') inQ = true;
            else if (c == delim) { cells.Add(cur.ToString()); cur.Clear(); }
            else cur.Append(c);
        }
    }
    cells.Add(cur.ToString());
    return cells;
}

// Dependency-free XLSX (single-sheet workbook: bold header, frozen top row,
// autofilter, auto-width). Implemented here with System.IO.Packaging so the
// API container stays dependency-free.
static void TryConvertToXlsx(string csvPath)
{
    try
    {
        string xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");
        ExchangeAuditSaaS.Xlsx.FromCsv(csvPath, xlsxPath, "Audit");
    }
    catch { }
}

sealed class JobRecord
{
    public string Id = "";
    public string SectionId = "";
    public string Organization = "";
    public string Status = "queued";
    public string Error = "";
    public string Log = "";
    public bool IncludeXlsx;
    public DateTime CreatedUtc;
}
