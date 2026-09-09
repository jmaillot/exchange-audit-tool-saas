#!/usr/bin/env python3
"""Exchange Audit worker: executes the API-generated audit script with pwsh.

POST /run {jobId, organization, script, includeXlsx}
  Authorization: Bearer <EXO delegated access token for https://outlook.office365.com>
  Writes /data/<jobId>.csv (+ .log). Token is kept in memory only, never logged.

GET /health -> {ok, demoMode, exoModule}

EAT_DEMO_MODE=true generates a fake CSV from the Export-Csv column list so the
UI can be tested without a tenant. Protection sections use Connect-IPPSSession.
"""
import csv
import json
import os
import re
import subprocess
import tempfile
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs

DATA = os.environ.get("EAT_DATA_DIR", "/data")
DEMO = os.environ.get("EAT_DEMO_MODE", "false").lower() == "true"
AUDIT_TIMEOUT = int(os.environ.get("EAT_AUDIT_TIMEOUT_SEC", "1800"))
os.makedirs(DATA, exist_ok=True)


def redact(text):
    return re.sub(r"-AccessToken\s+\S+", "-AccessToken ***", text or "", flags=re.I)


def demo_csv(job_id, script):
    # Extract selected column names from the generated `Select-Object a,b,c` tail.
    cols = ["DisplayName", "PrimarySmtpAddress"]
    m = re.search(r"Select-Object\s+([^|\n]+?)\s*\|", script)
    if m:
        raw = [c.strip().strip('"').strip("'") for c in m.group(1).split(",")]
        cols = [c for c in raw if c and not c.startswith("@")] or cols
    path = os.path.join(DATA, job_id + ".csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(cols)
        for i in range(1, 4):
            w.writerow([f"demo-{c.lower()}-{i}" for c in cols])
    with open(os.path.join(DATA, job_id + ".log"), "w", encoding="utf-8") as f:
        f.write("DEMO MODE: no Exchange connection, 3 sample rows written.\n")


def read_tail(job_id, max_chars=6000):
    try:
        with open(os.path.join(DATA, job_id + ".log"), encoding="utf-8") as f:
            t = f.read()
        return t[-max_chars:] if len(t) > max_chars else t
    except OSError:
        return ""


def run_real(job_id, org, script, token):
    log_path = os.path.join(DATA, job_id + ".log")
    csv_path = os.path.join(DATA, job_id + ".csv")
    is_protection = "Get-HostedContentFilterPolicy" in script or "Get-MalwareFilterPolicy" in script \
        or "Get-AntiPhishPolicy" in script or "Get-SafeLinksPolicy" in script or "Get-DlpPolicy" in script
    connect = (
        "Import-Module ExchangeOnlineManagement -ErrorAction Stop\n"
        "$ProgressPreference='SilentlyContinue'; $ErrorActionPreference='Stop'\n"
        "Connect-ExchangeOnline -AccessToken $env:EAT_TOKEN -Organization $env:EAT_ORG -ShowBanner:$false\n"
    )
    if is_protection:
        connect += "Connect-IPPSSession -AccessToken $env:EAT_TOKEN -Organization $env:EAT_ORG\n"
    wrapper = connect + "\n" + script + "\nDisconnect-ExchangeOnline -Confirm:$false\n"
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as tf:
        tf.write(wrapper)
        ps1 = tf.name
    env = dict(os.environ, EAT_TOKEN=token, EAT_ORG=org)
    try:
        with open(log_path, "w", encoding="utf-8") as log:
            p = subprocess.run(["pwsh", "-NoProfile", "-File", ps1],
                               capture_output=True, text=True, timeout=AUDIT_TIMEOUT, env=env)
            log.write((p.stdout or "") + ("\n" + p.stderr if p.stderr else ""))
            log.write(f"\n[exit={p.returncode}]\n")
            if p.returncode != 0:
                raise RuntimeError(f"pwsh exit {p.returncode}")
            if not os.path.exists(csv_path):
                raise RuntimeError("audit produced no CSV")
    finally:
        try:
            os.unlink(ps1)
        except OSError:
            pass


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/health":
            self._json(200, {"ok": True, "demoMode": DEMO})
            return
        # Serve job outputs (the API pulls them; storages are not shared).
        u = urlparse(self.path)
        if u.path.startswith("/file/"):
            job_id = re.sub(r"[^a-z0-9]", "", u.path[len("/file/"):].lower())
            kind = parse_qs(u.query).get("kind", [""])[0]
            if not job_id or kind not in ("csv", "log"):
                self._json(422, {"error": "job and kind=csv|log required"})
                return
            path = os.path.join(DATA, job_id + (".csv" if kind == "csv" else ".log"))
            if not os.path.exists(path):
                self._json(404, {"error": "not found"})
                return
            try:
                with open(path, "rb") as f:
                    body = f.read()
            except OSError as ex:
                self._json(500, {"error": str(ex)})
                return
            self.send_response(200)
            self.send_header("Content-Type", "text/csv" if kind == "csv" else "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self._json(404, {"error": "not found"})

    def do_DELETE(self):
        u = urlparse(self.path)
        if u.path.startswith("/file/"):
            job_id = re.sub(r"[^a-z0-9]", "", u.path[len("/file/"):].lower())
            if not job_id:
                self._json(422, {"error": "job required"})
                return
            for ext in (".csv", ".log"):
                try:
                    os.unlink(os.path.join(DATA, job_id + ext))
                except OSError:
                    pass
            self._json(200, {"ok": True})
            return
        self._json(404, {"error": "not found"})

    def do_POST(self):
        if self.path != "/run":
            self._json(404, {"error": "not found"})
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
        except ValueError:
            length = 0
        try:
            payload = json.loads(self.rfile.read(length) or b"{}")
        except json.JSONDecodeError:
            self._json(400, {"error": "invalid JSON"})
            return
        auth = self.headers.get("Authorization", "")
        token = auth[7:].strip() if auth.lower().startswith("bearer ") else ""
        job_id = re.sub(r"[^a-z0-9]", "", str(payload.get("jobId", "")).lower())
        org = str(payload.get("organization", ""))
        script = str(payload.get("script", ""))
        if not job_id or not script or not org:
            self._json(422, {"error": "jobId, organization and script are required"})
            return
        if not token and not DEMO:
            self._json(401, {"error": "missing Bearer EXO access token"})
            return
        try:
            if DEMO:
                demo_csv(job_id, script)
            else:
                run_real(job_id, org, script, token)
            self._json(200, {"ok": True})
        except subprocess.TimeoutExpired:
            self._json(500, {"error": "audit timed out", "log": read_tail(job_id)})
        except Exception as ex:
            self._json(500, {"error": str(ex), "log": read_tail(job_id)})
        finally:
            token = ""


if __name__ == "__main__":
    HTTPServer(("0.0.0.0", 8081), Handler).serve_forever()
