FROM mcr.microsoft.com/powershell:7.4-ubuntu-22.04
# run_job.py is stdlib-only but the image ships without python3.
RUN apt-get update && apt-get install -y --no-install-recommends python3 curl \
    && rm -rf /var/lib/apt/lists/* \
    && pwsh -NoProfile -Command "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; Install-Module ExchangeOnlineManagement -Scope AllUsers -Force -AllowClobber"
WORKDIR /worker
COPY run_job.py .
# Baked-in Microsoft licensing display-name database (Entra "product names and
# service plan identifiers" CSV). Audits try the live URL first (always
# current); this copy is the offline fallback when the download fails, and raw
# IDs are the last resort (see Sections.Licenses.cs).
# Freshness: Docker caches this layer, so a plain rebuild does NOT re-check.
# Bump CSV_REFRESH_TOKEN at build time to force an online check, e.g.
#   CSV_REFRESH_TOKEN=$(date +%Y%m%d) docker compose up -d --build worker
# The download itself is conditional (If-Modified-Since): if Microsoft hasn't
# changed the file, the server answers 304 and nothing is re-downloaded.
# The build itself never fails if the download is down.
ARG CSV_REFRESH_TOKEN=2026-01
RUN echo "licensing CSV refresh token: $CSV_REFRESH_TOKEN" && \
    ONLINE_DATE=$(curl -sI --max-time 60 "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | grep -i '^last-modified:' | tr -d '\r' || true) && \
    echo "online file: ${ONLINE_DATE:-last-modified unknown}" && \
    if [ -f licensing-names.csv ]; then COND="-z licensing-names.csv"; ls -l --time-style=long-iso licensing-names.csv; else COND=""; echo "no baked copy yet, full download"; fi && \
    curl -fSL --retry 2 --max-time 120 $COND "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" -o licensing-names.csv && \
    ls -la licensing-names.csv || echo "WARNING: licensing-names.csv not baked in; license audits fall back to live download / raw IDs"
ENV EAT_DATA_DIR=/data EAT_CONNECT_TIMEOUT_SEC=300 EAT_AUDIT_TIMEOUT_SEC=1800 EAT_DEMO_MODE=false
VOLUME /data
EXPOSE 8081
CMD ["python3", "run_job.py"]
