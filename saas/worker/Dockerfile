FROM mcr.microsoft.com/powershell:7.4-ubuntu-22.04
RUN pwsh -NoProfile -Command "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; Install-Module ExchangeOnlineManagement -Scope AllUsers -Force -AllowClobber"
WORKDIR /worker
COPY run_job.py .
ENV EAT_DATA_DIR=/data EAT_CONNECT_TIMEOUT_SEC=300 EAT_AUDIT_TIMEOUT_SEC=1800 EAT_DEMO_MODE=false
VOLUME /data
EXPOSE 8081
CMD ["python3", "run_job.py"]
