# ============================================
# SPO Permissions Analyzer - Container Image
# ============================================
# Web UI on port 8080
#
# Build:  podman build -t spo-analyzer .
# Run:    podman run -p 8080:8080 spo-analyzer

FROM mcr.microsoft.com/powershell:7.4-ubuntu-22.04

# Install xsel for clipboard support (required by PnP PowerShell on Linux)
RUN apt-get update && apt-get install -y xsel && rm -rf /var/lib/apt/lists/*

# Install PnP.PowerShell module
RUN pwsh -Command "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; \
    Install-Module -Name PnP.PowerShell -Scope AllUsers -Force -AcceptLicense"

WORKDIR /app

# Copy application files
COPY Functions/ ./Functions/
COPY Web/ ./Web/
COPY Start-SPOTool-Web.ps1 .
COPY docker-entrypoint.ps1 .

# Create runtime directories
RUN mkdir -p /app/Logs /app/Reports/Generated

EXPOSE 8080

ENV SPO_HEADLESS=true

ENTRYPOINT ["pwsh", "-File", "/app/docker-entrypoint.ps1"]
