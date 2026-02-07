# SharePoint Online Permissions Analyzer  [![CodeFactor](https://www.codefactor.io/repository/github/creativeacer/spo-permissions-analyzer/badge)](https://www.codefactor.io/repository/github/creativeacer/spo-permissions-analyzer)

<div align="center">

![SharePoint](https://img.shields.io/badge/SharePoint-Online-0078D4?style=for-the-badge&logo=microsoft-sharepoint&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-7.0+-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![PnP](https://img.shields.io/badge/PnP-PowerShell_3.x-orange?style=for-the-badge)
![Container](https://img.shields.io/badge/Container-Podman%20%7C%20Docker-blue?style=for-the-badge&logo=podman&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

Analyze SharePoint Online permissions, users, groups, and security settings.
Runs as a **container**, **desktop app (WPF)**, or **local web server**.

[Quick Start](#quick-start) | [Container](#container-deployment) | [Features](#features) | [Screenshots](#screenshots) | [App Registration](#app-registration)

</div>

---

## Quick Start

### Option A: Container (recommended)

No local PowerShell or module installation required.

```bash
git clone https://github.com/CreativeAcer/SPO-Permissions-Analyzer.git
cd SPO-Permissions-Analyzer
podman compose up        # or: docker compose up
```

Open `http://localhost:8080` in your browser.

### Option B: Desktop (WPF)

Requires PowerShell 7+ and PnP.PowerShell on Windows.

```powershell
git clone https://github.com/CreativeAcer/SPO-Permissions-Analyzer.git
cd SPO-Permissions-Analyzer
.\Install-Prerequisites.ps1
.\Start-SPOTool.ps1
```

### Option C: Web UI (direct)

Same backend as the desktop app, browser-based frontend with Chart.js charts.

```powershell
.\Start-SPOTool-Web.ps1    # opens http://localhost:8080
```

### Demo Mode

All three options support **Demo Mode** — click the button on the Connection tab to explore all features with sample data, no SharePoint connection required.

---

## Container Deployment

### Prerequisites

[Podman](https://podman.io/) or [Docker](https://docs.docker.com/get-docker/)

### Commands

| Command | Description |
|---------|-------------|
| `podman compose up` | Web UI at localhost:8080 (default) |
| `podman compose down` | Stop the container |
| `podman build -t spo-analyzer .` | Build image only |
| `podman run -p 8080:8080 spo-analyzer` | Run without compose |

### Live SharePoint Connection

The container uses **device code flow** for authentication (no browser popup needed).

**Auto-connect on startup** — uncomment and set env vars in `compose.yaml`:

```yaml
environment:
  - SPO_TENANT_URL=https://yourtenant.sharepoint.com
  - SPO_CLIENT_ID=your-app-registration-guid
```

The device code appears in the container terminal. Open `https://microsoft.com/devicelogin`, enter the code, and the web server starts already connected.

**Connect via the UI** — click "Connect to SharePoint" in the browser. The device code appears in the container terminal (`podman logs <container>`). Authenticate at the device login URL; the UI updates when complete.

---

## Features

### Permission Analysis
- Site-level permissions with inheritance detection
- User enumeration with internal/external classification
- Group analysis with member counts
- Role assignment mapping — who has what permission on what resource
- Permission inheritance tree — detects broken inheritance at site, list, and library levels
- Sharing link audit — anonymous, company-wide, and specific-people links

### Deep Dive Views
- **Sites** — storage analysis, health scoring, hub site tracking, filterable grid
- **Users** — permission breakdown, internal vs external, security risk assessment
- **Groups** — membership analysis, empty group detection, size distribution
- **External Users** — domain analysis, access level audit, security findings
- **Role Assignments** — principal-to-permission mapping with chart and security review
- **Inheritance** — broken inheritance overview with percentage scoring
- **Sharing Links** — link type distribution, anonymous edit detection, exportable findings

### Export
- CSV export from every deep dive view
- Per-type data export (sites, users, groups, role assignments, inheritance, sharing links)

---

## Screenshots

<div align="center">
<p>
  <img src="Images/main.png" width="45%" />
  <img src="Images/Dashboard.png" width="45%" />
</p>
<p>
  <img src="Images/Export.png" width="45%" />
  <img src="Images/analyze.png" width="45%" />
</p>
<em>Desktop (WPF) interface shown. The web UI provides the same features in a browser.</em>
</div>

---

## App Registration

You need an Azure AD App Registration to connect to SharePoint Online.

### 1. Create the registration

1. **Azure Portal** > **App registrations** > **New registration**
2. Name: `SharePoint Permissions Analyzer`
3. Account types: **Single tenant**
4. Redirect URI: **Public client/native** > `http://localhost`

### 2. Add API permissions

| API | Permission | Type |
|-----|-----------|------|
| Microsoft Graph | Sites.FullControl.All | Delegated |
| Microsoft Graph | User.Read.All | Delegated |
| Microsoft Graph | GroupMember.Read.All | Delegated |
| SharePoint | AllSites.FullControl | Delegated |

> **Why FullControl?** This is a read-only tool, but SharePoint treats reading RoleAssignments and RoleDefinitionBindings as a privileged operation. `Sites.Read.All` / `AllSites.Read` is insufficient.
>
> The signed-in user must also be a **SharePoint Administrator** for tenant-wide site enumeration to work.

Click **Grant admin consent** after adding all permissions.

### 3. Enable public client

**Authentication** > **Allow public client flows** > **Yes** > **Save**

---

## Project Structure

```
SPO-Permissions-Analyzer/
├── Start-SPOTool.ps1              # WPF desktop entry point
├── Start-SPOTool-Web.ps1          # Web UI entry point
├── Dockerfile                     # Container image
├── compose.yaml                   # Podman/Docker compose
├── docker-entrypoint.ps1          # Container entrypoint
├── Install-Prerequisites.ps1      # Module installer
├── Functions/
│   ├── Core/                      # Shared by both UIs
│   │   ├── Settings.ps1
│   │   ├── SharePointDataManager.ps1
│   │   └── Logging.ps1
│   ├── SharePoint/                # SharePoint operations (shared)
│   │   └── SPOConnection.ps1
│   ├── Server/                    # Web UI backend
│   │   ├── WebServer.ps1
│   │   └── ApiHandlers.ps1
│   └── UI/                        # WPF interface
│       ├── MainWindow.ps1
│       ├── ConnectionTab.ps1
│       ├── OperationsTab.ps1      # Data collection (shared by both UIs)
│       ├── VisualAnalyticsTab.ps1
│       └── DeepDive/              # 7 deep dive windows
├── Views/                         # WPF XAML definitions
│   ├── Windows/MainWindow.xaml
│   └── DeepDive/                  # 7 deep dive XAMLs
├── Web/                           # Web UI frontend
│   ├── index.html
│   ├── css/app.css
│   └── js/ (api.js, charts.js, app.js)
├── Logs/                          # Auto-created
└── Reports/Generated/             # Auto-created
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Access is denied" | Verify app registration permissions and admin consent |
| "PnP PowerShell module not found" | Run `Install-Prerequisites.ps1` or `Install-Module PnP.PowerShell -Force` |
| "Connection timeout" | Check network; ensure redirect URI is `http://localhost` |
| Execution policy error | `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser` |
| Container auth not working | Check terminal for device code; ensure `SPO_HEADLESS=true` is set |

Check `./Logs/` for detailed error information. Try **Demo Mode** to isolate connection issues.

---

## Contributing

Contributions welcome! Fork the repo, create a feature branch, and open a pull request.

---

## License

MIT — see [LICENSE](LICENSE).

---

<div align="center">

**[Report an Issue](https://github.com/CreativeAcer/SPO-Permissions-Analyzer/issues)** | **[Discussions](https://github.com/CreativeAcer/SPO-Permissions-Analyzer/discussions)**

Made with care by [CreativeAcer](https://github.com/CreativeAcer)

</div>
