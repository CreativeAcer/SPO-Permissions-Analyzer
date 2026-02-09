# SharePoint Online Permissions Analyzer  [![CodeFactor](https://www.codefactor.io/repository/github/creativeacer/spo-permissions-analyzer/badge)](https://www.codefactor.io/repository/github/creativeacer/spo-permissions-analyzer)

<div align="center">

![SharePoint](https://img.shields.io/badge/SharePoint-Online-0078D4?style=for-the-badge&logo=microsoft-sharepoint&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-7.0+-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![PnP](https://img.shields.io/badge/PnP-PowerShell_3.x-orange?style=for-the-badge)
![Container](https://img.shields.io/badge/Container-Podman%20%7C%20Docker-blue?style=for-the-badge&logo=podman&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

SharePoint Online security and permissions analyzer with **risk assessment**, **external user enrichment**, and **visual analytics**.

Identifies security risks, stale accounts, anonymous sharing, and permission issues across your SharePoint environment.

Runs as a **container** or **local web server** with a modern browser-based interface.

> **Note**: As of 08/02/2026, the XAML/WPF desktop version has been removed to reduce complexity. The web interface provides all features with better cross-platform support.

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

### Option B: Local Web Server

Requires PowerShell 7+ and PnP.PowerShell.

```powershell
git clone https://github.com/CreativeAcer/SPO-Permissions-Analyzer.git
cd SPO-Permissions-Analyzer
.\Install-Prerequisites.ps1
.\Start-SPOTool-Web.ps1    # opens http://localhost:8080
```

### Demo Mode

Both options support **Demo Mode** — click the button on the Connection tab to explore all features with sample data, no SharePoint connection required.

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
- Role assignment mapping across all resources
- Permission inheritance tree with broken inheritance detection
- Sharing link audit (anonymous, company-wide, specific-people)

### Risk Assessment
- **11 security rules** across Critical, High, Medium, and Low severity levels
- **Overall risk score** (0-100) with color-coded dashboard banner
- **Identifies**: external admins, anonymous edit links, excessive permissions, broken inheritance, stale accounts, empty groups
- **Filterable findings** by severity with detailed remediation guidance

### External User Enrichment
- **Microsoft Graph integration** — enriches external users with live account data
- **Stale account detection** — flags accounts inactive for 90+ days
- **Account status tracking** — Active, Disabled, or Never signed in
- **Domain analysis** — aggregates external users by organization

### Deep Dive Views
- **Sites** — storage analysis, health scoring, filterable grid
- **Users** — permission breakdown, internal vs external classification
- **Groups** — membership analysis, empty group detection
- **External Users** — domain analysis, access audit, enrichment with stale warnings
- **Role Assignments** — principal-to-permission mapping with security review
- **Inheritance** — broken inheritance overview with percentage scoring, **interactive tree view** showing parent-child hierarchy
- **Sharing Links** — link type distribution, anonymous edit detection

### Interactive UI Features
- **Global search** — Omnibox search across sites, users, groups with keyboard shortcuts (Ctrl+K / Cmd+K)
- **Clickable charts** — Click any chart bar or segment to drill down into detailed data
- **Tree visualizations** — Collapsible hierarchical view of permission inheritance
- **Format selection** — Export any data as CSV or JSON with format chooser modal
- Sortable/filterable tables with real-time search
- Responsive design for desktop and mobile

---

## Security Rules

11 rules evaluate your environment across External Access, Sharing Links, Permissions, Inheritance, and Groups.

**Critical**: External site admins, anonymous edit links
**High**: External users with elevated permissions, anonymous links, excessive Full Control, broken inheritance
**Medium**: Multiple external domains, excessive org-wide links, direct user assignments
**Low**: Empty groups

Risk score (0-100) calculated from top 5 findings. Levels: Critical (80+), High (60-79), Medium (30-59), Low (1-29), None (0).

---

## Screenshots

<div align="center">
<p>
  <img src="Images/Visual-Analytics-web.png" width="90%" />
</p>
<p>
  <img src="Images/Risk-assessment-web.png" width="45%" />
  <img src="Images/Enrich-Users-Web.png" width="45%" />
</p>
<em>Modern web interface with risk assessment, visual analytics, and external user enrichment.</em>
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
├── Start-SPOTool-Web.ps1          # Web UI entry point
├── Dockerfile                     # Container image
├── compose.yaml                   # Podman/Docker compose
├── docker-entrypoint.ps1          # Container entrypoint
├── Install-Prerequisites.ps1      # Module installer
├── Functions/
│   ├── Core/                      # Core business logic
│   │   ├── Settings.ps1
│   │   ├── SharePointDataManager.ps1
│   │   ├── RiskScoring.ps1
│   │   ├── GraphEnrichment.ps1
│   │   └── Logging.ps1
│   ├── SharePoint/                # SharePoint operations
│   │   └── SPOConnection.ps1
│   └── Server/                    # Web backend
│       ├── WebServer.ps1
│       └── ApiHandlers.ps1
├── Web/                           # Web frontend
│   ├── index.html
│   ├── css/app.css
│   └── js/
│       ├── api.js
│       ├── app.js
│       ├── charts.js
│       └── ui-helpers.js
├── Images/                        # Screenshots
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
