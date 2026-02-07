<#
.SYNOPSIS
    Container entrypoint - dispatches to web or local (WPF) mode
.DESCRIPTION
    Reads UI_MODE environment variable:
      - "web"   (default) → launches browser-based UI on port 8080
      - "local" → launches WPF desktop UI (requires Windows container with display)
#>

$mode = $env:UI_MODE
if (-not $mode) { $mode = "web" }

switch ($mode.ToLower()) {
    "web" {
        Write-Host ""
        Write-Host "  Starting in WEB mode (container-friendly)" -ForegroundColor Cyan
        Write-Host "  Access the UI at: http://localhost:8080" -ForegroundColor Green
        Write-Host ""
        & /app/Start-SPOTool-Web.ps1 -Port 8080 -ListenAddress "+" -NoBrowser
    }
    "local" {
        # WPF requires Windows desktop assemblies + a display
        $hasWpf = $false
        try {
            Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
            $hasWpf = $true
        }
        catch {
            # Not available on this platform
        }

        if ($hasWpf) {
            Write-Host ""
            Write-Host "  Starting in LOCAL mode (WPF desktop)" -ForegroundColor Cyan
            Write-Host ""
            & /app/Start-SPOTool.ps1
        }
        else {
            Write-Host ""
            Write-Host "  ERROR: LOCAL mode (WPF/XAML) is not available in this container." -ForegroundColor Red
            Write-Host ""
            Write-Host "  WPF requires a Windows container with display support." -ForegroundColor Yellow
            Write-Host "  Options:" -ForegroundColor Yellow
            Write-Host "    1. Run directly on your Windows host:  .\Start-SPOTool.ps1" -ForegroundColor White
            Write-Host "    2. Use web mode instead:  podman compose up" -ForegroundColor White
            Write-Host ""
            exit 1
        }
    }
    default {
        Write-Host "  Unknown UI_MODE: '$mode'. Use 'web' (default) or 'local'." -ForegroundColor Red
        exit 1
    }
}
