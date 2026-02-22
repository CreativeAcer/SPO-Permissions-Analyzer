# ============================================
# WebServer.ps1 - Lightweight HTTP Server
# ============================================
# Replaces WPF UI with a local web server + browser frontend.
# Uses System.Net.HttpListener (built-in, no external deps).

function Start-WebServer {
    <#
    .SYNOPSIS
    Starts a local HTTP server and opens the browser
    .PARAMETER Port
    Port number to listen on (default 8080)
    .PARAMETER ListenAddress
    Address to bind to. Use 'localhost' for local-only, '+' or '*' for all interfaces (required in containers)
    .PARAMETER NoBrowser
    Skip opening the default browser (used in container/headless mode)
    #>
    param(
        [int]$Port = 8080,
        [string]$ListenAddress = "localhost",
        [switch]$NoBrowser
    )

    $webRoot = Join-Path $PSScriptRoot "..\..\Web"
    $webRoot = [System.IO.Path]::GetFullPath($webRoot)

    if (-not (Test-Path $webRoot)) {
        throw "Web root not found at: $webRoot"
    }

    $prefix = "http://${ListenAddress}:$Port/"
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($prefix)

    # Shared operation state - synchronized for thread-safe access from background runspaces
    $script:ServerState = [hashtable]::Synchronized(@{
        Listener         = $listener
        WebRoot          = $webRoot
        Port             = $Port
        Running          = $true
        OperationLog     = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        OperationRunning = $false
        OperationComplete = $true
        OperationError   = $null
        BackgroundJob    = $null
        SharePointData   = $script:SharePointData
    })

    try {
        $listener.Start()
        Write-Host ""
        Write-Host "  =======================================" -ForegroundColor Cyan
        Write-Host "  PermiX (Web)" -ForegroundColor Cyan
        Write-Host "  =======================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Server running at: " -NoNewline
        Write-Host "http://localhost:$Port/" -ForegroundColor Green
        if ($ListenAddress -ne "localhost") {
            Write-Host "  Listening on: $prefix"
        }
        Write-Host "  Web root: $webRoot"
        Write-Host ""
        Write-Host "  Press Ctrl+C to stop the server" -ForegroundColor Yellow
        Write-Host ""

        # Open default browser (skip in container/headless mode)
        if (-not $NoBrowser) {
            try {
                Start-Process "http://localhost:$Port/"
            }
            catch {
                Write-Host "  Could not open browser automatically." -ForegroundColor Yellow
                Write-Host "  Please navigate to http://localhost:$Port/ manually." -ForegroundColor Yellow
            }
        }

        Write-ActivityLog "Web server started on port $Port" -Level "Information"

        # Main request loop - uses async GetContext so we can service requests
        # while background operations (permissions analysis, site fetching) are running
        while ($script:ServerState.Running -and $listener.IsListening) {
            try {
                $asyncResult = $listener.BeginGetContext($null, $null)
                # Wait up to 500ms for a request, then loop to check Running flag
                while (-not $asyncResult.AsyncWaitHandle.WaitOne(500)) {
                    if (-not $script:ServerState.Running -or -not $listener.IsListening) { break }
                }
                if ($asyncResult.IsCompleted -and $script:ServerState.Running -and $listener.IsListening) {
                    $context = $listener.EndGetContext($asyncResult)
                    Invoke-RequestHandler -Context $context
                }
            }
            catch [System.Net.HttpListenerException] {
                if ($script:ServerState.Running) {
                    Write-ActivityLog "Listener error: $($_.Exception.Message)" -Level "Warning"
                }
            }
            catch {
                Write-ActivityLog "Request error: $($_.Exception.Message)" -Level "Error"
            }
        }
    }
    finally {
        $listener.Stop()
        $listener.Close()
        Write-Host ""
        Write-Host "  Server stopped." -ForegroundColor Yellow
        Write-ActivityLog "Web server stopped" -Level "Information"
    }
}

function Invoke-RequestHandler {
    <#
    .SYNOPSIS
    Routes an incoming HTTP request to the appropriate handler
    #>
    param(
        [System.Net.HttpListenerContext]$Context
    )

    $request  = $Context.Request
    $response = $Context.Response
    $path     = $request.Url.LocalPath

    try {
        # API routes
        if ($path.StartsWith("/api/")) {
            Invoke-ApiHandler -Request $request -Response $response -Path $path
        }
        # Static files
        else {
            Send-StaticFile -Path $path -Response $response
        }
    }
    catch {
        Send-JsonResponse -Response $response -Data @{
            error = $true
            message = $_.Exception.Message
        } -StatusCode 500
    }
}

function Send-StaticFile {
    <#
    .SYNOPSIS
    Serves a static file from the Web/ directory
    #>
    param(
        [string]$Path,
        [System.Net.HttpListenerResponse]$Response
    )

    # Default to index.html
    if ($Path -eq "/" -or $Path -eq "") { $Path = "/index.html" }

    $filePath = Join-Path $script:ServerState.WebRoot $Path.TrimStart("/")
    $filePath = [System.IO.Path]::GetFullPath($filePath)

    # Security: ensure file is within web root
    if (-not $filePath.StartsWith($script:ServerState.WebRoot)) {
        Send-JsonResponse -Response $Response -Data @{ error = "Forbidden" } -StatusCode 403
        return
    }

    if (-not (Test-Path $filePath -PathType Leaf)) {
        Send-JsonResponse -Response $Response -Data @{ error = "Not found: $Path" } -StatusCode 404
        return
    }

    # MIME types
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    $mimeTypes = @{
        ".html" = "text/html; charset=utf-8"
        ".css"  = "text/css; charset=utf-8"
        ".js"   = "application/javascript; charset=utf-8"
        ".json" = "application/json; charset=utf-8"
        ".png"  = "image/png"
        ".svg"  = "image/svg+xml"
        ".ico"  = "image/x-icon"
    }

    $contentType = if ($mimeTypes.ContainsKey($extension)) { $mimeTypes[$extension] } else { "application/octet-stream" }
    $Response.ContentType = $contentType

    # Cache static assets (CSS/JS) for 1 hour, HTML never
    if ($extension -in @(".css", ".js", ".png", ".svg")) {
        $Response.Headers.Add("Cache-Control", "public, max-age=3600")
    } else {
        $Response.Headers.Add("Cache-Control", "no-cache")
    }

    $content = [System.IO.File]::ReadAllBytes($filePath)
    $Response.ContentLength64 = $content.Length
    $Response.OutputStream.Write($content, 0, $content.Length)
    $Response.OutputStream.Close()
}

function Send-JsonResponse {
    <#
    .SYNOPSIS
    Sends a JSON response
    #>
    param(
        [System.Net.HttpListenerResponse]$Response,
        [object]$Data,
        [int]$StatusCode = 200
    )

    $Response.StatusCode = $StatusCode
    $Response.ContentType = "application/json; charset=utf-8"
    $Response.Headers.Add("Cache-Control", "no-cache")

    $json = $Data | ConvertTo-Json -Depth 10 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Read-RequestBody {
    <#
    .SYNOPSIS
    Reads and parses JSON request body
    #>
    param(
        [System.Net.HttpListenerRequest]$Request
    )

    if ($Request.HasEntityBody) {
        $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
        $body = $reader.ReadToEnd()
        $reader.Close()

        if ($body -and $body.Length -gt 0) {
            return ($body | ConvertFrom-Json)
        }
    }
    return $null
}

function Stop-WebServer {
    <#
    .SYNOPSIS
    Gracefully stops the web server
    #>
    $script:ServerState.Running = $false
    if ($script:ServerState.Listener -and $script:ServerState.Listener.IsListening) {
        $script:ServerState.Listener.Stop()
    }
}
