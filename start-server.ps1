param(
    [int]$Port = 8080,
    [string]$Root = (Get-Location).Path,
    [switch]$LaunchBrowser
)

Set-Location -Path $Root
Write-Host "Serving files from: $Root"
Write-Host "Listening on http://localhost:$Port/"

if (Get-Command npx -ErrorAction SilentlyContinue) {
    Write-Host "Using: npx http-server"
    & npx http-server -c-1 -p $Port
    exit $LASTEXITCODE
}

if (Get-Command python -ErrorAction SilentlyContinue) {
    Write-Host "Using: python -m http.server"
    & python -m http.server $Port
    exit $LASTEXITCODE
}

Write-Host "No npx/python found — using PowerShell HttpListener fallback. Press Ctrl+C to stop."
if ($LaunchBrowser) { Start-Process "http://localhost:$Port" }

function Get-MimeType($path) {
    switch ([System.IO.Path]::GetExtension($path).ToLower()) {
        '.html' { 'text/html' }
        '.htm'  { 'text/html' }
        '.js'   { 'application/javascript' }
        '.css'  { 'text/css' }
        '.json' { 'application/json' }
        '.png'  { 'image/png' }
        '.jpg'  { 'image/jpeg' }
        '.jpeg' { 'image/jpeg' }
        '.gif'  { 'image/gif' }
        '.svg'  { 'image/svg+xml' }
        '.txt'  { 'text/plain' }
        default { 'application/octet-stream' }
    }
}

$listener = New-Object System.Net.HttpListener
$prefix = "http://*:$Port/"
$listener.Prefixes.Add($prefix)
$listener.Start()

try {
    while ($true) {
        $context = $listener.GetContext()
        $req = $context.Request
        $resp = $context.Response

        $rawPath = $req.Url.LocalPath.TrimStart('/')
        if ([string]::IsNullOrWhiteSpace($rawPath)) { $rawPath = 'index.html' }
        $filePath = Join-Path -Path $Root -ChildPath $rawPath

        if (Test-Path $filePath -PathType Leaf) {
            try {
                $bytes = [System.IO.File]::ReadAllBytes($filePath)
                $resp.ContentLength64 = $bytes.Length
                $resp.ContentType = Get-MimeType $filePath
                $resp.OutputStream.Write($bytes, 0, $bytes.Length)
            } catch {
                $resp.StatusCode = 500
                $err = [System.Text.Encoding]::UTF8.GetBytes("500 Internal Server Error")
                $resp.ContentLength64 = $err.Length
                $resp.OutputStream.Write($err,0,$err.Length)
            }
        } else {
            $resp.StatusCode = 404
            $not = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
            $resp.ContentLength64 = $not.Length
            $resp.OutputStream.Write($not,0,$not.Length)
        }
        $resp.OutputStream.Close()
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
