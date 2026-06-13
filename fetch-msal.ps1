<#
fetch-msal.ps1
Installs @azure/msal-browser via npm and copies the minified bundle to ./libs/msal-browser.min.js

Usage:
  .\fetch-msal.ps1
  or
  .\fetch-msal.ps1 -Version 2.46.1
#>

param(
  [string]$Version = "latest"
)

$destDir = Join-Path -Path (Get-Location).Path -ChildPath 'libs'
if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }

function Try-Download($url, $outFile) {
  try {
    Write-Host "Attempting download: $url"
    Invoke-WebRequest -Uri $url -OutFile $outFile -UseBasicParsing -ErrorAction Stop
    Write-Host "Downloaded: $outFile"
    return $true
  } catch {
    Write-Warning "Download failed: $url -> $_"
    return $false
  }
}

# Try direct CDN downloads first (prefer jsDelivr + latest tag), then fall back to npm install if available
$localOut = Join-Path $destDir 'msal-browser.min.js'

$cdnUrls = @()
if ($Version -and $Version -ne 'latest') {
  $cdnUrls += "https://cdn.jsdelivr.net/npm/@azure/msal-browser@$Version/dist/msal-browser.min.js"
  $cdnUrls += "https://unpkg.com/@azure/msal-browser@$Version/dist/msal-browser.min.js"
}
# always try latest tags
$cdnUrls += "https://cdn.jsdelivr.net/npm/@azure/msal-browser@latest/dist/msal-browser.min.js"
$cdnUrls += "https://unpkg.com/@azure/msal-browser/dist/msal-browser.min.js"

$downloaded = $false
foreach ($u in $cdnUrls) {
  if (Try-Download $u $localOut) { $downloaded = $true; break }
}

if (-not $downloaded) {
  if (Get-Command npm -ErrorAction SilentlyContinue) {
    Write-Host "CDN download failed; attempting to install via npm and copy bundle..."
    try {
      if ($Version -eq 'latest') { npm install @azure/msal-browser --no-audit --no-fund --silent } else { npm install @azure/msal-browser@$Version --no-audit --no-fund --silent }
    } catch {
      Write-Error "npm install failed: $_"
      exit 1
    }
    $nmRoot = Join-Path (Get-Location).Path 'node_modules\@azure\msal-browser'
    if (-not (Test-Path $nmRoot)) { Write-Error "Installed package not found at $nmRoot"; exit 1 }

    # Find any browser bundle (prefer *.min.js)
    $candidates = Get-ChildItem -Path $nmRoot -Recurse -Include '*msal-browser*.min.js','*msal-browser*.js' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '*.map' }
    if ($null -eq $candidates -or $candidates.Count -eq 0) {
      Write-Error "Could not find any msal browser bundle under $nmRoot"
      exit 1
    }

    # prefer minified files
    $ordered = $candidates | Sort-Object -Property @{Expression={$_.Name -like '*.min.js'};Descending=$true}
    $picked = $ordered[0]
    Copy-Item -Path $picked.FullName -Destination $localOut -Force
    Write-Host "Copied $($picked.FullName) -> $localOut"
    exit 0
  } else {
    Write-Error "Failed to download msal-browser from CDNs and npm is not available. Install Node.js/npm or provide the file at $localOut"
    exit 1
  }
}

Write-Host "msal-browser saved to: $localOut"
Write-Host "You can now open index.html which will prefer the local copy under ./libs/"
