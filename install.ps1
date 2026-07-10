# officecli installer for Windows (PowerShell 5.1+ / pwsh).
#
#   irm https://github.com/pariharshyamu/office-agent/releases/latest/download/install.ps1 | iex
#
# Options (environment variables):
#   OFFICECLI_VERSION   install a specific tag (e.g. v0.2.0); default: latest
#   OFFICECLI_INSTALL   install directory; default: %LOCALAPPDATA%\Programs\officecli
$ErrorActionPreference = 'Stop'

$repo = 'pariharshyamu/office-agent'
$version = if ($env:OFFICECLI_VERSION) { $env:OFFICECLI_VERSION } else { 'latest' }
$installDir = if ($env:OFFICECLI_INSTALL) { $env:OFFICECLI_INSTALL } else { Join-Path $env:LOCALAPPDATA 'Programs\officecli' }

$arch = $env:PROCESSOR_ARCHITECTURE
if ($env:PROCESSOR_ARCHITEW6432) { $arch = $env:PROCESSOR_ARCHITEW6432 }
switch ($arch) {
    'AMD64' { $asset = 'officecli-win-x64.exe' }
    'ARM64' { $asset = 'officecli-win-arm64.exe' }
    default { throw "Unsupported Windows architecture '$arch'" }
}

if ($version -eq 'latest') {
    $url = "https://github.com/$repo/releases/latest/download/$asset"
} else {
    $url = "https://github.com/$repo/releases/download/$version/$asset"
}

Write-Host "Downloading $asset ($version) ..."
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
$exe = Join-Path $installDir 'officecli.exe'
Invoke-WebRequest -Uri $url -OutFile $exe -UseBasicParsing

$ver = & $exe --version
Write-Host "Installed $ver to $exe"

# Add the install directory to the user PATH if it is not there yet.
$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
if (-not (($userPath -split ';') -contains $installDir)) {
    [Environment]::SetEnvironmentVariable('Path', "$userPath;$installDir", 'User')
    Write-Host "Added $installDir to your user PATH (restart your terminal to pick it up)."
}

Write-Host ''
Write-Host 'Get started:   officecli help'
Write-Host "MCP server:    claude mcp add officecli -- $exe mcp"
