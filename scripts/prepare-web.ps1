[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
& (Join-Path $PSScriptRoot 'prepare-assets.ps1')

$web = Join-Path $root 'web'
New-Item -ItemType Directory -Force -Path $web | Out-Null
Copy-Item (Join-Path $root 'index.html') (Join-Path $web 'index.html') -Force
$webVendor = Join-Path $web 'vendor'
New-Item -ItemType Directory -Force -Path $webVendor | Out-Null
Copy-Item (Join-Path $root 'vendor\*') $webVendor -Recurse -Force
$webJs = Join-Path $web 'js'
if (Test-Path (Join-Path $root 'js')) {
    New-Item -ItemType Directory -Force -Path $webJs | Out-Null
    Copy-Item (Join-Path $root 'js\*') $webJs -Recurse -Force
}
if (Test-Path (Join-Path $root 'version.txt')) {
    Copy-Item (Join-Path $root 'version.txt') (Join-Path $web 'version.txt') -Force
}
