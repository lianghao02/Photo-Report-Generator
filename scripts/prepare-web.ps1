[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
& (Join-Path $PSScriptRoot 'prepare-assets.ps1')

$web = Join-Path $root 'web'
New-Item -ItemType Directory -Force -Path $web | Out-Null
Copy-Item (Join-Path $root 'index.html') (Join-Path $web 'index.html') -Force
Copy-Item (Join-Path $root 'vendor') (Join-Path $web 'vendor') -Recurse -Force
