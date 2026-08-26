[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$vendor = Join-Path $root 'vendor'
$fontDir = Join-Path $vendor 'fontawesome'

New-Item -ItemType Directory -Force -Path $vendor, $fontDir | Out-Null

Copy-Item (Join-Path $root 'node_modules\docx\build\index.umd.js') (Join-Path $vendor 'docx.umd.js') -Force
Copy-Item (Join-Path $root 'node_modules\file-saver\dist\FileSaver.min.js') (Join-Path $vendor 'FileSaver.min.js') -Force
Copy-Item (Join-Path $root 'node_modules\xlsx\dist\xlsx.full.min.js') (Join-Path $vendor 'xlsx.full.min.js') -Force
Copy-Item (Join-Path $root 'node_modules\jszip\dist\jszip.min.js') (Join-Path $vendor 'jszip.min.js') -Force
$destCss = Join-Path $fontDir 'css'
$destWebfonts = Join-Path $fontDir 'webfonts'
if (Test-Path $destCss) { Remove-Item $destCss -Recurse -Force }
if (Test-Path $destWebfonts) { Remove-Item $destWebfonts -Recurse -Force }
Copy-Item (Join-Path $root 'node_modules\@fortawesome\fontawesome-free\css') $destCss -Recurse -Force
Copy-Item (Join-Path $root 'node_modules\@fortawesome\fontawesome-free\webfonts') $destWebfonts -Recurse -Force

& (Join-Path $root 'node_modules\.bin\tailwindcss.cmd') -i (Join-Path $PSScriptRoot 'tailwind-input.css') -o (Join-Path $vendor 'tailwind.css') --content (Join-Path $root 'index.html') --minify
if ($LASTEXITCODE -ne 0) { throw 'Tailwind CSS build failed.' }

Write-Output 'Local frontend assets updated.'
