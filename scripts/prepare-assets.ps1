[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$vendor = Join-Path $root 'vendor'
$fontDir = Join-Path $vendor 'fontawesome'

New-Item -ItemType Directory -Force -Path $vendor, $fontDir | Out-Null

$assets = @(
    @{ Source = 'node_modules\docx\build\index.umd.js'; Destination = 'docx.umd.js' },
    @{ Source = 'node_modules\file-saver\dist\FileSaver.min.js'; Destination = 'FileSaver.min.js' },
    @{ Source = 'node_modules\xlsx\dist\xlsx.full.min.js'; Destination = 'xlsx.full.min.js' },
    @{ Source = 'node_modules\jszip\dist\jszip.min.js'; Destination = 'jszip.min.js' },
    @{ Source = 'node_modules\jspdf\dist\jspdf.umd.min.js'; Destination = 'jspdf.umd.min.js' }
)
foreach ($asset in $assets) {
    $sourcePath = Join-Path $root $asset.Source
    $destinationPath = Join-Path $vendor $asset.Destination
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath $destinationPath -Force
    } elseif (Test-Path $destinationPath) {
        Write-Warning "找不到 $($asset.Source)，保留既有離線資源：$($asset.Destination)"
    } else {
        throw "找不到必要資源：$($asset.Source)"
    }
}

$destCss = Join-Path $fontDir 'css'
$destWebfonts = Join-Path $fontDir 'webfonts'
$sourceFontDir = Join-Path $root 'node_modules\@fortawesome\fontawesome-free'
if (Test-Path $sourceFontDir) {
    if (Test-Path $destCss) { Remove-Item $destCss -Recurse -Force }
    if (Test-Path $destWebfonts) { Remove-Item $destWebfonts -Recurse -Force }
    Copy-Item (Join-Path $sourceFontDir 'css') $destCss -Recurse -Force
    Copy-Item (Join-Path $sourceFontDir 'webfonts') $destWebfonts -Recurse -Force
} elseif ((Test-Path $destCss) -and (Test-Path $destWebfonts)) {
    Write-Warning '找不到 Font Awesome 套件，保留既有離線字型資源。'
} else {
    throw '找不到必要的 Font Awesome 資源。'
}

$tailwind = Join-Path $root 'node_modules\.bin\tailwindcss.cmd'
if (Test-Path $tailwind) {
    & $tailwind -i (Join-Path $PSScriptRoot 'tailwind-input.css') -o (Join-Path $vendor 'tailwind.css') --content (Join-Path $root 'index.html') --minify
    if ($LASTEXITCODE -ne 0) { throw 'Tailwind CSS build failed.' }
} elseif (Test-Path (Join-Path $vendor 'tailwind.css')) {
    Write-Warning '找不到 Tailwind CSS 套件，保留既有離線樣式。'
} else {
    throw '找不到必要的 Tailwind CSS 資源。'
}

Write-Output 'Local frontend assets updated.'
