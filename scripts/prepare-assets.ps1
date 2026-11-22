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
    # 掃描範圍必須包含由 index.html 拆出的 css/app.css 與 js/app.js（主程式字串裡有大量動態類別名稱），
    # 否則重建後會少掉規則、動態產生的畫面元素失去樣式。所有含類別的檔案合計產出需涵蓋完整規則。
    $tailwindContent = (Join-Path $root 'index.html') + ',' + (Join-Path $root 'css\app.css') + ',' + (Join-Path $root 'js\app.js') + ',' + (Join-Path $root 'js\ui\audit-ui.js') + ',' + (Join-Path $root 'js\ui\merge-report-ui.js')
    & $tailwind -i (Join-Path $PSScriptRoot 'tailwind-input.css') -o (Join-Path $vendor 'tailwind.css') --content $tailwindContent --minify
    if ($LASTEXITCODE -ne 0) { throw 'Tailwind CSS build failed.' }
} elseif (Test-Path (Join-Path $vendor 'tailwind.css')) {
    Write-Warning '找不到 Tailwind CSS 套件，保留既有離線樣式。'
} else {
    throw '找不到必要的 Tailwind CSS 資源。'
}

# 自動依靜態資源內容產生 8 碼雜湊並更新 index.html 的 ?v= 快取參數（冪等寫入）
$indexPath = Join-Path $root 'index.html'
if (Test-Path $indexPath) {
    $verTag = if (Test-Path (Join-Path $root 'version.txt')) {
        (Get-Content (Join-Path $root 'version.txt') -Raw -Encoding UTF8).Trim() -replace '[^A-Za-z0-9._-]', ''
    } else { 'v0' }
    $hashInputs = @()
    foreach ($candidate in @((Join-Path $root 'css\app.css'), (Join-Path $vendor 'tailwind.css'))) {
        if (Test-Path $candidate) { $hashInputs += Get-Item $candidate }
    }
    if (Test-Path (Join-Path $root 'js')) {
        $hashInputs += Get-ChildItem (Join-Path $root 'js') -Recurse -File -Filter '*.js' | Sort-Object FullName
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        foreach ($item in $hashInputs) {
            $relBytes = [System.Text.Encoding]::UTF8.GetBytes($item.FullName.Substring($root.Length) + "`n")
            [void]$sha.TransformBlock($relBytes, 0, $relBytes.Length, $relBytes, 0)
            $fileBytes = [System.IO.File]::ReadAllBytes($item.FullName)
            [void]$sha.TransformBlock($fileBytes, 0, $fileBytes.Length, $fileBytes, 0)
        }
        [void]$sha.TransformFinalBlock([byte[]]::new(0), 0, 0)
        $shortHash = ([System.BitConverter]::ToString($sha.Hash) -replace '-', '').Substring(0, 8).ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
    $cacheTag = "$verTag-$shortHash"
    $rawHtml = [System.IO.File]::ReadAllText($indexPath, [System.Text.UTF8Encoding]::new($false))
    $updatedHtml = [regex]::Replace(
        $rawHtml,
        '((?:vendor/tailwind\.css|css/app\.css|js/[^"'']+?\.js))\?v=[^"'' >]+',
        ('${1}?v=' + $cacheTag)
    )
    if ($updatedHtml -ne $rawHtml) {
        [System.IO.File]::WriteAllText($indexPath, $updatedHtml, [System.Text.UTF8Encoding]::new($false))
    }
}

Write-Output 'Local frontend assets updated.'
