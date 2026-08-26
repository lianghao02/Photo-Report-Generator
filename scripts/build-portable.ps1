# build-portable.ps1
# Usage: .\scripts\build-portable.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$env:PATH = "$env:USERPROFILE\.cargo\bin;$env:USERPROFILE\scoop\apps\mingw\current\bin;$env:USERPROFILE\scoop\shims;" + $env:PATH

$Root        = Split-Path $PSScriptRoot -Parent
$Target      = "x86_64-pc-windows-gnu"
$ReleaseDir  = Join-Path $Root "src-tauri\target\$Target\release"
$BundleDir   = Join-Path $ReleaseDir "bundle\nsis"
$RawExe      = Join-Path $ReleaseDir "photo-report-generator.exe"
$WebView2Dll = Join-Path $ReleaseDir "WebView2Loader.dll"

Set-Location $Root

Write-Host ">> Step 1: sync web assets" -ForegroundColor Cyan
& (Join-Path $PSScriptRoot "prepare-web.ps1")

Write-Host ">> Step 2: tauri build" -ForegroundColor Cyan
$npxCmd = Get-Command npx.cmd, npx -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
if (-not $npxCmd) { throw "找不到 npx 命令，請確認 Node.js 已正確安裝。" }
& $npxCmd tauri build --target $Target
if ($LASTEXITCODE -ne 0) {
    throw "Tauri 建置失敗（結束碼：$LASTEXITCODE）。"
}
if (-not (Test-Path $RawExe)) {
    Write-Host "FAILED: exe not found" -ForegroundColor Red
    exit 1
}

$conf    = Get-Content (Join-Path $Root "src-tauri\tauri.conf.json") -Encoding UTF8 | ConvertFrom-Json
$ver     = $conf.version
$appName = $conf.productName

Write-Host ">> Step 3: build portable folder" -ForegroundColor Cyan
$folderName  = $appName + "_" + $ver + "_x64_portable"
$portableDir = Join-Path $BundleDir $folderName
$portableZip = Join-Path $BundleDir ($folderName + ".zip")

if (Test-Path $portableDir) { Remove-Item $portableDir -Recurse -Force }
New-Item -ItemType Directory -Path $portableDir | Out-Null

Copy-Item $RawExe (Join-Path $portableDir ($appName + ".exe")) -Force

if (Test-Path $WebView2Dll) {
    Copy-Item $WebView2Dll (Join-Path $portableDir "WebView2Loader.dll") -Force
    Write-Host "   WebView2Loader.dll included" -ForegroundColor Gray
}

$readmePath = Join-Path $portableDir "README.txt"
$lines = @(
    "Portable - v" + $ver,
    "================================",
    "Double-click " + $appName + ".exe to launch.",
    "Keep all files in this folder together.",
    "Requires Windows 10/11 with WebView2 runtime."
)
Set-Content -Path $readmePath -Value $lines -Encoding UTF8

Write-Host ">> Step 4: compress to ZIP" -ForegroundColor Cyan
if (Test-Path $portableZip) { Remove-Item $portableZip -Force }
Compress-Archive -Path $portableDir -DestinationPath $portableZip -CompressionLevel Optimal

$setupExe = Join-Path $BundleDir ($appName + "_" + $ver + "_x64-setup.exe")
Write-Host ""
Write-Host "Build Complete!" -ForegroundColor Green

if (Test-Path $setupExe) {
    $mb = [math]::Round((Get-Item $setupExe).Length / 1MB, 2)
    Write-Host "  [Installer] $((Get-Item $setupExe).Name)  ($mb MB)"
}
if (Test-Path $portableZip) {
    $mb = [math]::Round((Get-Item $portableZip).Length / 1MB, 2)
    Write-Host "  [Portable]  $((Get-Item $portableZip).Name)  ($mb MB)"
}
Write-Host "  Output: $BundleDir"
