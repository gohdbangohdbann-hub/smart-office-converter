$ErrorActionPreference = "Stop"
$InstallRoot = Join-Path $env:LOCALAPPDATA "NawaOCROffice"
$ManifestSource = Join-Path $PSScriptRoot "manifest.xml"
$ManifestTarget = Join-Path $InstallRoot "manifest.xml"
$ReadmeSource = Join-Path $PSScriptRoot "INSTALL-WINDOWS-AR.md"

New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
Copy-Item -Force $ManifestSource $ManifestTarget
Copy-Item -Force $ReadmeSource (Join-Path $InstallRoot "INSTALL-WINDOWS-AR.md")

$word = Get-Command "winword.exe" -ErrorAction SilentlyContinue
$excel = Get-Command "excel.exe" -ErrorAction SilentlyContinue

Write-Host "تم تجهيز Nawa OCR Office في: $InstallRoot" -ForegroundColor Green
Write-Host "الخطوة الأخيرة: افتح Word أو Excel ثم Insert > Get Add-ins > Manage My Add-ins > Upload My Add-in، واختر:" -ForegroundColor Yellow
Write-Host $ManifestTarget -ForegroundColor Cyan

if ($word) { Start-Process $word.Source }
elseif ($excel) { Start-Process $excel.Source }
else { Write-Host "لم يتم العثور على Word أو Excel تلقائيًا. افتحهما يدويًا." -ForegroundColor Yellow }
