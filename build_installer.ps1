$ErrorActionPreference = "Stop"
$AppName = -join ([char[]](22270, 20070, 30005, 21830, 32447, 19978, 27963, 21160, 20215, 26684, 33258, 21160, 29983, 25104, 22120))
$AppExe = Join-Path "dist" ($AppName + ".exe")

if (-not (Test-Path $AppExe)) {
  .\build_exe.ps1
}

$iscc = Get-Command iscc -ErrorAction SilentlyContinue
if (-not $iscc) {
  $defaultPath = Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"
  if (Test-Path $defaultPath) {
    $iscc = Get-Item $defaultPath
  }
}

if (-not $iscc) {
  $userPath = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
  if (Test-Path $userPath) {
    $iscc = Get-Item $userPath
  }
}

if (-not $iscc) {
  throw "ISCC.exe was not found. Install Inno Setup first."
}

$isccPath = if ($iscc.Source) { $iscc.Source } else { $iscc.FullName }
& $isccPath "installer.iss"

Write-Host "Installer generated in installer folder."
