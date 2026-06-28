$ErrorActionPreference = "Stop"
$AppName = -join ([char[]](22270, 20070, 30005, 21830, 32447, 19978, 27963, 21160, 20215, 26684, 33258, 21160, 29983, 25104, 22120))

python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --onefile `
  --icon "assets\app.ico" `
  --add-data "assets\app.ico;." `
  --name $AppName `
  "main.py"

Write-Host "Generated app exe."
