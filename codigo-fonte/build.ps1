# build.ps1 - Gera os executaveis (backup_tool.exe e backup_gui.exe) em ..\programa
#
# Uso: abra o PowerShell nesta pasta (codigo-fonte) e rode:  .\build.ps1
# Requer: Python 3.x 64-bit com os pacotes de requirements.txt + pyinstaller.

$ErrorActionPreference = "Stop"

# Caminho do Python (ajuste se necessario). Cai para "python" do PATH se nao achar.
$py = "C:\Users\admin\AppData\Local\Programs\Python\Python315\python.exe"
if (-not (Test-Path $py)) { $py = "python" }

$fonte = $PSScriptRoot
$saida = Join-Path $PSScriptRoot "..\programa"
$trabalho = Join-Path $env:TEMP "bt_build"

New-Item -ItemType Directory -Force $saida | Out-Null

Write-Host ">>> Gerando backup_tool.exe (CLI)..." -ForegroundColor Cyan
& $py -m PyInstaller --onefile --name backup_tool --noconfirm --clean `
    --distpath $saida --workpath $trabalho --specpath $fonte `
    --hidden-import win32timezone --hidden-import openpyxl --exclude-module tkinter `
    (Join-Path $fonte "main.py")

Write-Host ">>> Gerando backup_gui.exe (interface grafica)..." -ForegroundColor Cyan
& $py -m PyInstaller --onefile --windowed --name backup_gui --noconfirm --clean `
    --distpath $saida --workpath $trabalho --specpath $fonte `
    (Join-Path $fonte "gui.py")

Write-Host ""
Write-Host "Pronto! Executaveis em: $saida" -ForegroundColor Green
Write-Host "Lembre-se de manter config.ini e a planilha ao lado dos .exe." -ForegroundColor Yellow
