# ================================================================
#  BUILD SCRIPT — Ferramenta EACE NOC
#  Executa os 3 passos em sequência:
#  1. PNG -> ICO (ícone do instalador)
#  2. PyInstaller -> .exe
#  3. Inno Setup -> Instalador .exe
#
#  USO: Execute este script na pasta do projeto com PowerShell
#  Para atualizar versão: altere $AppVersion e $ISSVersion em
#  Ferramenta_EACE_NOC_Setup.iss
# ================================================================

$ProjectDir  = $PSScriptRoot
if (-not $ProjectDir) { $ProjectDir = Get-Location }
$PythonScript = "FerramentaEACEv6.py"
$ISSFile     = "$ProjectDir\Ferramenta_EACE_NOC_Setup.iss"
$OutputDir   = "$ProjectDir\installer_output"

# Localizações conhecidas do Inno Setup (ordem de preferência)
$InnoLocations = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 7\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 7\ISCC.exe",
    "C:\Program Files\Inno Setup 7\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  BUILD — Ferramenta EACE NOC v6.0 — Bitnet Telecom" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------
# PASSO 1: Converter logo_bitnet.png -> logo_bitnet.ico
# ---------------------------------------------------------------
Write-Host "[1/3] Convertendo logo_bitnet.png para .ico..." -ForegroundColor Yellow

# Usar arquivo Python temporário para evitar problemas de escape no PowerShell
$convertPy = "$ProjectDir\_convert_ico_temp.py"
@"
from PIL import Image
import sys, os
src = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo_bitnet.png')
dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo_bitnet.ico')
try:
    img = Image.open(src).convert('RGBA')
    img.save(dst, format='ICO', sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])
    print('ICO gerado:', dst)
except Exception as e:
    print('ERRO ao gerar ICO:', e)
    sys.exit(1)
"@ | Out-File -FilePath $convertPy -Encoding UTF8

Set-Location $ProjectDir
python $convertPy
if ($LASTEXITCODE -ne 0) {
    Remove-Item $convertPy -Force -ErrorAction SilentlyContinue
    Write-Host "  ERRO: Falha ao converter PNG para ICO. Abortando." -ForegroundColor Red
    exit 1
}
Remove-Item $convertPy -Force -ErrorAction SilentlyContinue
Write-Host "  OK: logo_bitnet.ico gerado." -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------
# PASSO 2: Compilar o executável com PyInstaller
# ---------------------------------------------------------------
Write-Host "[2/3] Compilando .exe com PyInstaller..." -ForegroundColor Yellow

Set-Location $ProjectDir
python -m PyInstaller `
    --noconfirm `
    --onefile `
    --windowed `
    --splash="Gemini_Generated_logo_bitnet.png" `
    --icon=logo_bitnet.ico `
    --add-data "logo_bitnet.png;." `
    --add-data "Gemini_Generated_logo_bitnet.png;." `
    --name "Ferramenta_EACE_NOC" `
    $PythonScript

if ($LASTEXITCODE -ne 0) {
    Write-Host "  ERRO: Falha no PyInstaller. Abortando." -ForegroundColor Red
    exit 1
}
Write-Host "  OK: dist\Ferramenta_EACE_NOC.exe gerado." -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------
# PASSO 3: Compilar o instalador com Inno Setup
# ---------------------------------------------------------------
Write-Host "[3/3] Compilando instalador com Inno Setup..." -ForegroundColor Yellow

$InnoCompiler = $null
foreach ($loc in $InnoLocations) {
    if (Test-Path $loc) {
        $InnoCompiler = $loc
        break
    }
}

if ($null -eq $InnoCompiler) {
    Write-Host "  ERRO: Inno Setup não encontrado nos caminhos padrão." -ForegroundColor Red
    Write-Host "  Caminhos verificados:" -ForegroundColor Red
    $InnoLocations | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    exit 1
}

Write-Host "  Usando: $InnoCompiler" -ForegroundColor Gray
& $InnoCompiler $ISSFile

if ($LASTEXITCODE -ne 0) {
    Write-Host "  ERRO: Inno Setup retornou erro. Verifique o script .iss." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  BUILD CONCLUÍDO COM SUCESSO!" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Instalador gerado em:" -ForegroundColor White
Write-Host "  $OutputDir\" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Para instalar/atualizar em qualquer PC:" -ForegroundColor White
Write-Host "  Execute o arquivo Ferramenta_EACE_NOC_Setup_v*.exe" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Para atualização silenciosa (sem cliques):" -ForegroundColor White
Write-Host '  Ferramenta_EACE_NOC_Setup_v*.exe /SILENT' -ForegroundColor Cyan
Write-Host ""
