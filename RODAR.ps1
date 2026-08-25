param([string]$Modo = "teste")
# Automacao ASA - abre o app, espera o login e roda os lancamentos.
$ErrorActionPreference = "SilentlyContinue"
Set-Location $PSScriptRoot

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   AUTOMACAO ASA - Lancamento de Atendimentos  (modo: $Modo)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1) Planilha
$planilha = Read-Host "Arraste a planilha (.xlsx) para esta janela e tecle ENTER"
$planilha = $planilha.Trim().Trim('"')
if (-not (Test-Path $planilha)) {
    Write-Host "Planilha nao encontrada: $planilha" -ForegroundColor Red
    Read-Host "Tecle ENTER para sair"; exit
}

# 2) Fecha somente o Chrome do perfil da automacao (nao mexe no seu Chrome normal)
Write-Host "Preparando o navegador do perfil da automacao..." -ForegroundColor Yellow
Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
    Where-Object { $_.CommandLine -like "*perfil-chrome*" } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Start-Sleep -Seconds 2

# 3) Abre o app do ASA nesse perfil com a porta de depuracao
$chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if (-not (Test-Path $chrome)) { $chrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" }
Start-Process $chrome -ArgumentList @(
    "--user-data-dir=$PSScriptRoot\perfil-chrome",
    "--app=https://asa-externo.am.sebrae.com.br/",
    "--remote-debugging-port=9222",
    "--window-size=1400,1000"
)
Start-Sleep -Seconds 3

Write-Host ""
Write-Host ">>> Se pedir, FACA O LOGIN na janela do ASA que abriu." -ForegroundColor Green
Write-Host ">>> Quando estiver vendo a LISTA de atendimentos, volte aqui." -ForegroundColor Green
Read-Host "Tecle ENTER para comecar os lancamentos"

# 4) Roda o script
$env:ASA_CDP = "9222"
$env:ASA_PLANILHA = $planilha
if ($Modo -eq "real") {
    $env:ASA_MODO_TESTE = "0"; $env:ASA_LIMITE = "0"
} else {
    $env:ASA_MODO_TESTE = "1"; $env:ASA_LIMITE = "1"   # preenche 1 e NAO finaliza
}
python "$PSScriptRoot\lancar_atendimentos.py"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   Terminou. Confira a lista no app." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Read-Host "Tecle ENTER para fechar"
