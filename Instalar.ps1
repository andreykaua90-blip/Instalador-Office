$ArquivoXml = "Configuração.xml"
$ArquivoOdt = "officedeploymenttool_20131-20090.exe"

$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"
Clear-Host

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Permissao negada! Por favor, abra o PowerShell como Administrador e rode o comando novamente."
    Break
}

$BaseUrl = "https://raw.githubusercontent.com/andreykaua90-blip/Instalador-Office/main"

$Dir = "C:\Office_Temp"
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
Set-Location -Path $Dir

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "   Instalador Automatico - Office LTSC" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/3] Baixando arquivos do GitHub e extraindo..." -ForegroundColor Yellow
Invoke-WebRequest -Uri "$BaseUrl/$ArquivoXml" -OutFile $ArquivoXml
Invoke-WebRequest -Uri "$BaseUrl/$ArquivoOdt" -OutFile $ArquivoOdt

Start-Process -FilePath ".\$ArquivoOdt" -ArgumentList "/extract:.\ /quiet" -Wait -NoNewWindow

if (-Not (Test-Path "setup.exe")) {
    Write-Error "[ERRO] Falha ao extrair o setup.exe. Verifique o nome do executavel."
    Set-Location -Path "C:\"
    Remove-Item -Path $Dir -Recurse -Force
    Break
}

Write-Host ""
Write-Host "[2/3] Instalando o Office..." -ForegroundColor Yellow
Write-Host "Aguarde a janela da Microsoft concluir o progresso. Isso pode demorar alguns minutos." -ForegroundColor Gray
Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure $ArquivoXml" -Wait -NoNewWindow

Write-Host ""
Write-Host "[3/3] Iniciando o processo de ativacao (Ohook)..." -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "ATENCAO: Siga a sequencia no menu azul que vai carregar:" 
Write-Host ""
Write-Host "1. No menu principal, digite a OPCAO [2] (Ohook)."
Write-Host "2. No proximo menu, digite a OPCAO [1] (Install Ohook)."
Write-Host "3. Quando a ativacao terminar, digite [0] para sair."
Write-Host "========================================================" -ForegroundColor Cyan
Start-Sleep -Seconds 4

Invoke-Expression (Invoke-RestMethod -Uri "https://get.activated.win")

Write-Host ""
Write-Host "Limpando arquivos temporarios..." -ForegroundColor Gray
Set-Location -Path "C:\"
Remove-Item -Path $Dir -Recurse -Force

Write-Host "========================================================" -ForegroundColor Green
Write-Host "   Processo finalizado com sucesso!" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
