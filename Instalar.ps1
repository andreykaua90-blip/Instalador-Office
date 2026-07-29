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

$Arch = if ([Environment]::Is64BitOperatingSystem) { "64" } else { "32" }

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "   Instalador Automatico - Office LTSC" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Selecione APENAS os programas que deseja INSTALAR:" -ForegroundColor Yellow
Write-Host "[1] Word"
Write-Host "[2] Excel"
Write-Host "[3] PowerPoint"
Write-Host "[4] Access"
Write-Host "[5] Outlook"
Write-Host "[6] OneNote"
Write-Host "[7] Publisher"
Write-Host "[8] OneDrive"
Write-Host "[9] Lync (Skype)"
Write-Host ""
$Opcoes = Read-Host "Digite os numeros desejados separados por virgula (ex: 1,2,3)"

$OpcoesArray = $Opcoes -split "," | ForEach-Object { $_.Trim() }

$Excludes = ""
if ("1" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"Word`" />`n" }
if ("2" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"Excel`" />`n" }
if ("3" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"PowerPoint`" />`n" }
if ("4" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"Access`" />`n" }
if ("5" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"Outlook`" />`n" }
if ("6" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"OneNote`" />`n" }
if ("7" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"Publisher`" />`n" }
if ("8" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"OneDrive`" />`n" }
if ("9" -notin $OpcoesArray) { $Excludes += "      <ExcludeApp ID=`"Lync`" />`n" }

$XmlContent = @"
<Configuration ID="0e135344-f323-430f-9825-667da45ea0b2">
  <Add OfficeClientEdition="$Arch" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB">
      <Language ID="MatchOS" />
$Excludes    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
</Configuration>
"@

$XmlContent | Out-File -FilePath ".\config.xml" -Encoding UTF8

Write-Host ""
Write-Host "[1/3] Baixando ferramenta de instalacao e extraindo..." -ForegroundColor Yellow
Invoke-WebRequest -Uri "$BaseUrl/$ArquivoOdt" -OutFile $ArquivoOdt

Start-Process -FilePath ".\$ArquivoOdt" -ArgumentList "/extract:.\ /quiet" -Wait -NoNewWindow

if (-Not (Test-Path "setup.exe")) {
    Write-Error "[ERRO] Falha ao extrair o setup.exe. Verifique o nome do executavel."
    Set-Location -Path "C:\"
    Remove-Item -Path $Dir -Recurse -Force
    Break
}

Write-Host ""
Write-Host "[2/3] Instalando o Office ($Arch bits)..." -ForegroundColor Yellow
Write-Host "Aguarde a janela da Microsoft concluir o progresso. Isso pode demorar alguns minutos." -ForegroundColor Gray
Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure config.xml" -Wait -NoNewWindow

Write-Host ""
Write-Host "[3/3] Iniciando o processo de ativacao automatica (Ohook)..." -ForegroundColor Yellow

$MAS = Invoke-RestMethod -Uri 'https://get.activated.win'
Invoke-Command -ScriptBlock ([ScriptBlock]::Create($MAS)) -ArgumentList "/ohook"

Write-Host ""
Write-Host "Limpando arquivos temporarios..." -ForegroundColor Gray
Set-Location -Path "C:\"
Remove-Item -Path $Dir -Recurse -Force

Write-Host "========================================================" -ForegroundColor Green
Write-Host "   Processo finalizado com sucesso!" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
