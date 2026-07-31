$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"
Clear-Host
$ProgressPreference = 'SilentlyContinue'

# Força TLS 1.2 para evitar erros de download de rede em sistemas mais antigos
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$SysLang = (Get-UICulture).TwoLetterISOLanguageName

if ($SysLang -ne "pt") {
    # ==================== ENGLISH ====================
    $MsgAdmin       = "Permission denied! Please open PowerShell as Administrator and run the command again."
    $Title          = "Automatic Installer / Uninstaller - Office"
    $MsgMenu        = "What do you want to do?"
    $MsgOpt1        = "[1] Install Office"
    $MsgOpt2        = "[2] Uninstall Office"
    $MsgOpt0        = "[0] Exit"
    $MsgSelectVer   = "Select the Office version:"
    $MsgSelect      = "Select the programs you want to INSTALL:"
    $MsgSelectRem   = "Select the programs you want to REMOVE:"
    $MsgInput       = "Enter the numbers separated by comma (e.g. 1,2,3) or A for ALL or 0 to go back"
    $MsgStep1       = "[1/3] Downloading Office Deployment Tool and extracting..."
    $MsgError       = "[ERROR] Failed to extract setup.exe."
    $MsgStep2       = "[2/3] Installing Office"
    $MsgWait        = "Please wait... This may take a few minutes."
    $MsgStep3       = "[3/3] Starting activation (Ohook)..."
    $MsgClean       = "Cleaning up temporary files..."
    $MsgSuccess     = "Process finished successfully!"
    $MsgODTFail     = "Could not locate the Office Deployment Tool download link."
    $MsgActivate    = "Do you want to activate Office now using Ohook? (Y/N)"
    $MsgSkipped     = "Activation skipped."
    $MsgUninstOk    = "Office uninstalled successfully!"
    $MsgAskAct      = "Do you also want to remove the activation (Ohook)? (Y/N)"
    $MsgRemAct      = "Removing activation..."
    $MsgInvalid     = "Invalid option."
    $MsgConfirm     = "Confirm? (Y = Continue / N = Back)"
    $MsgBack        = "[0] Back"
    $MsgAll         = "[A] All"
    $MsgPrompt      = "Enter your choice"
    $MsgGettingODT  = "Getting Office Deployment Tool download link..."
    $MsgDownloading = "Downloading ODT..."
    $MsgExtracting  = "Extracting files..."
    $MsgWaitWindow  = "The Microsoft window should appear now. Wait for it to close by itself."
    $MsgRemoved     = "Applications that will be REMOVED:"
    $MsgInstalled   = "Applications that will be INSTALLED:"
    $MsgSummary     = "INSTALLATION SUMMARY"
    $MsgVersion     = "Version:"
    $MsgArch        = "Architecture:"
    $MsgUninstDone  = "Uninstallation completed."
    $MsgPressExit   = "Press Enter to exit..."
} else {
    # ==================== PORTUGUÊS ====================
    $MsgAdmin       = "Permissao negada! Por favor, abra o PowerShell como Administrador e rode o comando novamente."
    $Title          = "Instalador / Desinstalador Automatico - Office"
    $MsgMenu        = "O que deseja fazer?"
    $MsgOpt1        = "[1] Instalar Office"
    $MsgOpt2        = "[2] Desinstalar Office"
    $MsgOpt0        = "[0] Sair"
    $MsgSelectVer   = "Selecione a versao do Office:"
    $MsgSelect      = "Selecione os programas que deseja INSTALAR:"
    $MsgSelectRem   = "Selecione os programas que deseja REMOVER:"
    $MsgInput       = "Digite os numeros separados por virgula (ex: 1,2,3) ou A para TODOS ou 0 para voltar"
    $MsgStep1       = "[1/3] Baixando Office Deployment Tool e extraindo..."
    $MsgError       = "[ERRO] Falha ao extrair o setup.exe."
    $MsgStep2       = "[2/3] Instalando o Office"
    $MsgWait        = "Aguarde... Isso pode demorar alguns minutos."
    $MsgStep3       = "[3/3] Iniciando ativacao (Ohook)..."
    $MsgClean       = "Limpando arquivos temporarios..."
    $MsgSuccess     = "Processo finalizado com sucesso!"
    $MsgODTFail     = "Nao foi possivel localizar o link de download do Office Deployment Tool."
    $MsgActivate    = "Deseja ativar o Office agora usando Ohook? (S/N)"
    $MsgSkipped     = "Ativacao ignorada."
    $MsgUninstOk    = "Office desinstalado com sucesso!"
    $MsgAskAct      = "Deseja tambem remover a ativacao (Ohook)? (S/N)"
    $MsgRemAct      = "Removendo ativacao..."
    $MsgInvalid     = "Opcao invalida."
    $MsgConfirm     = "Confirmar? (S = Continuar / N = Voltar)"
    $MsgBack        = "[0] Voltar"
    $MsgAll         = "[A] Todos"
    $MsgPrompt      = "Digite sua opcao"
    $MsgGettingODT  = "Obtendo link de download do ODT..."
    $MsgDownloading = "Baixando ODT..."
    $MsgExtracting  = "Extraindo arquivos..."
    $MsgWaitWindow  = "A janela da Microsoft deve aparecer agora. Aguarde ela fechar sozinha."
    $MsgRemoved     = "Aplicativos que SERAO REMOVIDOS:"
    $MsgInstalled   = "Aplicativos que SERAO instalados:"
    $MsgSummary     = "RESUMO DA INSTALACAO"
    $MsgVersion     = "Versao:"
    $MsgArch        = "Arquitetura:"
    $MsgUninstDone  = "Desinstalacao concluida."
    $MsgPressExit   = "Pressione Enter para sair..."
}

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning $MsgAdmin
    Break
}

$Dir = "C:\Office_Temp"
$Arch = if ([Environment]::Is64BitOperatingSystem) { "64" } else { "32" }
$AppsList = @("Word","Excel","PowerPoint","Access","Outlook","OneNote","Publisher","OneDrive","Lync")

function Show-Header {
    Clear-Host
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Reset-TempDir {
    Set-Location C:\
    if (Test-Path $Dir) { Remove-Item -Path $Dir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Force -Path $Dir | Out-Null
    Set-Location -Path $Dir
}

function Clear-TempDir {
    Set-Location C:\
    if (Test-Path $Dir) { Remove-Item -Path $Dir -Recurse -Force -ErrorAction SilentlyContinue }
}

function Get-ODT {
    Write-Host $MsgGettingODT -ForegroundColor Gray
    try {
        $html = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing -ErrorAction Stop
        $regex = 'https://download\.microsoft\.com/download/[^\s"''<>]+officedeploymenttool[^\s"''<>]+\.exe'
        if ($html.Content -match $regex) {
            $DownloadUrl = $matches[0]
            $Arquivo = Join-Path $Dir "OfficeDeploymentTool.exe"
            Write-Host $MsgDownloading -ForegroundColor Gray
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $Arquivo -ErrorAction Stop
            Write-Host $MsgExtracting -ForegroundColor Gray
            Start-Process -FilePath $Arquivo -ArgumentList "/quiet /extract:`"$Dir`"" -WindowStyle Hidden -Wait
            return $true
        } else {
            Write-Error $MsgODTFail
            return $false
        }
    } catch {
        Write-Error "Erro ao baixar o ODT: $($_.Exception.Message)"
        return $false
    }
}

function Get-ExcludesXml($SelectedArray) {
    $ExcludesStr = ""
    for ($i = 0; $i -lt $AppsList.Count; $i++) {
        $AppNum = ($i + 1).ToString()
        if ($AppNum -notin $SelectedArray) {
            $ExcludesStr += "      <ExcludeApp ID=`"$($AppsList[$i])`" />`n"
        }
    }
    return $ExcludesStr
}

# ====================== MENU PRINCIPAL ======================
:MenuPrincipal
while ($true) {
    Show-Header
    Write-Host $MsgMenu -ForegroundColor Yellow
    Write-Host $MsgOpt1
    Write-Host $MsgOpt2
    Write-Host ""
    Write-Host $MsgOpt0
    Write-Host ""
    $Acao = Read-Host $MsgPrompt

    if ($Acao -eq "0") {
        Clear-TempDir
        exit
    }

    # ====================== DESINSTALAR ======================
    if ($Acao -eq "2") {
        :EscolherAppsRemover
        while ($true) {
            Show-Header
            Write-Host $MsgSelectRem -ForegroundColor Yellow
            for ($i = 0; $i -lt $AppsList.Count; $i++) {
                Write-Host "[$($i+1)] $($AppsList[$i])"
            }
            Write-Host ""
            Write-Host $MsgAll
            Write-Host $MsgBack
            Write-Host ""
            $Opcoes = Read-Host $MsgInput

            if ($Opcoes -eq "0") { continue MenuPrincipal }

            $RemoverTudo = ($Opcoes -match '^[aA]$')
            if ($RemoverTudo) {
                $OpcoesArray = 1..9 | ForEach-Object { $_.ToString() }
            } else {
                $OpcoesArray = $Opcoes -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[1-9]$' }
            }

            if ($OpcoesArray.Count -eq 0) {
                Write-Host $MsgInvalid -ForegroundColor Red
                Start-Sleep 1
                continue EscolherAppsRemover
            }

            Show-Header
            Write-Host $MsgRemoved -ForegroundColor Yellow
            for ($i = 0; $i -lt $AppsList.Count; $i++) {
                $AppNum = ($i + 1).ToString()
                if ($AppNum -in $OpcoesArray) {
                    Write-Host "  [X] $($AppsList[$i])" -ForegroundColor Red
                } else {
                    Write-Host "  [ ] $($AppsList[$i])" -ForegroundColor DarkGray
                }
            }

            Write-Host ""
            $Confirma = Read-Host $MsgConfirm
            if ($Confirma -notmatch '^[sSyY]') { continue EscolherAppsRemover }

            Reset-TempDir

            if (-not (Get-ODT) -or -not (Test-Path ".\setup.exe")) { 
                if (-not (Test-Path ".\setup.exe")) { Write-Error $MsgError }
                Write-Host $MsgPressExit -ForegroundColor Gray
                $null = Read-Host
                continue MenuPrincipal
            }

            if ($RemoverTudo) {
                $RemoveXml = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
            } else {
                $Excludes = Get-ExcludesXml -SelectedArray $OpcoesArray
                $RemoveXml = @"
<Configuration>
  <Add OfficeClientEdition="$Arch" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="MatchOS" />
$Excludes    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <RemoveMSI />
</Configuration>
"@
            }

            $RemoveXml | Out-File ".\remove.xml" -Encoding UTF8

            Write-Host ""
            Write-Host "Iniciando a desinstalacao silenciosa..." -ForegroundColor Yellow
            Write-Host "O Office esta sendo removido no fundo. Aguarde" -NoNewline -ForegroundColor Gray
            
            # Inicia o desinstalador oculto
            $SetupProc = Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure remove.xml" -WindowStyle Hidden -PassThru
            
            # Loop que imprime pontinhos no terminal para mostrar que o script nao travou
            while (-not $SetupProc.HasExited) {
                Start-Sleep -Seconds 2
                Write-Host "." -NoNewline -ForegroundColor Gray
            }
            
            Write-Host " Concluido!" -ForegroundColor Green

            Write-Host ""
            Write-Host $MsgUninstDone -ForegroundColor Green
            Write-Host ""
            $RemoverAtivacao = Read-Host $MsgAskAct

            if ($RemoverAtivacao -match '^[sSyY]') {
                Write-Host $MsgRemAct -ForegroundColor Yellow
                try {
                    $MAS = Invoke-RestMethod -Uri 'https://get.activated.win'
                    Invoke-Command -ScriptBlock ([ScriptBlock]::Create($MAS)) -ArgumentList "/Ohook-Uninstall"
                } catch {
                    Write-Warning "Nao foi possivel remover a ativacao."
                }
            }

            Clear-TempDir

            Write-Host ""
            Write-Host "========================================================" -ForegroundColor Green
            Write-Host " $MsgUninstOk" -ForegroundColor Green
            Write-Host "========================================================" -ForegroundColor Green
            Write-Host ""
            Write-Host $MsgPressExit -ForegroundColor Gray
            $null = Read-Host
            continue MenuPrincipal
        }
    }

    if ($Acao -ne "1") {
        Write-Host $MsgInvalid -ForegroundColor Red
        Start-Sleep 1
        continue MenuPrincipal
    }

    # ====================== INSTALAR ======================
    :EscolherVersao
    while ($true) {
        Show-Header
        Write-Host $MsgSelectVer -ForegroundColor Yellow
        Write-Host "[1] Office LTSC Professional Plus 2024"
        Write-Host "[2] Office LTSC Professional Plus 2021"
        Write-Host "[3] Office Professional Plus 2019"
        Write-Host ""
        Write-Host $MsgBack
        Write-Host ""
        $Versao = Read-Host $MsgPrompt

        if ($Versao -eq "0") { continue MenuPrincipal }

        switch ($Versao) {
            "1" {
                $Channel = "PerpetualVL2024"; $ProductID = "ProPlus2024Volume"
                $PIDKEY = "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"; $NomeVersao = "Office LTSC Professional Plus 2024"
            }
            "2" {
                $Channel = "PerpetualVL2021"; $ProductID = "ProPlus2021Volume"
                $PIDKEY = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"; $NomeVersao = "Office LTSC Professional Plus 2021"
            }
            "3" {
                $Channel = "PerpetualVL2019"; $ProductID = "ProPlus2019Volume"
                $PIDKEY = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"; $NomeVersao = "Office Professional Plus 2019"
            }
            default {
                Write-Host $MsgInvalid -ForegroundColor Red
                Start-Sleep 1
                continue EscolherVersao
            }
        }

        :EscolherApps
        while ($true) {
            Show-Header
            Write-Host "$MsgVersion $NomeVersao" -ForegroundColor Green
            Write-Host ""
            Write-Host $MsgSelect -ForegroundColor Yellow
            
            for ($i = 0; $i -lt $AppsList.Count; $i++) {
                Write-Host "[$($i+1)] $($AppsList[$i])"
            }
            
            Write-Host ""
            Write-Host $MsgAll
            Write-Host $MsgBack
            Write-Host ""
            $Opcoes = Read-Host $MsgInput

            if ($Opcoes -eq "0") { continue EscolherVersao }

            if ($Opcoes -match '^[aA]$') {
                $OpcoesArray = 1..9 | ForEach-Object { $_.ToString() }
            } else {
                $OpcoesArray = $Opcoes -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[1-9]$' }
            }

            if ($OpcoesArray.Count -eq 0) {
                Write-Host $MsgInvalid -ForegroundColor Red
                Start-Sleep 1
                continue EscolherApps
            }

            $Excludes = Get-ExcludesXml -SelectedArray $OpcoesArray
            if ($Versao -eq "3") { $Excludes += "      <ExcludeApp ID=`"Groove`" />`n" }

            Show-Header
            Write-Host $MsgSummary -ForegroundColor Cyan
            Write-Host "$MsgVersion $NomeVersao" -ForegroundColor Yellow
            Write-Host "$MsgArch $Arch bits" -ForegroundColor Yellow
            Write-Host ""
            Write-Host $MsgInstalled -ForegroundColor Green
            
            for ($i = 0; $i -lt $AppsList.Count; $i++) {
                $AppNum = ($i + 1).ToString()
                if ($AppNum -in $OpcoesArray) {
                    Write-Host "  [X] $($AppsList[$i])" -ForegroundColor Green
                } else {
                    Write-Host "  [ ] $($AppsList[$i])" -ForegroundColor DarkGray
                }
            }

            Write-Host ""
            $Confirma = Read-Host $MsgConfirm
            if ($Confirma -notmatch '^[sSyY]') { continue EscolherApps }

            Reset-TempDir

            $XmlContent = @"
<Configuration ID="$(([guid]::NewGuid()).ToString())">
  <Add OfficeClientEdition="$Arch" Channel="$Channel">
    <Product ID="$ProductID" PIDKEY="$PIDKEY">
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
            $XmlContent | Out-File ".\config.xml" -Encoding UTF8

            Write-Host ""
            Write-Host $MsgStep1 -ForegroundColor Yellow
            if (-not (Get-ODT) -or -not (Test-Path ".\setup.exe")) { 
                if (-not (Test-Path ".\setup.exe")) { Write-Error $MsgError }
                Write-Host $MsgPressExit -ForegroundColor Gray
                $null = Read-Host
                continue MenuPrincipal
            }

            Write-Host ""
            Write-Host "$MsgStep2 - $NomeVersao ($Arch bits)..." -ForegroundColor Yellow
            Write-Host $MsgWait -ForegroundColor Gray
            Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure config.xml" -Wait -NoNewWindow

            Write-Host ""
            $Ativar = Read-Host $MsgActivate
            if ($Ativar -match '^[sSyY]') {
                Write-Host $MsgStep3 -ForegroundColor Yellow
                $MAS = Invoke-RestMethod -Uri 'https://get.activated.win'
                Invoke-Command -ScriptBlock ([ScriptBlock]::Create($MAS)) -ArgumentList "/ohook"
            } else {
                Write-Host $MsgSkipped -ForegroundColor Gray
            }

            Write-Host ""
            Write-Host $MsgClean -ForegroundColor Gray
            Clear-TempDir

            Write-Host ""
            Write-Host "========================================================" -ForegroundColor Green
            Write-Host " $MsgSuccess" -ForegroundColor Green
            Write-Host "========================================================" -ForegroundColor Green
            Write-Host ""
            Write-Host $MsgPressExit -ForegroundColor Gray
            $null = Read-Host
            continue MenuPrincipal
        }
    }
}
