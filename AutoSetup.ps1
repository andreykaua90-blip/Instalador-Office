$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"
Clear-Host
$ProgressPreference = 'SilentlyContinue'

$SysLang = (Get-UICulture).TwoLetterISOLanguageName
if ($SysLang -ne "pt") {
    $MsgAdmin     = "Permission denied! Please open PowerShell as Administrator and run the command again."
    $Title        = "Automatic Installer / Uninstaller - Office"
    $MsgMenu      = "What do you want to do?"
    $MsgSelectVer = "Select the Office version:"
    $MsgSelect    = "Select the programs you want to INSTALL:"
    $MsgInput     = "Enter the numbers separated by comma (e.g. 1,2,3) or 0 to go back"
    $MsgStep1     = "[1/3] Downloading Office Deployment Tool and extracting..."
    $MsgError     = "[ERROR] Failed to extract setup.exe."
    $MsgStep2     = "[2/3] Installing Office"
    $MsgWait      = "Please wait... This may take a few minutes."
    $MsgStep3     = "[3/3] Starting activation (Ohook)..."
    $MsgClean     = "Cleaning up temporary files..."
    $MsgSuccess   = "Process finished successfully!"
    $MsgODTFail   = "Could not locate the Office Deployment Tool download link."
    $MsgActivate  = "Do you want to activate Office now using Ohook? (Y/N)"
    $MsgSkipped   = "Activation skipped."
    $MsgUninst    = "Uninstalling Office completely..."
    $MsgUninstOk  = "Office uninstalled successfully!"
    $MsgAskAct    = "Do you also want to remove the activation (Ohook)? (Y/N)"
    $MsgRemAct    = "Removing activation..."
    $MsgInvalid   = "Invalid option."
    $MsgConfirm   = "Confirm installation? (Y = Install / N = Back)"
    $MsgCancel    = "Returning..."
} else {
    $MsgAdmin     = "Permissao negada! Por favor, abra o PowerShell como Administrador e rode o comando novamente."
    $Title        = "Instalador / Desinstalador Automatico - Office"
    $MsgMenu      = "O que deseja fazer?"
    $MsgSelectVer = "Selecione a versao do Office:"
    $MsgSelect    = "Selecione os programas que deseja INSTALAR:"
    $MsgInput     = "Digite os numeros separados por virgula (ex: 1,2,3) ou 0 para voltar"
    $MsgStep1     = "[1/3] Baixando Office Deployment Tool e extraindo..."
    $MsgError     = "[ERRO] Falha ao extrair o setup.exe."
    $MsgStep2     = "[2/3] Instalando o Office"
    $MsgWait      = "Aguarde... Isso pode demorar alguns minutos."
    $MsgStep3     = "[3/3] Iniciando ativacao (Ohook)..."
    $MsgClean     = "Limpando arquivos temporarios..."
    $MsgSuccess   = "Processo finalizado com sucesso!"
    $MsgODTFail   = "Nao foi possivel localizar o link de download do Office Deployment Tool."
    $MsgActivate  = "Deseja ativar o Office agora usando Ohook? (S/N)"
    $MsgSkipped   = "Ativacao ignorada."
    $MsgUninst    = "Desinstalando o Office completamente..."
    $MsgUninstOk  = "Office desinstalado com sucesso!"
    $MsgAskAct    = "Deseja tambem remover a ativacao (Ohook)? (S/N)"
    $MsgRemAct    = "Removendo ativacao..."
    $MsgInvalid   = "Opcao invalida."
    $MsgConfirm   = "Confirmar instalacao? (S = Instalar / N = Voltar)"
    $MsgCancel    = "Voltando..."
}

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning $MsgAdmin
    Break
}

$Dir = "C:\Office_Temp"
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
Set-Location -Path $Dir
$Arch = if ([Environment]::Is64BitOperatingSystem) { "64" } else { "32" }

function Show-Header {
    Clear-Host
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
}

# ====================== MENU PRINCIPAL ======================
:MenuPrincipal
while ($true) {
    Show-Header
    Write-Host $MsgMenu -ForegroundColor Yellow
    Write-Host "[1] Instalar Office"
    Write-Host "[2] Desinstalar Office"
    Write-Host "[3] Sair"
    Write-Host ""
    $Acao = Read-Host "Digite 1, 2 ou 3"

    if ($Acao -eq "3") {
        Set-Location C:\
        Remove-Item -Path $Dir -Recurse -Force -ErrorAction SilentlyContinue
        exit
    }

    if ($Acao -eq "2") {
        # ========== DESINSTALAR ==========
        Show-Header
        Write-Host $MsgUninst -ForegroundColor Yellow
        Write-Host ""

        $DownloadPage = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
        Write-Host "Obtendo link de download do ODT..." -ForegroundColor Gray
        try {
            $html = Invoke-WebRequest -Uri $DownloadPage -UseBasicParsing -ErrorAction Stop
            $regex = 'https://download\.microsoft\.com/download/[^\s"''<>]+officedeploymenttool[^\s"''<>]+\.exe'
            if ($html.Content -match $regex) {
                $DownloadUrl = $matches[0]
                $Arquivo = Join-Path $Dir "OfficeDeploymentTool.exe"
                Write-Host "Baixando ODT..." -ForegroundColor Gray
                Invoke-WebRequest -Uri $DownloadUrl -OutFile $Arquivo -ErrorAction Stop
                Write-Host "Extraindo arquivos..." -ForegroundColor Gray
                Start-Process -FilePath $Arquivo -ArgumentList "/quiet /extract:`"$Dir`"" -Wait -NoNewWindow
            } else {
                Write-Error $MsgODTFail
                Pause
                continue MenuPrincipal
            }
        } catch {
            Write-Error "Erro ao baixar o ODT: $($_.Exception.Message)"
            Pause
            continue MenuPrincipal
        }

        if (-Not (Test-Path ".\setup.exe")) {
            Write-Error $MsgError
            Pause
            continue MenuPrincipal
        }

        $RemoveXml = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
        $RemoveXml | Out-File ".\remove.xml" -Encoding UTF8

        Write-Host ""
        Write-Host $MsgWait -ForegroundColor Yellow
        Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure remove.xml" -Wait -WindowStyle Hidden

        Write-Host ""
        Write-Host "Desinstalacao concluida." -ForegroundColor Green
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

        Set-Location C:\
        Remove-Item -Path $Dir -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Force -Path $Dir | Out-Null
        Set-Location $Dir

        Write-Host ""
        Write-Host "========================================================" -ForegroundColor Green
        Write-Host " $MsgUninstOk" -ForegroundColor Green
        Write-Host "========================================================" -ForegroundColor Green
        Pause
        continue MenuPrincipal
    }

    if ($Acao -ne "1") {
        Write-Host $MsgInvalid -ForegroundColor Red
        Start-Sleep -Seconds 1
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
        Write-Host "[0] Voltar"
        Write-Host ""
        $Versao = Read-Host "Digite o numero"

        if ($Versao -eq "0") { continue MenuPrincipal }

        switch ($Versao) {
            "1" {
                $Channel    = "PerpetualVL2024"
                $ProductID  = "ProPlus2024Volume"
                $PIDKEY     = "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"
                $NomeVersao = "Office LTSC Professional Plus 2024"
                break
            }
            "2" {
                $Channel    = "PerpetualVL2021"
                $ProductID  = "ProPlus2021Volume"
                $PIDKEY     = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
                $NomeVersao = "Office LTSC Professional Plus 2021"
                break
            }
            "3" {
                $Channel    = "PerpetualVL2019"
                $ProductID  = "ProPlus2019Volume"
                $PIDKEY     = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
                $NomeVersao = "Office Professional Plus 2019"
                break
            }
            default {
                Write-Host $MsgInvalid -ForegroundColor Red
                Start-Sleep -Seconds 1
                continue EscolherVersao
            }
        }

        # ===== ESCOLHER APLICATIVOS =====
        :EscolherApps
        while ($true) {
            Show-Header
            Write-Host "Versao: $NomeVersao" -ForegroundColor Green
            Write-Host ""
            Write-Host $MsgSelect -ForegroundColor Yellow
            Write-Host "[1] Word"
            Write-Host "[2] Excel"
            Write-Host "[3] PowerPoint"
            Write-Host "[4] Access"
            Write-Host "[5] Outlook"
            Write-Host "[6] OneNote"
            Write-Host "[7] Publisher"
            Write-Host "[8] OneDrive"
            Write-Host "[9] Lync (Skype)"
            Write-Host "[0] Voltar"
            Write-Host ""
            $Opcoes = Read-Host $MsgInput

            if ($Opcoes -eq "0") { continue EscolherVersao }

            $OpcoesArray = $Opcoes -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[1-9]$' }

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
            if ($Versao -eq "3") { $Excludes += "      <ExcludeApp ID=`"Groove`" />`n" }

            # ===== CONFIRMAÇÃO =====
            Show-Header
            Write-Host "RESUMO DA INSTALACAO" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "Versao: $NomeVersao" -ForegroundColor Yellow
            Write-Host "Arquitetura: $Arch bits" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Aplicativos que SERAO instalados:" -ForegroundColor Green

            $Apps = @("Word","Excel","PowerPoint","Access","Outlook","OneNote","Publisher","OneDrive","Lync")
            for ($i = 1; $i -le 9; $i++) {
                if ("$i" -in $OpcoesArray) {
                    Write-Host "  [X] $($Apps[$i-1])" -ForegroundColor Green
                } else {
                    Write-Host "  [ ] $($Apps[$i-1])" -ForegroundColor DarkGray
                }
            }

            Write-Host ""
            $Confirma = Read-Host $MsgConfirm

            if ($Confirma -match '^[nN]') {
                continue EscolherApps   # Volta para escolher os apps
            }

            if ($Confirma -notmatch '^[sSyY]') {
                continue EscolherApps
            }

            # ===== USUÁRIO CONFIRMOU → CRIA O XML E INSTALA =====
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

            $DownloadPage = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
            try {
                $html = Invoke-WebRequest -Uri $DownloadPage -UseBasicParsing -ErrorAction Stop
                $regex = 'https://download\.microsoft\.com/download/[^\s"''<>]+officedeploymenttool[^\s"''<>]+\.exe'
                if ($html.Content -match $regex) {
                    $DownloadUrl = $matches[0]
                    $Arquivo = Join-Path $Dir "OfficeDeploymentTool.exe"
                    Write-Host "Baixando ODT..." -ForegroundColor Gray
                    Invoke-WebRequest -Uri $DownloadUrl -OutFile $Arquivo -ErrorAction Stop
                    Write-Host "Extraindo arquivos..." -ForegroundColor Gray
                    Start-Process -FilePath $Arquivo -ArgumentList "/quiet /extract:`"$Dir`"" -Wait -NoNewWindow
                } else {
                    Write-Error $MsgODTFail
                    Pause
                    continue MenuPrincipal
                }
            } catch {
                Write-Error "Erro ao baixar o ODT: $($_.Exception.Message)"
                Pause
                continue MenuPrincipal
            }

            if (-Not (Test-Path ".\setup.exe")) {
                Write-Error $MsgError
                Pause
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
            Set-Location C:\
            Remove-Item -Path $Dir -Recurse -Force -ErrorAction SilentlyContinue

            Write-Host ""
            Write-Host "========================================================" -ForegroundColor Green
            Write-Host " $MsgSuccess" -ForegroundColor Green
            Write-Host "========================================================" -ForegroundColor Green
            Pause
            exit
        }
    }
}
