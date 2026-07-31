$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"
Clear-Host
$ProgressPreference = 'SilentlyContinue'

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
            Start-Process -FilePath $Arquivo -ArgumentList "/quiet /extract:`"$Dir`"" -Wait -NoNewWindow
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
        Set-Location C:\
        Remove-Item -Path $Dir -Recurse -Force -ErrorAction SilentlyContinue
        exit
    }

    # ====================== DESINSTALAR ======================
    if ($Acao -eq "2") {
        :EscolherAppsRemover
        while ($true) {
            Show-Header
            Write-Host $MsgSelectRem -ForegroundColor Yellow
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
            Write-Host $MsgAll
            Write-Host $MsgBack
            Write-Host ""
            $Opcoes = Read-Host $MsgInput

            if ($Opcoes -eq "0") { continue MenuPrincipal }

            if ($Opcoes -match '^[aA]$') {
                $OpcoesArray = @("1","2","3","4","5","6","7","8","9")
                $RemoverTudo = $true
            } else {
                $OpcoesArray = $Opcoes -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[1-9]$' }
                $RemoverTudo = $false
            }

            if ($OpcoesArray.Count -eq 0) {
                Write-Host $MsgInvalid -ForegroundColor Red
                Start-Sleep 1
                continue EscolherAppsRemover
            }

            Show-Header
            Write-Host $MsgRemoved -ForegroundColor Yellow
            $Apps = @("Word","Excel","PowerPoint","Access","Outlook","OneNote","Publisher","OneDrive","Lync")
            for ($i = 1; $i -le 9; $i++) {
                if ("$i" -in $OpcoesArray) {
                    Write-Host "  [X] $($Apps[$i-1])" -ForegroundColor Red
                } else {
                    Write-Host "  [ ] $($Apps[$i-1])" -ForegroundColor DarkGray
                }
            }

            Write-Host ""
            $Confirma = Read-Host $MsgConfirm
            if ($Confirma -notmatch '^[sSyY]') { continue EscolherAppsRemover }

            if (-not (Get-ODT)) { 
                Write-Host $MsgPressExit -ForegroundColor Gray
                $null = Read-Host
                exit
            }
            if (-Not (Test-Path ".\setup.exe")) { 
                Write-Error $MsgError
                Write-Host $MsgPressExit -ForegroundColor Gray
                $null = Read-Host
                exit
            }

            if ($RemoverTudo) {
                $RemoveXml = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
            } else {
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

                $RemoveXml = @"
<Configuration>
  <Add OfficeClientEdition="$Arch" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="MatchOS" />
$Excludes    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <RemoveMSI />
</Configuration>
"@
            }

            $RemoveXml | Out-File ".\remove.xml" -Encoding UTF8

            Write-Host ""
            Write-Host $MsgWait -ForegroundColor Yellow
            Write-Host $MsgWaitWindow -ForegroundColor Gray
            Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure remove.xml" -Wait -WindowStyle Hidden

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

            Set-Location C:\
            Remove-Item -Path $Dir -Recurse -Force -ErrorAction SilentlyContinue

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
            Write-Host $MsgAll
            Write-Host $MsgBack
            Write-Host ""
            $Opcoes = Read-Host $MsgInput

            if ($Opcoes -eq "0") { continue EscolherVersao }

            if ($Opcoes -match '^[aA]$') {
                $OpcoesArray = @("1","2","3","4","5","6","7","8","9")
            } else {
                $OpcoesArray = $Opcoes -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[1-9]$' }
            }

            if ($OpcoesArray.Count -eq 0) {
                Write-Host $MsgInvalid -ForegroundColor Red
                Start-Sleep 1
                continue EscolherApps
            }

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

            Show-Header
            Write-Host $MsgSummary -ForegroundColor Cyan
            Write-Host "$MsgVersion $NomeVersao" -ForegroundColor Yellow
            Write-Host "$MsgArch $Arch bits" -ForegroundColor Yellow
            Write-Host ""
            Write-Host $MsgInstalled -ForegroundColor Green
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
            if ($Confirma -notmatch '^[sSyY]') { continue EscolherApps }

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
            if (-not (Get-ODT)) { 
                Write-Host $MsgPressExit -ForegroundColor Gray
                $null = Read-Host
                exit
            }
            if (-Not (Test-Path ".\setup.exe")) { 
                Write-Error $MsgError
                Write-Host $MsgPressExit -ForegroundColor Gray
                $null = Read-Host
                exit
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
            Write-Host ""
            Write-Host $MsgPressExit -ForegroundColor Gray
            $null = Read-Host
            continue MenuPrincipal
        }
    }
}
