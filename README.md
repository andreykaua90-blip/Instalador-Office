# AutoSetup - Office LTSC Deployment

PowerShell script developed to automate the configuration, download, installation, and activation processes of Microsoft Office LTSC 2024.

The script operates autonomously (via remote execution) and handles official Microsoft deployment tools, as well as integrating local command-line activation methods.

## Methodology and Tools

* **Office Deployment Tool (ODT):** The script dynamically fetches the latest official Office Deployment Tool directly from the Microsoft Download Center. It scrapes the download page, downloads the most recent `officedeploymenttool_*.exe`, and silently extracts it to obtain `setup.exe`. This ensures the tool is always up to date without relying on a fixed binary stored in the repository.
* **Dynamic XML Generation:** The script eliminates the need for a static `config.xml` file in the repository. The PowerShell code processes user input and generates the XML configuration file at runtime, injecting `<ExcludeApp ID="..." />` tags for the software the user chooses not to install.
* **Environment Detection:**
  * **Architecture:** Uses the `[Environment]::Is64BitOperatingSystem` method to automatically set the `OfficeClientEdition` tag (32 or 64-bit).
  * **Localization (Language):** The `<Language ID="MatchOS" />` parameter instructs the ODT to download Office in the operating system's native language. The script also uses `(Get-UICulture).TwoLetterISOLanguageName` to determine the output language of the terminal strings.
* **Activation (Ohook Method):** Activation is achieved through integration with Microsoft Activation Scripts (MAS). The script calls the URL get.activated.win passing the `/ohook` argument, which overrides the local license validation routines (SPP - Software Protection Platform), resulting in a permanent activation.

## Execution Flow

1. Creation of the temporary working directory (`C:\Office_Temp`).
2. System data collection (Language and Architecture).
3. User input collection via console to define the package scope.
4. Writing the `config.xml` file to the temporary directory.
5. Dynamic download of the latest Office Deployment Tool from the official Microsoft Download Center.
6. Silent extraction of the ODT executable to generate the `setup.exe` binary, followed by the installation initialization via `setup.exe /configure config.xml`.
7. Background execution of the MAS script (Ohook).
8. Recursive deletion of the `C:\Office_Temp` directory to clean the environment.

## Usage

1. Open **Windows PowerShell** or **Windows Terminal** with Administrator privileges.
2. Run the following command:
   ```powershell
   irm https://raw.githubusercontent.com/andreykaua90-blip/Office-AutoSetup/main/AutoSetup.ps1 | iex

---
---

# AutoSetup - Implantação do Office LTSC

Script em PowerShell desenvolvido para a automação dos processos de configuração, download, instalação e ativação do Microsoft Office LTSC 2024.

O script opera de forma autônoma (via execução remota) e manipula as ferramentas oficiais de implantação da Microsoft, além de integrar métodos de ativação locais via linha de comando.

## Metodologia e Ferramentas

* **Office Deployment Tool (ODT):** O script obtém dinamicamente a versão mais recente do Office Deployment Tool diretamente do Microsoft Download Center. Ele faz o scrape da página de download, baixa o `officedeploymenttool_*.exe` mais atual e o extrai silenciosamente para obter o `setup.exe`. Isso garante que a ferramenta esteja sempre atualizada, sem depender de um binário fixo armazenado no repositório.
* **Geração Dinâmica de XML:** O script dispensa o uso de um arquivo `config.xml` estático no repositório. O código PowerShell processa o input do usuário e gera o arquivo XML em tempo de execução, injetando as tags `<ExcludeApp ID="..." />` para os softwares que o usuário optar por não instalar.
* **Detecção de Ambiente:**
  * **Arquitetura:** Utiliza o método `[Environment]::Is64BitOperatingSystem` para definir automaticamente a tag `OfficeClientEdition` (32 ou 64 bits).
  * **Localização (Idioma):** O parâmetro `<Language ID="MatchOS" />` instrui o ODT a baixar o Office no mesmo idioma do sistema operacional. O script também utiliza `(Get-UICulture).TwoLetterISOLanguageName` para definir o idioma das strings de saída do próprio terminal.
* **Ativação (Método Ohook):** A ativação é realizada através da integração com o Microsoft Activation Scripts (MAS). O script aciona a URL get.activated.win passando o argumento `/ohook`, que atua substituindo as rotinas locais de validação de licença (SPP - Software Protection Platform), resultando em uma ativação permanente.

## Fluxo de Execução

1. Criação do diretório temporário de trabalho (`C:\Office_Temp`).
2. Coleta de dados do sistema (Idioma e Arquitetura).
3. Coleta do input do usuário via console para definir o escopo dos pacotes.
4. Escrita do arquivo `config.xml` no diretório temporário.
5. Download dinâmico da versão mais recente do Office Deployment Tool a partir do Microsoft Download Center oficial.
6. Extração silenciosa do executável do ODT para gerar o binário `setup.exe`, seguida da inicialização da instalação via `setup.exe /configure config.xml`.
7. Execução em segundo plano do script do MAS (Ohook).
8. Exclusão recursiva do diretório `C:\Office_Temp` para limpeza do ambiente.

## Uso

1. Abra o **Windows PowerShell** ou o **Terminal do Windows** com privilégios de Administrador.
2. Execute o comando abaixo:
   ```powershell
   irm https://raw.githubusercontent.com/andreykaua90-blip/Office-AutoSetup/main/AutoSetup.ps1 | iex
