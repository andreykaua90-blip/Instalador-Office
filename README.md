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
