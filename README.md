*Read this in other languages: [🇧🇷 Português](#-instalador-automático-e-inteligente---office-ltsc) | [🇺🇸 English](#-smart-automatic-installer---office-ltsc)*

---

# 🇧🇷 Instalador Automático e Inteligente - Office LTSC

Este repositório contém um script avançado em PowerShell projetado para automatizar totalmente o processo de download, personalização, instalação e ativação do Microsoft Office LTSC (2024).

Esqueça tutoriais complexos ou arquivos de configuração manuais. O script faz todo o trabalho pesado por você, desde identificar a arquitetura e o idioma do seu sistema até realizar a ativação permanente e silenciosa.

## ✨ Principais Recursos

* **🌐 Multi-idioma Automático:** O script detecta silenciosamente o idioma do seu Windows (Português ou Inglês) e adapta toda a interface do terminal, além de baixar o Office no idioma nativo do sistema (`MatchOS`).
* **🧠 Criação Dinâmica de XML:** Você não precisa de arquivos `config.xml` pré-prontos. O script gera o arquivo de configuração na hora.
* **🎯 Instalação Sob Medida:** O menu interativo permite que você escolha *exatamente* quais aplicativos quer instalar (Word, Excel, PowerPoint, etc.), ignorando o que você não usa.
* **⚙️ Detecção Automática de Arquitetura:** O script identifica sozinho se o seu Windows é de 32 ou 64 bits e adapta a instalação.
* **⚡ Ativação Zero-Touch:** Graças à integração direta com o Microsoft Activation Scripts (MAS), o Office é ativado permanentemente via método Ohook em segundo plano.
* **🧹 100% Limpo:** Todos os downloads ocorrem em uma pasta temporária (`C:\Office_Temp`) que é deletada automaticamente assim que o processo termina.

## 💻 Como Usar (Comando Único)

Você **não precisa baixar** nenhum arquivo manualmente. Basta executar um único comando no seu computador e seguir o menu de seleção de aplicativos na tela.

1. Clique com o botão direito no menu Iniciar e abra o **Windows PowerShell (Administrador)** ou **Terminal (Administrador)**.
2. Copie o comando exato abaixo e pressione **Enter**:

```powershell
irm [https://raw.githubusercontent.com/andreykaua90-blip/Instalador-Office/main/AutoSetup.ps1](https://raw.githubusercontent.com/andreykaua90-blip/Instalador-Office/main/AutoSetup.ps1) | iex
