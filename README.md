# 🍯 Instalador Automático - Office LTSC

Este repositório contém um script em PowerShell projetado para automatizar totalmente o processo de download, instalação e ativação do Microsoft Office LTSC.

Tudo é feito de forma limpa, utilizando o Office Deployment Tool (ODT) oficial da Microsoft e ativando através do método Ohook, sem deixar arquivos temporários no seu computador.

## 🚀 Como usar (Comando Único)

Você não precisa baixar os arquivos manualmente. Basta executar um único comando para que o script faça todo o trabalho.

1. Clique com o botão direito no menu Iniciar e abra o **Windows PowerShell (Administrador)** ou **Terminal (Administrador)**.
2. Copie e cole o comando abaixo e pressione **Enter**:

```powershell
irm [https://raw.githubusercontent.com/andreykaua90-blip/Instalador-Office/main/Instalar.ps1](https://raw.githubusercontent.com/andreykaua90-blip/Instalador-Office/main/Instalar.ps1) | iex
