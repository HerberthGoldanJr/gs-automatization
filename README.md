# 📦 Automação de Planilha de Carregamento Logístico (Google Apps Script)

Este projeto é um script desenvolvido em **Google Apps Script**, usando .JS para automatizar o preenchimento e envio de uma **planilha de carregamento logístico**, enviado por email.

## 🚀 Funcionalidades

- Copia automaticamente os dados relevantes de uma planilha central com todos os processos pendentes.
- Preenche uma planilha secundária (planilha de carregamento) com os dados do processo em destaque.
- Formata os campos conforme o tipo de modal, recinto e origem do processo.
- Cria automaticamente um **rascunho no Gmail**, pronto para envio ao responsável pelo carregamento logístico.
- Inclui regras para anexar documentos específicos (PDFs) com base no tipo de processo (aéreo, marítimo, etc).

## 🛠️ Tecnologias Utilizadas

- Google Apps Script
- Google Sheets
- Google Drive (para localizar e anexar arquivos PDF)
- Gmail API (para criar rascunhos de e-mail)

## 📁 Estrutura do Projeto

- `aereo_maritimo_MPR.gs` — Arquivo principal com as funções de automação.
- `appsscript.json` — Configuração do projeto Apps Script.

## 📝 Como Usar

1. Abra o Google Sheets com a planilha de carregamento.
2. Acesse `Extensões > Apps Script`.
3. Cole o código do repositório.
4. Configure os gatilhos conforme necessário (ex: ao editar uma célula com referência).
5. Conceda permissões ao script na primeira execução.
6. Teste com uma referência de processo válida para ver o preenchimento automático e o rascunho gerado.

## 📌 Observações

- Os nomes dos arquivos PDF seguem um padrão específico (ex: `PXXXXXX-XX_DI_REF`).
- Os processos são identificados por uma **referência**, e os dados adicionais são extraídos diretamente dos PDFs armazenados no Google Drive.
- O script foi criado para reduzir o retrabalho manual e garantir padronização no envio de processos logísticos, garatindo minimização de erros e ganho de tempo nas funções.

## 📬 Contato

Caso queira adaptar esse script para sua operação ou tenha dúvidas, entre em contato via herberthgoldanjr.@gmail.com ou abra uma issue aqui no GitHub.

---

# 📦 Logistics Loading Sheet Automation (Google Apps Script)

This project is a script developed in **Google Apps Script** to automate the filling in and sending of a **logistics loading sheet**.

## 🚀 Features

- Automatically copies the relevant data from a central spreadsheet with all pending processes.
- Fills in a secondary spreadsheet (loading sheet) with the data from the highlighted process.
- Formats the fields according to the type of modal, enclosure and origin of the process.
- Automatically creates a **draft in Gmail**, ready to send to the person responsible for the logistics shipment.
- Includes rules for attaching specific documents (PDFs) based on the type of process (air, sea, etc).

## 🛠️ Technologies Used

- Google Apps Script
- Google Sheets
- Google Drive (to locate and attach PDF files)
- Gmail API (to create draft emails)

## 📁 Project Structure

- `Code.gs` - Main file with the automation functions.
- `appsscript.json` - Configuration of the Apps Script project.

## 📝 How to use

1. Open Google Sheets with the upload spreadsheet.
2. Go to `Extensions > Apps Script`.
3. Paste the code from the repository.
4. Configure the triggers as necessary (e.g. when editing a cell with a reference).
5. Grant permissions to the script on the first run.
6. Test with a valid process reference to see the autofill and draft generated.

## 📌 Observations

- PDF file names follow a specific pattern (e.g., PXXXXXX-XX_DI_REF).
- Processes are identified by a reference, and additional data is extracted directly from PDFs stored on Google Drive.
- The script was created to reduce manual rework and ensure standardization in logistics process submissions, guaranteeing error minimization and time savings in operations.

## 📬 Contact

If you want to adapt this script for your operation or have questions, contact via herberthgoldanjr.@gmail.com or open an issue here on GitHub.



