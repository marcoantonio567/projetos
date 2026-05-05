# Certificate Automation with Python / Automação de Certificados com Python

## English

### Project Idea

I have an Excel spreadsheet with student data from those who completed a course. I want to explore the possibility of creating a Python program to automate the process of taking data from the spreadsheet and filling in the variable fields on a standard certificate template.

**Fields to be filled automatically:**

- Course name
- Participant name
- Type of participation
- Start date
- End date
- Workload (hours)
- Certificate issue date
- Signatures of:
  - General Manager
  - Coordinator
  - Student

> **Important Note on Work Ethics:**  
> The spreadsheet used in this project contains **fictitious data only**, not real information. This is to ensure ethical handling of data during development and testing.

### Possible Tools and Libraries (Python)

- `pandas` – to read Excel files
- `python-docx` or `reportlab` – to edit Word/PDF certificates
- `Pillow` – for image processing (if certificates are images with signature fields)
- `PyPDF2` or `fpdf2` – for PDF manipulation

---

## Português

### Ideia do Projeto

Tenho uma planilha no Excel com os dados dos alunos que fizeram um curso. Quero ver a possibilidade de criar um programa usando Python para automatizar o envio dos dados da planilha, preenchendo os campos mutáveis em um modelo padrão de certificado.

**Campos a serem preenchidos automaticamente:**

- Nome do curso
- Nome do participante
- Tipo de participação
- Data de início
- Data de término
- Carga horária
- Data de emissão do certificado
- Assinaturas do:
  - Gestor Geral
  - Coordenador
  - Aluno

> **Atenção – Ética de Trabalho:**  
> A planilha utilizada neste projeto contém **apenas dados fictícios**, não informações reais. Isso garante o tratamento ético dos dados durante o desenvolvimento e os testes.

### Possíveis Ferramentas e Bibliotecas (Python)

- `pandas` – para ler arquivos Excel
- `python-docx` ou `reportlab` – para editar certificados em Word/PDF
- `Pillow` – para processamento de imagens (caso os certificados tenham campos de assinatura em imagens)
- `PyPDF2` ou `fpdf2` – para manipulação de PDFs

---

## Proposed Workflow (English) / Fluxo de Trabalho Proposto (Português)

| Step | English | Português |
|------|---------|-----------|
| 1 | Load fictitious data from Excel file using `pandas` | Carregar dados fictícios do arquivo Excel usando `pandas` |
| 2 | Read the standard certificate template (e.g., DOCX or PDF) | Ler o modelo padrão do certificado (ex.: DOCX ou PDF) |
| 3 | Replace variable fields with student data | Substituir os campos variáveis pelos dados do aluno |
| 4 | Insert signatures (images or text) | Inserir assinaturas (imagens ou texto) |
| 5 | Generate final certificate for each student | Gerar o certificado final para cada aluno |
| 6 | Save or export automatically | Salvar ou exportar automaticamente |

---

## Ethical Reminder (English/Português)

> **All student data used in this project is fictitious and generated solely for testing and learning purposes. No real personal information is stored or processed.**  
> **Todos os dados de alunos usados neste projeto são fictícios e gerados exclusivamente para fins de teste e aprendizado. Nenhuma informação pessoal real é armazenada ou processada.**

---

## Sample Fictitious Data / Exemplo de Dados Fictícios

| Nome Participante | Curso               | Participação | Início     | Fim        | Carga Horária |
|-------------------|---------------------|--------------|------------|------------|---------------|
| Ana Silva         | Python Básico       | Remota       | 01/03/2025 | 15/03/2025 | 20h           |
| João Souza        | Excel Avançado      | Presencial   | 10/04/2025 | 25/04/2025 | 30h           |
| Maria Oliveira    | Gestão de Projetos  | Híbrida      | 05/05/2025 | 20/05/2025 | 24h           |
