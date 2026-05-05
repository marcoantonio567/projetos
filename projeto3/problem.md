# Bot Assistant for Accounting Entries / Assistente de Lançamentos Contábeis

## English

### Problem Description

I need to have an assistant bot that performs my company's accounting entry functions, as the system itself is very tedious. I would like to make the entries in Excel, then deliver the ready Excel file to the bot so it can post the entries while I do other activities.

I tried to create one here. PyAutoGUI works perfectly, but when trying to read an .xlsx file, pylint indicates that the line of code is wrong. As a result, when executing the parameters, it gets stuck when trying to post the data.

**I need urgent help.**

> **Note:** Due to client privacy concerns, the spreadsheet data has been replaced with randomly generated data. I tried to recreate the company's system layout to make it as similar as possible.

---

## Português

### Descrição do Problema

Preciso ter um bot assistente que faça minhas funções de lançamento contábil da empresa que trabalho, pois o próprio sistema é muito maçante. Gostaria de fazer os lançamentos no Excel, entregar ao bot o Excel pronto para o bot ir lançando enquanto faço outras atividades.

Tentei criar um aqui. O PyAutoGUI funciona perfeitamente, mas na hora de ler o arquivo .xlsx ele avisa pelo pylint que está errado a linha de código. Por isso, na hora de executar os parâmetros, ele trava na hora de lançar com os dados.

**Preciso de ajuda urgentemente.**

> **Observação:** Por questão de respeito à privacidade do cliente, os dados da planilha foram trocados por dados aleatórios criados. Tentei recriar o layout do sistema da empresa para ficar o mais parecido possível.

---

## Technical Notes / Notas Técnicas

- **PyAutoGUI** - Working perfectly / Funciona perfeitamente
- **Excel (.xlsx) reading** - Error reported by pylint / Erro relatado pelo pylint
- **Issue** - Process stops when trying to post data / Processo trava ao tentar lançar dados
