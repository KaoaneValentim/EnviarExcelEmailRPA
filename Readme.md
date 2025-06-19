
#  Requisitos para Executar o Projeto UiPath

Para que este projeto funcione corretamente, você deve configurar os seguintes itens:

---

## 1. Planilha Excel (`Funcionarios.xlsx`)

- Crie uma planilha com **dados tabulares** (ex: nomes, datas, valores etc).
- Salve como `Funcionarios.xlsx`.

 **Local esperado**: mesma pasta onde o projeto `.xaml` está localizado.

Ou, altere no código:
```vb
workbookPath := "C:\caminho\completo\Funcionarios.xlsx"
```

---

##  2. Documento Word (Modelo Base)

- Crie um documento `.docx` que servirá como **base para o relatório**.
- Ele será preenchido com a tabela do Excel e convertido para PDF.

📍 Altere a propriedade `FilePath` na atividade `Use Word File`:
```vb
FilePath := "C:\caminho\para\modelo.docx"
```

---

##  3. Destinatário do E-mail

- Defina o campo `To` com o e-mail do destinatário na atividade `Send Email`.

📍 Exemplo:
```vb
To := "exemplo@dominio.com"
Subject := "Relatório gerado automaticamente"
```

---

##  4. Outlook Instalado

- É necessário ter o **Microsoft Outlook instalado e configurado** no mesmo perfil do Windows que executa o robô.
- O projeto usa `Use Desktop Outlook App`, que depende do Outlook local.

---

##  5. (Opcional) Corpo do E-mail HTML

- O corpo do e-mail é carregado de um arquivo externo em HTML.

 Local esperado:
```
.data/HtmlContent0.html
```

🛠️ Você pode editar esse arquivo ou substituir o conteúdo HTML pela propriedade `Body` com texto simples.

---


