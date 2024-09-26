# README

## Descrição

Este script em Python utiliza as bibliotecas `openpyxl` e `win32com` para automatizar o envio de e-mails através do Microsoft Outlook, com base em uma planilha Excel contendo informações de nomes e e-mails. Ele também permite o envio de anexos personalizados para cada destinatário.

## Requisitos

### Softwares necessários:

- **Microsoft Outlook**: O script depende do Outlook para criar e enviar e-mails.
- **Excel**: As informações dos destinatários e os arquivos anexados estão organizados em uma planilha Excel.

### Bibliotecas Python:

1. **openpyxl**: Utilizada para carregar e ler os dados da planilha.
   - Instalação: 
     ```bash
     pip install openpyxl
     ```
2. **pywin32 (win32com.client)**: Utilizada para interação com o Microsoft Outlook.
   - Instalação:
     ```bash
     pip install pywin32
     ```

### Estrutura do Projeto

- `Lista de Email.xlsx`: Planilha Excel que contém as informações dos destinatários (nome, nome completo e e-mail) na aba "Dados".
- Diretório `C:\Users\fgpaula\Desktop\automatização\HORARIO INSTRUTORES\SET 2\`: Contém os arquivos Excel anexados a cada e-mail.
- Diretório `C:\Assinatura\`: Contém a imagem que será usada como assinatura nos e-mails.

## Funcionamento

1. **Carregamento da planilha Excel**: O script carrega a planilha `Lista de Email.xlsx` que contém os dados dos destinatários.
   
2. **Iteração sobre as linhas da planilha**: O código percorre todas as linhas da aba "Dados", a partir da segunda linha, onde as informações dos destinatários estão organizadas.

3. **Criação do e-mail no Outlook**:
   - Para cada destinatário, o script cria um novo e-mail no Outlook com o nome e o e-mail correspondente.
   - O campo de assunto é composto pelo nome completo do destinatário.
   - O corpo do e-mail é em HTML e contém o nome do destinatário de forma personalizada.

4. **Anexo de arquivo**: 
   - O script tenta anexar um arquivo específico localizado em `C:\Users\fgpaula\Desktop\automatização\HORARIO INSTRUTORES\SET 2\`. O nome do arquivo é baseado no nome completo do destinatário.
   - Se o arquivo não for encontrado, uma mensagem de erro é exibida no console.

5. **Salvar e-mail**: Ao final do processo, o e-mail é salvo no Outlook (não enviado automaticamente).

## Estrutura da Planilha Excel

A planilha `Lista de Email.xlsx` deve conter os seguintes dados na aba "Dados":

| Coluna A (Nome) | Coluna B (Nome Completo) | Coluna C (E-mail) |
|-----------------|--------------------------|-------------------|
| João            | João da Silva             | joao@email.com    |
| Maria           | Maria Oliveira            | maria@email.com   |

## Como Executar

1. Certifique-se de que a planilha Excel e os arquivos anexos estão nos diretórios corretos.
2. Execute o script Python.
3. Verifique o Microsoft Outlook para os e-mails criados, prontos para envio.

## Observações

- **Envio Manual**: O script salva os e-mails, mas não os envia automaticamente. Para enviar, você pode abrir os e-mails salvos no Outlook e clicar em "Enviar".
- **Verificação de anexo**: O script faz uma verificação para garantir que o arquivo a ser anexado existe no caminho especificado.


## Possíveis Melhorias

- Implementar o envio automático de e-mails utilizando `emailOutlook.Send()`.
- Adicionar tratamento de erros mais robusto, por exemplo, no caso de problemas com a conexão ao Outlook ou falhas na leitura da planilha.

