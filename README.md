# Envio Automatizado de E-mail em Lote com Arquivo em Anexo
Este projeto facilita o envio em massa de certificados personalizados por email, utilizando Python. É especialmente útil para organizadores de eventos que precisam distribuir certificados de participação para vários destinatários. O script automatiza o processo de envio de emails, anexando os arquivos corretos a cada mensagem.

## Funcionalidades
- Importa uma lista de destinatários e seus respectivos anexos de um arquivo Excel.
- Lê o corpo da mensagem e o assunto do email de um arquivo de texto.
- Envia emails com anexos personalizados para cada destinatário na lista.
- Utiliza o protocolo SMTP para enviar emails através de um servidor.

## Requisitos
- `Pandas`
- `smtplib`
- `email`

## Estrutura do Projeto
- `main.py`: Script principal que executa o envio de emails.
- `lista_destinatarios.xlsx`: Arquivo Excel contendo os destinatários e os nomes dos arquivos anexos.
- `corpo_email.txt`: Arquivo de texto contendo o assunto e o corpo da mensagem do email.

## Uso
1. Prepare o arquivo `lista_destinatarios.xlsx` com a estrutura:

   | NOME     | E-mail           | ANEXO                |
   |----------|------------------|----------------------|
   | João Doe | joao@example.com | certificado_joao.pdf |
   | Maria Doe| maria@example.com| certificado_maria.pdf|
2. Crie o arquivo `corpo_email.txt` com o seguinte formato:
   ```
   Assunto do Email
   (Linha em branco)
   Corpo do email...
   ```
3. Edite o script `main.py` e substitua `'SEU-ENDEREÇO-EMAIL'` e `'SUA-SENHA'` com suas credenciais.
4. Execute o script:
   ```bash
   python main.py
   ```

## Autor
- [@abner-lucas](https://github.com/abner-lucas)
  
## Licença
Este projeto está licenciado sob a licença MIT. Veja o arquivo [MIT](https://choosealicense.com/licenses/mit/) para mais detalhes
