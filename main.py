#IMPORTANDO A TABELA COM NOME, EMAIL E NOME DO ARQUIVO
import pandas as pd
tb_frequencia = pd.read_excel('lista_destinatarios.xlsx')

#ABRINDO ARQUIVO TXT QUE TEM O CORPO DA MENSAGEM
txt_email = open('corpo_email.txt', encoding="utf8")
txt_email = txt_email.readlines()
txt_assunto_email = txt_email[0]
impofor i in txt_assunto_email:
  cont = cont + 1
txt_assunto_email = txt_assunto_email[:62]
txt_corpo_email = ''.join(txt_email[2:])

#DEFINDO FUNÇÃO PARA ENVIAR E-MAIL
def enviar_email(destinatario, anexo):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    remetente = 'SEU-ENDEREÇO-EMAIL'
    password = 'SUA-SENHA'

    msg = MIMEMultipart()

    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f'{txt_assunto_email}'
    msg.attach(MIMEText(f'{txt_corpo_email}', 'plain'))

    # abrindo e convertendo arquivo de anexo
    attachment = open(anexo, 'rb')
    part = MIMEBase('application', 'pdf')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)

    part.add_header('Content-Disposition', 'attachment', filename=anexo)

    # inserindo anexo na mensagem
    msg.attach(part)

    # encerrando arquivo
    attachment.close()

    servidor = smtplib.SMTP('smtp.gmail.com:587')
    servidor.starttls()
    servidor.login(remetente, password)
    text = msg.as_string()
    servidor.sendmail(remetente, destinatario, text)

    print(anexo)
    print('Email enviado')

#CRIANDO VETOR DE REFERÊNCIA DE QUANTIDADE DE DESTINATÁRIOS
destinatarios = tb_frequencia['NOME']

#PERCORRENDO CADA REGISTRO DA TABELA
for nome in destinatarios:
    reg_destinatarios = tb_frequencia.loc[tb_frequencia['NOME'] == nome]

    email = ''.join(reg_destinatarios['E-mail'])
    anexo = ''.join(reg_destinatarios['ANEXO'])

    enviar_email(email, anexo)
