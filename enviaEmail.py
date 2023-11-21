# pip install secure-smtplib
import smtplib
import datetime
from email.mime.text import MIMEText

def EnviaEmail(mensagem):
    remetente = 'emailRemetente'
    destinatarios = ['listaComEmailsDosDestinatarios']
    assunto = 'Teste Robô Python'
    
    # Formatando a data e hora
    data_hora_atualizacao = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    
    # Corpo do e-mail
    corpo_email = f'{mensagem} Data e hora da atualização: {data_hora_atualizacao}'
    
    # Configuração do e-mail
    msg = MIMEText(corpo_email, 'plain', 'utf-8')
    msg['Subject'] = assunto
    msg['From'] = remetente
    msg['To'] = ', '.join(destinatarios)

    try:
        # Configuração do servidor SMTP, padrao: Outlook
        servidor_email = smtplib.SMTP('SMTP.office365.com', 587)
        servidor_email.ehlo()
        servidor_email.starttls()
              
        servidor_email.login(remetente, 'senhaRemetente')

        # Envio do e-mail
        servidor_email.sendmail(remetente, destinatarios, msg.as_string())
        print("E-mail enviado com sucesso!")

    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")

    finally:
        # Encerramento da conexão SMTP, mesmo em caso de erro
        servidor_email.quit()

