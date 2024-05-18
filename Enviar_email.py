import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import main

emails = main.emails
matricula = main.mat
def enviar_email():
    i = 0
    for email1 in main.emails:
        corpo_email = "alguma coisa para teste"
        subject = "Teste1234"
        from_email = 'teste5@gmail.com'
        password = 'password'
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        attachment_paths = [f'certificados/certificado_{matricula[i]}.jpg', f'certificados/certificado_{matricula[i]}.pdf']  # Insira os caminhos dos arquivos aqui

        # Cria a mensagem MIMEMultipart
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = email1[i]
        msg['Subject'] = subject

        # Anexa o corpo do email
        msg.attach(MIMEText(corpo_email, 'plain'))

        # Função para anexar arquivos
        def anexar_arquivo(msg, file_path):
            with open(file_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={file_path.split("/")[-1]}')
            msg.attach(part)

        # Anexar cada arquivo especificado
        for file_path in attachment_paths:
            anexar_arquivo(msg, file_path)

        # Conecta ao servidor SMTP
        s = smtplib.SMTP(smtp_server, smtp_port)
        s.starttls()
        s.login(from_email, password)
        s.sendmail(from_email, email1, msg.as_string())
        s.quit()
        print('Email enviado para', email1)
        i = i+1

enviar_email()
