import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#Entre com o e-mail a ser usado para enviar
email_user = input("Entre com o e-mail: ")
email_password = input("Entre com senha do e-mail: ")
print("\n")

#Lista de e-mails fica em um arquivo 'email.txt' na mesma pasta do script
arquivo = open('email.txt', 'r')
emails = []

'''
Os arquivos precisam ser numerados de 1.pdf a 1000.pdf, por exemplo.
a ordem dos e-mails no .txt é um de/para os arquivos .pdf.
A quantidade de e-mails e PDF's devem ser a mesma
'''

for linha in arquivo:
    linha = linha.strip()
    emails.append(linha)

arquivo.close()

for x in range(len(emails)):

    email_send = emails[x]
    subject = 'Informe de rendimentos teste'
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject

    body = 'Segue seu informe de rendimentos como teste'
    msg.attach(MIMEText(body,'plain'))

    filename="" + str(x + 1) + '.pdf' + ""
    attachment=open(filename,'rb')

    part = MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment; filename= "+filename)

    msg.attach(part)
    text = msg.as_string()
    #Servidor SMTP pode ser alterado de acordo com o e-mail
    server = smtplib.SMTP('smtp.office365.com',587)
    server.starttls()
    server.login(email_user,email_password)

    server.sendmail(email_user,email_send,text)
    server.quit()
    print("Email enviado para %s com sucesso" % emails[x])
    print("Preparando para enviar o próximo")
    time.sleep(5)
    print("\n")
    attachment.close()


print("Todos os e-mails enviados com sucesso")
