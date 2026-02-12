import os
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()

def testar_conexao():
    email = os.getenv('EMAIL_ADDRESS')
    senha = os.getenv('EMAIL_PASSWORD')
    
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    print("Tentando conectar ao Outlook...")

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls() 
        server.login(email, senha)
        
        msg = MIMEText("DEUUUU CERTOOOOO EEEEEEEEE")
        msg['Subject'] = "Teste de Conexão Python"
        msg['From'] = email
        msg['To'] = email 

        server.sendmail(email, email, msg.as_string())
        server.quit()
        print("SUCESSO: E-mail enviado! Verifique sua caixa de entrada.")
        
    except Exception as e:
        print(f"ERRO: {e}")

testar_conexao()