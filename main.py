import os
import smtplib
import pyodbc
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from dotenv import load_dotenv

def registrar_log(mensagem):
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    texto = f"[{data_hora}] {mensagem}"
    print(texto)
    with open("registro_robo.txt", "a", encoding="utf-8") as arquivo:
        arquivo.write(texto + "\n")

load_dotenv()

def automacao_contratos_tlog():
    registrar_log("--- INICIANDO AUTOMATIZAÇÃO ---")
    
    hoje = datetime.now()
    data_alvo = hoje + timedelta(days=30)
    data_sql = data_alvo.strftime('%Y-%m-%d')
    data_br = data_alvo.strftime('%d/%m/%Y')

    registrar_log(f"Data de referência: {data_br}")

    contratos = []
    try:
        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};"
            f"SERVER={os.getenv('DB_SERVER')};"
            f"DATABASE={os.getenv('DB_DATABASE')};"
            f"UID={os.getenv('DB_USERNAME')};"
            f"PWD={os.getenv('DB_PASSWORD')};"
        )
        cursor = conn.cursor()
        
        query = f"""
            SELECT FILIAL, TERMINO
            FROM CO_CONTRATO 
            WHERE FILIAL IN (1, 4) 
            AND CAST(TERMINO AS DATE) = '{data_sql}'
        """
        cursor.execute(query)
        contratos = cursor.fetchall()
        conn.close()
        
    except Exception as e:
        registrar_log(f"ERRO NO BANCO: {e}")
        return

    if not contratos:
        registrar_log("Nenhum contrato vence em 30 dias.")
        return

    registrar_log(f"Encontrei {len(contratos)} contrato(s). Preparando envio...")

    try:
        email_remetente = os.getenv('EMAIL_ADDRESS')
        senha_remetente = os.getenv('EMAIL_PASSWORD')
        
        lista_destinatarios = [
            'suporte@grupotlog.com.br'
        ]

        lista_html = "<ul>"
        for c in contratos:
            nome_filial = "PATIO SJP" if c.FILIAL == 1 else "PATIO PNG"
            lista_html += f"<li><b>Filial:</b> {nome_filial} | <b>Vencimento:</b> {data_br}</li>"
        lista_html += "</ul>"

        corpo_email = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h3 style="color: #c0392b;">⚠️ Alerta de Vencimento de Contrato</h3>
            <p>Olá, equipe.</p>
            <p>O sistema identificou que os seguintes contratos vencem em <b>30 dias</b> ({data_br}):</p>
            {lista_html}
            <p>Favor iniciar o processo de renovação ou baixa.</p>
            <hr>
            <p style="font-size: 12px; color: gray;"><i>Mensagem automática - Sistema de Automação TLOG</i></p>
        </body>
        </html>
        """

        msg = MIMEMultipart()
        msg['From'] = email_remetente
        msg['To'] = ", ".join(lista_destinatarios)
        msg['Subject'] = f"ALERTA: Contratos Vencendo em {data_br}"
        msg.attach(MIMEText(corpo_email, 'html'))

        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, lista_destinatarios, msg.as_string())
        server.quit()
        
        registrar_log("E-mail enviado com sucesso para a equipe!")
        
    except Exception as e:
        registrar_log(f"ERRO AO ENVIAR E-MAIL: {e}")

if __name__ == "__main__":
    automacao_contratos_tlog()