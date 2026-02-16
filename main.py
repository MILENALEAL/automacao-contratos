import os
import smtplib
import pyodbc
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

def registrar_auto(mensagem):
    agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    linha_log = f"[{agora}] {mensagem}\n"
    with open("registro_auto.txt", "a", encoding="utf-8") as arquivo:
        arquivo.write(linha_log)

def enviar_email_vencimento(nome_cliente, numero_contrato, data_vencimento, dias_restantes):
    email_remetente = os.getenv('EMAIL_ADDRESS')
    senha_remetente = os.getenv('EMAIL_PASSWORD')
    destinatarios = [
    'evelise.batistela@grupotlog.com.br',
    'andre.frigotto@grupotlog.com.br',
    'viviane.suider@grupotlog.com.br'
]

    corpo_html = f"""
    <html>
        <body style="font-family: Arial, sans-serif; background-color: #f4f4f4; padding: 20px;">
            <div style="max-width: 600px; margin: auto; background-color: #ffffff; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
                <div style="background-color: #0056b3; color: white; padding: 30px; text-align: center; font-size: 26px; font-weight: bold;">
                    Notificação de Contratos Vencendo
                </div>
                <div style="padding: 30px; color: #333; line-height: 1.6; font-size: 16px;">
                    <p>Olá,</p>
                    <p>O contrato nº <span style="color: #007bff; font-weight: bold;">{numero_contrato}</span> irá vencer em {dias_restantes} dias.</p>
                    <p style="margin-top: 25px;">
                        <b>Cliente:</b> {nome_cliente}<br>
                        <b>Data de Vencimento:</b> {data_vencimento}
                    </p>
                </div>
                <div style="background-color: #fafafa; padding: 15px; text-align: center; font-size: 12px; color: #999; border-top: 1px solid #eee;">
                    Este é um aviso automático.
                </div>
            </div>
        </body>
    </html>
    """

    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = f"ALERTA: Contrato Vencendo - {nome_cliente}"
    msg.attach(MIMEText(corpo_html, 'html'))

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, destinatarios, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        registrar_auto(f"ERRO AO ENVIAR E-MAIL ({numero_contrato}): {e}")
        return False

def automacao_contratos_tlog():
    registrar_auto("--- INICIANDO AUTOMATIZAÇÃO ---")
    hoje = datetime.now()
    dias_antecedencia = 30
    data_alvo = hoje + timedelta(days=dias_antecedencia)
    data_sql = data_alvo.strftime('%Y-%m-%d')
    data_br = data_alvo.strftime('%d/%m/%Y')

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
            SELECT C.NUMERO, P.NOME, C.TERMINO, C.FILIAL
            FROM CO_CONTRATO C
            INNER JOIN MS_PESSOA P ON P.HANDLE = C.PESSOA
            WHERE C.FILIAL IN (1, 4) 
            AND C.TIPO = 10
            AND CAST(C.TERMINO AS DATE) = '{data_sql}'
        """
        cursor.execute(query)
        contratos = cursor.fetchall()
        conn.close()
    except Exception as e:
        registrar_auto(f"ERRO NO BANCO DE DADOS: {e}")
        return

    if not contratos:
        registrar_auto(f"Nenhum contrato vence em {data_br}.")
        return

    for c in contratos:
        num_contrato = c[0]
        nome_cliente = c[1]
        data_venc = c[2].strftime('%d/%m/%Y') if hasattr(c[2], 'strftime') else str(c[2])
        sucesso = enviar_email_vencimento(nome_cliente, num_contrato, data_venc, dias_antecedencia)
        if sucesso:
            registrar_auto(f"E-mail enviado: Contrato {num_contrato} - {nome_cliente}")

if __name__ == "__main__":
    automacao_contratos_tlog()
