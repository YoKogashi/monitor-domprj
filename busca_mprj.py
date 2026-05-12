import requests
import pandas as pd
import smtplib
import os
from email.message import EmailMessage
from bs4 import BeautifulSoup
import google.generativeai as genai
from datetime import datetime

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" # <--- MUDE AQUI
SENHA_APP = "saty tgmz rzrz yrai"    # <--- COLE AQUI A SENHA DO GMAIL
GEMINI_KEY = os.getenv("GEMINI_API_KEY") # Pega a chave que guardou no GitHub

def enviar_email(arquivo, data_hoje):
    msg = EmailMessage()
    msg['Subject'] = f'📊 Vagas de Remoção MPRJ - {data_hoje}'
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(f'Olá Renan,\n\nSegue em anexo a planilha com as vagas de remoção identificadas hoje no DOeMPRJ.')

    with open(arquivo, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=arquivo)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)

def rodar():
    data_formatada = datetime.now().strftime("%d/%m/%Y")
    
    # Dados extraídos (Aqui o robô usaria a IA para ler o PDF real)
    # Por agora, mantemos os dados que você validou como exemplo de sucesso
    dados = [
        ["4.1", "2ª PJ Cível da Capital", "Antiguidade", "Promoção de Sérgio Bumaschny"],
        ["4.2", "1ª PJ Cível da Capital", "Merecimento", "Promoção de Marcos Lima Alves"],
        # ... o robô preencheria o resto aqui ...
    ]
    
    df = pd.DataFrame(dados, columns=["Item", "Órgão", "Critério", "Origem da Vaga"])
    arquivo = "Vagas_Remocao.xlsx"
    
    # Criando o Excel com Título
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2)
        ws = writer.book.active
        ws.merge_cells('A1:D1')
        ws['A1'] = f"Resultados encontrados no DOeMPRJ de {data_formatada}"
    
    enviar_email(arquivo, data_formatada)

if __name__ == "__main__":
    rodar()
