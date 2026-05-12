import requests
import pandas as pd
import smtplib
from email.message import EmailMessage
from bs4 import BeautifulSoup
import os

# CONFIGURAÇÕES DO ROBÔ
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" # <--- MUDE AQUI
SENHA_APP = "saty tgmz rzrz yrai"    # <--- COLE AQUI A SENHA DO PASSO 0

URL_BUSCA = "https://www.mprj.mp.br/busca?p_p_id=br_mp_mprj_internet_busca_web_BuscaPortlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_count=1&_br_mp_mprj_internet_busca_web_BuscaPortlet_jspPage=%2Fhtml%2Fview.jsp&_br_mp_mprj_internet_busca_web_BuscaPortlet_exibicao_param=card&_br_mp_mprj_internet_busca_web_BuscaPortlet_filtro_param=doerj&_br_mp_mprj_internet_busca_web_BuscaPortlet_delta=15&_br_mp_mprj_internet_busca_web_BuscaPortlet_keywords=&_br_mp_mprj_internet_busca_web_BuscaPortlet_advancedSearch=false&_br_mp_mprj_internet_busca_web_BuscaPortlet_andOperator=true&_br_mp_mprj_internet_busca_web_BuscaPortlet_resetCur=false&_br_mp_mprj_internet_busca_web_BuscaPortlet_cur=1"

def enviar_email(arquivo_excel):
    msg = EmailMessage()
    msg['Subject'] = '📊 Vagas de Remoção MPRJ - Identificadas pelo Robô'
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content('Olá Renan,\n\nO robô identificou uma nova edição do Diário Oficial. Segue em anexo a tabela com as vagas extraídas.')

    with open(arquivo_excel, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=arquivo_excel)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)
    print("E-mail enviado com sucesso!")

def rodar():
    print("Buscando no site...")
    # (Aqui o robô faria a lógica de ler o PDF e extrair os dados automaticamente)
    # Por agora, para testar se o seu e-mail funciona, ele vai gerar um Excel de teste
    
    dados_exemplo = [
        {"Item": "4.1", "Órgão": "Exemplo", "Critério": "Teste", "Origem": "Robô Funcionando"}
    ]
    df = pd.DataFrame(dados_exemplo)
    nome_arquivo = "Vagas_Remocao.xlsx"
    df.to_excel(nome_arquivo, index=False)
    
    enviar_email(nome_arquivo)

if __name__ == "__main__":
    rodar()
