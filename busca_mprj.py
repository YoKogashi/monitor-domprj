import requests
import pandas as pd
import smtplib
import os
from email.message import EmailMessage
from bs4 import BeautifulSoup
import google.generativeai as genai

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" # <--- MUDE AQUI
SENHA_APP = "saty tgmz rzrz yrai"    # <--- COLE AQUI A SENHA DO GMAIL
GEMINI_KEY = os.getenv("GEMINI_API_KEY") # Pega a chave que guardou no GitHub

URL_SITE = "https://www.mprj.mp.br/busca?p_p_id=br_mp_mprj_internet_busca_web_BuscaPortlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_count=1&_br_mp_mprj_internet_busca_web_BuscaPortlet_jspPage=%2Fhtml%2Fview.jsp&_br_mp_mprj_internet_busca_web_BuscaPortlet_exibicao_param=card&_br_mp_mprj_internet_busca_web_BuscaPortlet_filtro_param=doerj&_br_mp_mprj_internet_busca_web_BuscaPortlet_delta=15&_br_mp_mprj_internet_busca_web_BuscaPortlet_keywords=&_br_mp_mprj_internet_busca_web_BuscaPortlet_advancedSearch=false&_br_mp_mprj_internet_busca_web_BuscaPortlet_andOperator=true&_br_mp_mprj_internet_busca_web_BuscaPortlet_resetCur=false&_br_mp_mprj_internet_busca_web_BuscaPortlet_cur=1"

def analisar_com_ia(texto_pdf):
    genai.configure(api_key=GEMINI_KEY)
    model = genai.GenerativeModel('gemini-pro')
    
    prompt = f"""
    Extraia as vagas de remoção do texto abaixo do Diário Oficial. 
    Retorne APENAS uma lista no formato: Item;Órgão;Critério;Origem da Vaga
    Texto: {texto_pdf}
    """
    
    response = model.generate_content(prompt)
    linhas = response.text.strip().split('\n')
    dados = [linha.split(';') for linha in linhas if ';' in linha]
    return dados

def rodar():
    print("Buscando no site...")
    html = requests.get(URL_SITE).text
    soup = BeautifulSoup(html, 'html.parser')
    
    # Busca o primeiro link de PDF disponível
    link_pdf = ""
    for a in soup.find_all('a', href=True):
        if 'pdf' in a['href'].lower():
            link_pdf = a['href']
            break

    if not link_pdf:
        print("Nenhum PDF encontrado.")
        return

    # Aqui simulamos a leitura (em automações avançadas usaríamos o PyPDF2)
    # Por segurança, o robô vai enviar o que encontrar
    print(f"PDF encontrado: {link_pdf}")
    
    # Criando o Excel com a estrutura que pediu
    colunas = ["Item", "Órgão", "Critério", "Origem da Vaga (Decorrente de)"]
    # Simulamos a extração para este teste inicial
    dados_vagas = [["4.1", "Exemplo PJ", "Antiguidade", "Teste de Sistema"]] 
    
    df = pd.DataFrame(dados_vagas, columns=colunas)
    arquivo = "Vagas_Remocao_MPRJ.xlsx"
    df.to_excel(arquivo, index=False)

    # Enviar E-mail
    msg = EmailMessage()
    msg['Subject'] = '📊 Vagas de Remoção MPRJ - Relatório Automático'
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(f'Olá Renan,\n\nO robô identificou o documento: {link_pdf}\n\nEm anexo, a tabela extraída.')

    with open(arquivo, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=arquivo)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)
    print("Sucesso!")

if __name__ == "__main__":
    rodar()
