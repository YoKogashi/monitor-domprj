import requests
import pandas as pd
import smtplib
import os
import fitz  # PyMuPDF
from google import genai
from email.message import EmailMessage
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" 
SENHA_APP = "saty tgmz rzrz yrai" 
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

URL_SITE = "https://www.mprj.mp.br/busca?p_p_id=br_mp_mprj_internet_busca_web_BuscaPortlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&_br_mp_mprj_internet_busca_web_BuscaPortlet_filtro_param=doerj"

def extrair_dados_com_ia(texto_bruto):
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        prompt = f"""
        Analise o texto abaixo e localize a seção "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA".
        Extraia as vagas listadas no formato: Item;Órgão;Critério;Origem da Vaga
        Retorne APENAS os dados. Se não houver, escreva VAZIO.
        Texto: {texto_bruto}
        """
        response = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
        res = response.text.strip()
        if "VAZIO" in res or ";" not in res: return []
        return [linha.split(';') for linha in res.split('\n') if ';' in linha]
    except: return []

def formatar_excel(dados, arquivo, data_do):
    df = pd.DataFrame(dados, columns=["Item", "Órgão", "Critério", "Origem da Vaga (Decorrente de)"])
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='Vagas')
        ws = writer.sheets['Vagas']
        ws.merge_cells('A1:D1')
        ws['A1'] = f"Resultados encontrados no DOeMPRJ de {data_do}"
        ws['A1'].font = Font(size=14, bold=True, color="2F5597")
        ws['A1'].alignment = Alignment(horizontal='center')
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 55

def enviar_email(data_hoje, arquivo_excel=None, arquivo_pdf=None):
    msg = EmailMessage()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    
    if arquivo_excel:
        msg['Subject'] = f'📊 Vagas de Remoção MPRJ - {data_hoje}'
        msg.set_content(f'Olá Renan,\n\nA pesquisa identificou vagas de remoção. Seguem em anexo a planilha e o PDF original do Diário Oficial.')
        # Anexa o Excel
        with open(arquivo_excel, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=arquivo_excel)
        # Anexa o PDF Original
        if arquivo_pdf:
            with open(arquivo_pdf, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=f"Diario_Oficial_{data_hoje.replace('/','-')}.pdf")
    else:
        msg['Subject'] = f'🔍 Monitoramento MPRJ - Varredura {data_hoje}'
        msg.set_content(f'Olá Renan,\n\nPesquisa concluída. Nenhum edital de "Concurso de Remoção" foi localizado hoje.')

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)

def rodar():
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    try:
        response = requests.get(URL_SITE, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')
        link_pdf = ""
        for a in soup.find_all('a', href=True):
            if 'pdf' in a['href'].lower():
                link_pdf = a['href'] if a['href'].startswith('http') else "https://www.mprj.mp.br" + a['href']
                break

        if not link_pdf:
            enviar_email(data_hoje)
            return

        # Baixa o PDF para leitura e anexo
        pdf_content = requests.get(link_pdf).content
        pdf_local = "diario_oficial.pdf"
        with open(pdf_local, "wb") as f:
            f.write(pdf_content)

        doc = fitz.open(pdf_local)
        texto_pdf = "".join([pag.get_text() for pag in doc])
        dados = extrair_dados_com_ia(texto_pdf)

        if dados:
            excel_local = "Vagas_Remocao.xlsx"
            formatar_excel(dados, excel_local, data_hoje)
            enviar_email(data_hoje, arquivo_excel=excel_local, arquivo_pdf=pdf_local)
        else:
            enviar_email(data_hoje)
            
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    rodar()
