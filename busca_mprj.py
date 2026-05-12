import requests
import pandas as pd
import smtplib
import os
import fitz  # PyMuPDF
from google import genai
from email.message import EmailMessage
from datetime import datetime, timedelta # Adicionado timedelta
from bs4 import BeautifulSoup
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" 
SENHA_APP = "saty tgmz rzrz yrai" 
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

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

def enviar_email(data_do, arquivo_excel=None, arquivo_pdf=None):
    msg = EmailMessage()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    
    if arquivo_excel:
        msg['Subject'] = f'📊 Vagas de Remoção MPRJ - {data_do}'
        msg.set_content(f'Olá Renan,\n\nIdentificamos vagas no Diário Oficial de {data_do}. Seguem a planilha e o PDF em anexo.')
        with open(arquivo_excel, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=arquivo_excel)
    else:
        msg['Subject'] = f'🔍 Monitoramento MPRJ - {data_do}'
        msg.set_content(f'Olá Renan,\n\nVarredura concluída para o dia {data_do}. Nenhuma vaga de remoção foi encontrada, mas o PDF segue em anexo para conferência.')

    # Esta linha garante que o PDF é enviado mesmo que não existam vagas
    if arquivo_pdf:
        with open(arquivo_pdf, 'rb') as f:
            nome_pdf = f"DO_MPRJ_{data_do.replace('/','-')}.pdf"
            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=nome_pdf)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)

def rodar():
    # ... (código anterior de busca e download do PDF permanece igual) ...
    
    if not link_pdf:
        enviar_email(data_str) # Caso nem o link do PDF seja encontrado
        return

    # Download do PDF
    pdf_content = requests.get(link_pdf).content
    pdf_local = "diario_oficial.pdf"
    with open(pdf_local, "wb") as f:
        f.write(pdf_content)

    doc = fitz.open(pdf_local)
    texto_pdf = "".join([pag.get_text() for pag in doc])
    dados = extrair_dados_com_ia(texto_pdf)

    if dados:
        excel_local = "Vagas_Remocao.xlsx"
        formatar_excel(dados, excel_local, data_str)
        enviar_email(data_str, arquivo_excel=excel_local, arquivo_pdf=pdf_local)
    else:
        # AGORA ENVIA O PDF MESMO SEM DADOS EXTRAÍDOS
        enviar_email(data_str, arquivo_pdf=pdf_local)
