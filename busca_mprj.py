import requests
import pandas as pd
import smtplib
import os
import fitz  # PyMuPDF
from google import genai
from email.message import EmailMessage
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" 
SENHA_APP = "saty tgmz rzrz yrai" 
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

def extrair_dados_com_ia(texto_relevante):
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        prompt = f'''
        Você é um analista de dados especialista em Diários Oficiais.
        Sua missão é extrair vagas de CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA do texto abaixo.

        ESTRUTURA DE BUSCA:
        - Localize itens que descrevam uma Promotoria de Justiça.
        - Identifique o trecho "vaga decorrente de..." ou "vaga decorrente da...".
        - Identifique o critério (Antiguidade ou Merecimento) que aparece entre parênteses.

        REGRAS DE EXTRAÇÃO:
        1. Capture o Identificador (ex: 4.1, 1, A, etc).
        2. Capture o Órgão (Nome da Promotoria).
        3. Capture o Critério (Apenas 'Antiguidade' ou 'Merecimento').
        4. Capture a Origem (O motivo completo: ex: "Promoção de Fulano de Tal").

        SAÍDA OBRIGATÓRIA (Separada por ponto e vírgula):
        Item;Órgão;Critério;Origem da Vaga

        Importante: Se encontrar as vagas mas elas não tiverem número, invente uma sequência (1, 2, 3). 
        Retorne APENAS as linhas de dados. Se não houver vagas de remoção, responda: VAZIO.

        TEXTO PARA ANÁLISE:
        {texto_relevante}
        '''
        response = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
        res = response.text.strip()
        if "VAZIO" in res or ";" not in res: return []
        return [l.strip().split(';') for l in res.split('\n') if ';' in l]
    except Exception as e:
        print(f"Erro IA: {e}")
        return []

def formatar_excel(dados, arquivo, data_do):
    df = pd.DataFrame(dados, columns=["Item", "Órgão", "Critério", "Origem da Vaga (Decorrente de)"])
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='Vagas')
        ws = writer.sheets['Vagas']
        ws.merge_cells('A1:D1')
        ws['A1'] = f"Vagas de Remoção - DOeMPRJ de {data_do}"
        ws['A1'].font = Font(size=14, bold=True, color="2F5597")
        ws['A1'].alignment = Alignment(horizontal='center')
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 50
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 55
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for row in ws.iter_rows(min_row=3, max_row=len(dados)+3):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical='center')

def enviar_email(data_do, url_pdf, localizado, tem_dados, arquivo_excel=None, arquivo_pdf=None):
    msg = EmailMessage()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg['Subject'] = f"Monitoramento DOeMPRJ - {data_do}"
    status_arquivo = "Localizado" if localizado else "Não localizado"
    endereco_url = url_pdf if localizado else "Não localizado"
    resultado_texto = "Dados de remoção informados no arquivo em anexo." if tem_dados else "Dados de remoção não encontrados."
    corpo = f"Pesquisa realizada para a data {data_do}.\n\nArquivo: {status_arquivo}\nEndereço: {endereco_url}\nResultado: {resultado_texto}"
    msg.set_content(corpo)
    if arquivo_excel:
        with open(arquivo_excel, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=arquivo_excel)
    if arquivo_pdf:
        with open(arquivo_pdf, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=f"DO_MPRJ_{data_do.replace('/','-')}.pdf")
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMETENTE, SENHA_APP)
            smtp.send_message(msg)
    except Exception as e: print(f"Erro Email: {e}")

def rodar():
    data_alvo = "08.05.2026"
    data_exibicao = data_alvo.replace(".", "/")
    url_pdf = f"https://www.mprj.mp.br/documents/20184/8887328/{data_alvo}.pdf"
    try:
        r = requests.get(url_pdf, timeout=30)
        if r.status_code != 200:
            enviar_email(data_exibicao, url_pdf, False, False)
            return
        pdf_local = "temp.pdf"
        with open(pdf_local, "wb") as f: f.write(r.content)
        doc = fitz.open(pdf_local)
        texto = ""
        for p in doc:
            t = p.get_text("blocks")
            f_t = "\n".join([b[4] for b in t])
            if "CONCURSO DE REMOÇÃO" in f_t.upper() and "PROMOTOR" in f_t.upper():
                texto += f_t + "\n"
        doc.close()
        if not texto:
            enviar_email(data_exibicao, url_pdf, True, False, arquivo_pdf=pdf_local)
            return
        dados = extrair_dados_com_ia(texto)
        if dados:
            excel = "Vagas_MPRJ.xlsx"
            formatar_excel(dados, excel, data_exibicao)
            enviar_email(data_exibicao, url_pdf, True, True, excel, pdf_local)
        else:
            enviar_email(data_exibicao, url_pdf, True, False, arquivo_pdf=pdf_local)
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    rodar()
