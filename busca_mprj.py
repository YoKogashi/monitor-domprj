import requests
import pandas as pd
import smtplib
import os
import fitz  # PyMuPDF
from google import genai
from email.message import EmailMessage
from datetime import datetime
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

def enviar_email(data_do, url_pdf, localizado, tem_dados, arquivo_excel=None, arquivo_pdf=None):
    msg = EmailMessage()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg['Subject'] = f"Monitoramento DOeMPRJ - {data_do}"
    
    # Padronização do corpo do e-mail conforme solicitado
    status_arquivo = "Localizado" if localizado else "Não localizado"
    endereco_url = url_pdf if localizado else "Não localizado"
    
    if tem_dados:
        resultado_texto = "Dados de remoção informados no arquivo em anexo."
    else:
        resultado_texto = "Dados de remoção não encontrados."

    corpo = (
        f"Pesquisa realizada para a data {data_do}.\n\n"
        f"Arquivo: {status_arquivo}\n"
        f"Endereço: {endereco_url}\n"
        f"Resultado: {resultado_texto}"
    )
    msg.set_content(corpo)

    # Anexos
    if arquivo_excel:
        with open(arquivo_excel, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=arquivo_excel)
    
    if arquivo_pdf:
        with open(arquivo_pdf, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=f"DO_MPRJ_{data_do.replace('/','-')}.pdf")

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)

def rodar():
    # --- DATA DO TESTE (08/05/2026) ---
    data_alvo = "08.05.2026"
    data_exibicao = "08/05/2026"
    
    # Montagem da URL baseada no padrão identificado
    url_pdf = f"https://www.mprj.mp.br/documents/20184/8887328/{data_alvo}.pdf"
    
    print(f"Tentando acessar: {url_pdf}")
    
    try:
        response = requests.get(url_pdf, timeout=30)
        
        if response.status_code != 200:
            print("Arquivo não localizado no servidor.")
            enviar_email(data_exibicao, url_pdf, localizado=False, tem_dados=False)
            return

        # Se localizou o arquivo
        pdf_local = "diario_oficial.pdf"
        with open(pdf_local, "wb") as f:
            f.write(response.content)

        doc = fitz.open(pdf_local)
        texto_pdf = "".join([pag.get_text() for pag in doc])
        dados = extrair_dados_com_ia(texto_pdf)

        if dados:
            excel_local = "Vagas_Remocao.xlsx"
            formatar_excel(dados, excel_local, data_exibicao)
            enviar_email(data_exibicao, url_pdf, localizado=True, tem_dados=True, arquivo_excel=excel_local, arquivo_pdf=pdf_local)
        else:
            enviar_email(data_exibicao, url_pdf, localizado=True, tem_dados=False, arquivo_pdf=pdf_local)
            
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    rodar()
