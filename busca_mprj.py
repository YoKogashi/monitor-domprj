import requests
import pandas as pd
import smtplib
import os
from google import genai
from email.message import EmailMessage
from datetime import datetime
from bs4 import BeautifulSoup
import fitz # Importante: precisa do PyMuPDF (instalado via pip install pymupdf)
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" 
SENHA_APP = "saty tgmz rzrz yrai" 
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

def extrair_dados_com_ia(texto_bruto):
    # Removi configurações extras para usar o padrão de fábrica
    client = genai.Client(api_key=GEMINI_KEY)
    
    # Tentaremos o modelo Pro, que costuma estar disponível em todas as regiões v1
    modelo_operacional = "gemini-1.5-pro"
    
    prompt = f"""
    Analise o texto abaixo e localize a seção "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA".
    Extraia as vagas listadas seguindo o formato: Item;Órgão;Critério;Origem da Vaga
    Texto: {texto_bruto}
    """
    
    try:
        # Chamada direta e simplificada
        response = client.models.generate_content(
            model=modelo_operacional, 
            contents=prompt
        )
        
        if not response.text:
            return []
            
        linhas = [l.strip() for l in response.text.split('\n') if ';' in l]
        return [linha.split(';') for linha in linhas]

    except Exception as e:
        print(f"Erro na tentativa com {modelo_operacional}: {e}")
        return []

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
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 55
        
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for row_idx, row in enumerate(ws.iter_rows(min_row=4, max_row=len(dados)+3), start=4):
            fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid") if row_idx % 2 == 0 else None
            for cell in row:
                if fill: cell.fill = fill
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical='center')

def rodar():
    print("Iniciando busca no PDF...")
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    
    # 1. Localiza o link do PDF no site
    html = requests.get(URL_SITE).text
    soup = BeautifulSoup(html, 'html.parser')
    link_pdf = ""
    for a in soup.find_all('a', href=True):
        if 'pdf' in a['href'].lower():
            link_pdf = a['href']
            if not link_pdf.startswith('http'):
                link_pdf = "https://www.mprj.mp.br" + link_pdf
            break

    if link_pdf:
        # 2. Baixa e lê o PDF
        response = requests.get(link_pdf)
        with open("diario.pdf", "wb") as f:
            f.write(response.content)
        
        doc = fitz.open("diario.pdf")
        texto_completo = ""
        for pagina in doc:
            texto_completo += pagina.get_text()
        
        # 3. Manda o texto REAL do PDF para a IA
        dados_extraidos = extrair_dados_com_ia(texto_completo)
        
        if dados_extraidos:
            nome_arquivo = "Vagas_Remocao_MPRJ.xlsx"
            formatar_excel(dados_extraidos, nome_arquivo, data_hoje)
            enviar_email(nome_arquivo, data_hoje)
            print("Sucesso: PDF lido e e-mail enviado!")
        else:
            print("IA não encontrou as palavras-chave dentro do PDF.")
