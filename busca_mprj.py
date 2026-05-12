import requests
import pandas as pd
import smtplib
import os
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

def extrair_dados_com_ia(texto_bruto):
    """Usa a nova biblioteca do Google GenAI para extrair os dados"""
    client = genai.Client(api_key=GEMINI_KEY)
    
    prompt = f"""
    Analise o texto de um Diário Oficial abaixo e localize a seção "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA".
    Analise as informações e extraia as vagas listadas seguindo este formato rigoroso, separando por ponto e vírgula (;):
    Item;Órgão;Critério;Origem da Vaga
    
    Exemplo: 4.1;2ª PJ Cível;Antiguidade;Promoção de Fulano
    
    Retorne APENAS os dados. Se não encontrar nada, retorne 'VAZIO'.
    Texto: {texto_bruto}
    """
    
    response = client.models.generate_content(
        model="gemini-1.5-flash", contents=prompt
    )
    
    texto_resposta = response.text.strip()
    if "VAZIO" in texto_resposta or ";" not in texto_resposta:
        return []
        
    linhas = texto_resposta.split('\n')
    return [linha.split(';') for linha in linhas if ';' in linha]

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
    print("Iniciando varredura com Gemini 1.5 Flash...")
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    
    url_site = "https://www.mprj.mp.br/busca?p_p_id=br_mp_mprj_internet_busca_web_BuscaPortlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&_br_mp_mprj_internet_busca_web_BuscaPortlet_jspPage=%2Fhtml%2Fview.jsp"
    
    try:
        html = requests.get(url_site).text
        soup = BeautifulSoup(html, 'html.parser')
        texto_pagina = soup.get_text()

        dados_extraidos = extrair_dados_com_ia(texto_pagina)

        if not dados_extraidos:
            print("Nenhuma vaga localizada no texto da página.")
            return

        nome_arquivo = "Vagas_Remocao_MPRJ.xlsx"
        formatar_excel(dados_extraidos, nome_arquivo, data_hoje)

        msg = EmailMessage()
        msg['Subject'] = f'📊 Vagas de Remoção MPRJ - {data_hoje}'
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = EMAIL_DESTINO
        msg.set_content(f'Olá Renan,\n\nO robô analisou as publicações de hoje e gerou a planilha em anexo.')

        with open(nome_arquivo, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=nome_arquivo)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMETENTE, SENHA_APP)
            smtp.send_message(msg)
        print("Sucesso!")
    except Exception as e:
        print(f"Erro durante a execução: {e}")

if __name__ == "__main__":
    rodar()
