import requests
import pandas as pd
import smtplib
import os
import google.generativeai as genai
from email.message import EmailMessage
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" # <--- MUDE AQUI
SENHA_APP = "saty tgmz rzrz yrai"    # <--- COLE AQUI A SENHA DO GMAIL
GEMINI_KEY = os.getenv("GEMINI_API_KEY") # Pega a chave que guardou no GitHub

def extrair_dados_com_ia(texto_bruto):
    """Instrui a IA a ler o texto e retornar apenas os dados estruturados"""
    genai.configure(api_key=GEMINI_KEY)
    model = genai.GenerativeModel('gemini-pro')
    
    prompt = f"""
    Analise o texto de um Diário Oficial abaixo e localize a seção "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA".
    Extraia as vagas listadas seguindo este formato rigoroso, separando por ponto e vírgula (;):
    Item;Órgão;Critério;Origem da Vaga
    
    Se houver mais de uma vaga, coloque uma em cada linha. 
    Não escreva mais nada além dos dados.
    Texto: {texto_bruto}
    """
    
    response = model.generate_content(prompt)
    linhas = response.text.strip().split('\n')
    # Transforma as linhas da IA em uma lista de listas para o Pandas
    return [linha.split(';') for linha in linhas if ';' in linha]

def formatar_excel(dados, arquivo, data_do):
    """Cria o Excel com a formatação visual profissional"""
    df = pd.DataFrame(dados, columns=["Item", "Órgão", "Critério", "Origem da Vaga (Decorrente de)"])
    
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='Vagas')
        ws = writer.sheets['Vagas']
        
        # Título
        ws.merge_cells('A1:D1')
        ws['A1'] = f"Resultados encontrados no DOeMPRJ de {data_do}"
        ws['A1'].font = Font(size=14, bold=True, color="2F5597")
        ws['A1'].alignment = Alignment(horizontal='center')

        # Cabeçalho
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        # Ajustes Visuais
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
    print("Iniciando varredura dinâmica...")
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    
    # 1. ACESSA O SITE E PEGA O TEXTO (Simulação de extração de texto do PDF)
    # Em uma implementação completa, usaríamos uma lib como PyMuPDF para ler o PDF do link
    # Por agora, o robô busca o conteúdo textual da página de busca que você enviou
    url_site = "https://www.mprj.mp.br/busca..." # URL completa aqui
    html = requests.get(url_site).text
    soup = BeautifulSoup(html, 'html.parser')
    texto_pagina = soup.get_text()

    # 2. IA PROCESSA O TEXTO
    dados_extraidos = extrair_dados_com_ia(texto_pagina)

    if not dados_extraidos:
        print("Nenhuma vaga de remoção encontrada hoje.")
        return

    # 3. GERA E FORMATA O EXCEL
    nome_arquivo = "Vagas_Remocao_MPRJ.xlsx"
    formatar_excel(dados_extraidos, nome_arquivo, data_hoje)

    # 4. ENVIA E-MAIL
    msg = EmailMessage()
    msg['Subject'] = f'📊 Vagas de Remoção MPRJ - {data_hoje}'
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg.set_content(f'Olá Renan,\n\nO robô analisou o DOeMPRJ de hoje e extraiu as vagas em anexo.')

    with open(nome_arquivo, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=nome_arquivo)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)
    print("Relatório dinâmico enviado!")

if __name__ == "__main__":
    rodar()
