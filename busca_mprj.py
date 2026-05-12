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

# URL de busca do MPRJ (Filtro Diário Oficial)
URL_SITE = "https://www.mprj.mp.br/busca?p_p_id=br_mp_mprj_internet_busca_web_BuscaPortlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&_br_mp_mprj_internet_busca_web_BuscaPortlet_filtro_param=doerj"

def extrair_dados_com_ia(texto_bruto):
    """Envia o texto do PDF para a IA organizar as vagas"""
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        modelo = "gemini-1.5-flash"
        
        prompt = f"""
        Analise o texto abaixo de um Diário Oficial e localize a seção "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA".
        Extraia as vagas listadas seguindo exatamente este formato, separado por ponto e vírgula (;):
        Item;Órgão;Critério;Origem da Vaga
        
        Exemplo de saída: 4.1;2ª PJ Cível;Antiguidade;Promoção de Sérgio Bumaschny
        
        Retorne APENAS os dados encontrados. Se não houver a seção de remoção, escreva VAZIO.
        Texto: {texto_bruto}
        """
        
        response = client.models.generate_content(model=modelo, contents=prompt)
        res = response.text.strip()
        
        if "VAZIO" in res or ";" not in res:
            return []
            
        linhas = [l.strip() for l in res.split('\n') if ';' in l]
        return [linha.split(';') for linha in linhas]
    except Exception as e:
        print(f"Erro na IA: {e}")
        return []

def formatar_excel(dados, arquivo, data_do):
    """Cria a planilha com o visual profissional solicitado"""
    df = pd.DataFrame(dados, columns=["Item", "Órgão", "Critério", "Origem da Vaga (Decorrente de)"])
    
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='Vagas')
        ws = writer.sheets['Vagas']
        
        # Título Superior
        ws.merge_cells('A1:D1')
        ws['A1'] = f"Resultados encontrados no DOeMPRJ de {data_do}"
        ws['A1'].font = Font(size=14, bold=True, color="2F5597")
        ws['A1'].alignment = Alignment(horizontal='center')

        # Cabeçalho da Tabela
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        # Largura das Colunas
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 55
        
        # Bordas e Zebrado (Linhas alternadas)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for row_idx, row in enumerate(ws.iter_rows(min_row=4, max_row=len(dados)+3), start=4):
            fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid") if row_idx % 2 == 0 else None
            for cell in row:
                if fill: cell.fill = fill
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical='center')

def enviar_email(arquivo, data_hoje, encontrou_vagas=True):
    """Envia o e-mail final para o usuário"""
    msg = EmailMessage()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    
    if encontrou_vagas:
        msg['Subject'] = f'📊 Vagas de Remoção MPRJ - {data_hoje}'
        msg.set_content(f'Olá Renan,\n\nO robô identificou novas vagas de remoção no Diário Oficial de hoje. Confira o arquivo anexo.')
        with open(arquivo, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=arquivo)
    else:
        msg['Subject'] = f'🔍 Monitoramento MPRJ - Varredura {data_hoje}'
        msg.set_content(f'Olá Renan,\n\nO robô realizou a busca hoje, mas não localizou editais de "Concurso de Remoção" publicados.')

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMETENTE, SENHA_APP)
            smtp.send_message(msg)
        print("E-mail enviado!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def rodar():
    print(f"Iniciando varredura em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    
    try:
        # 1. Busca o link do PDF
        response = requests.get(URL_SITE, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')
        link_pdf = ""
        for a in soup.find_all('a', href=True):
            if 'pdf' in a['href'].lower():
                link_pdf = a['href']
                if not link_pdf.startswith('http'):
                    link_pdf = "https://www.mprj.mp.br" + link_pdf
                break

        if not link_pdf:
            print("Nenhum Diário Oficial encontrado no site.")
            enviar_email(None, data_hoje, encontrou_vagas=False)
            return

        # 2. Baixa e lê o conteúdo do PDF
        print(f"Lendo PDF: {link_pdf}")
        pdf_data = requests.get(link_pdf).content
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        texto_pdf = "".join([pagina.get_text() for pagina in doc])
        
        # 3. IA processa e extrai
        dados = extrair_dados_com_ia(texto_pdf)

        if dados:
            arquivo_final = "Vagas_Remocao_MPRJ.xlsx"
            formatar_excel(dados, arquivo_final, data_hoje)
            enviar_email(arquivo_final, data_hoje, encontrou_vagas=True)
        else:
            print("Texto analisado, mas sem vagas de remoção.")
            enviar_email(None, data_hoje, encontrou_vagas=False)

    except Exception as e:
        print(f"Falha na execução: {e}")

if __name__ == "__main__":
    rodar()
