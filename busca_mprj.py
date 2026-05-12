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

def extrair_dados_com_ia(texto_pagina):
    """Envia apenas as páginas relevantes para a IA extrair a tabela"""
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        
        prompt = f"""
        Você é um assistente especialista em leitura de Diários Oficiais.
        No texto abaixo, localize a seção "4. CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA".
        
        Sua tarefa é listar TODAS as vagas numeradas (ex: 1, 2, 3...).
        Para cada item, extraia:
        1. O número do item (ex: 1)
        2. O nome da Promotoria (ex: 2ª Promotoria de Justiça Cível da Capital)
        3. O critério (Antiguidade ou Merecimento)
        4. A origem da vaga (ex: Promoção de Sérgio Bumaschny ou vacância)

        FORMATO DE SAÍDA OBRIGATÓRIO (use ponto e vírgula):
        Item;Órgão;Critério;Origem da Vaga
        
        Importante: Retorne APENAS os dados encontrados, um por linha. 
        Se não houver a seção ou itens buscados, responda apenas: VAZIO
        
        TEXTO PARA ANÁLISE:
        {texto_pagina}
        """
        
        response = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
        res = response.text.strip()
        
        if "VAZIO" in res or ";" not in res:
            return []
            
        # Filtra linhas que possuem o separador ;
        linhas = [l.strip() for l in res.split('\n') if ';' in l]
        return [linha.split(';') for linha in linhas]
    except Exception as e:
        print(f"Erro na comunicação com a IA: {e}")
        return []

def formatar_excel(dados, arquivo, data_do):
    """Gera a planilha Excel com formatação profissional"""
    df = pd.DataFrame(dados, columns=["Item", "Órgão", "Critério", "Origem da Vaga (Decorrente de)"])
    
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='Vagas')
        ws = writer.sheets['Vagas']
        
        # Título
        ws.merge_cells('A1:D1')
        ws['A1'] = f"Vagas de Remoção - DOeMPRJ de {data_do}"
        ws['A1'].font = Font(size=14, bold=True, color="2F5597")
        ws['A1'].alignment = Alignment(horizontal='center')

        # Cabeçalho
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        # Ajuste de colunas
        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 50
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 50
        
        # Bordas e alinhamento
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for row in ws.iter_rows(min_row=3, max_row=len(dados)+3):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical='center')

def enviar_email(data_do, url_pdf, localizado, tem_dados, arquivo_excel=None, arquivo_pdf=None):
    """Envia o e-mail no formato padronizado para mapeamento de erros"""
    msg = EmailMessage()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg['Subject'] = f"Monitoramento DOeMPRJ - {data_do}"
    
    status_arquivo = "Localizado" if localizado else "Não localizado"
    endereco_url = url_pdf if localizado else "Não localizado"
    resultado_texto = "Dados de remoção informados no arquivo em anexo." if tem_dados else "Dados de remoção não encontrados."

    corpo = (
        f"Pesquisa realizada para a data {data_do}.\n\n"
        f"Arquivo: {status_arquivo}\n"
        f"Endereço: {endereco_url}\n"
        f"Resultado: {resultado_texto}"
    )
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
        print("E-mail enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def rodar():
    # --- CONFIGURAÇÃO DE DATA ---
    # Para voltar ao automático (ontem), use: (datetime.now() - timedelta(days=1))
    data_teste = "08.05.2026" 
    data_exibicao = "08/05/2026"
    
    # URL Padrão identificada
    url_pdf = f"https://www.mprj.mp.br/documents/20184/8887328/{data_teste}.pdf"
    
    print(f"Iniciando verificação: {url_pdf}")
    
    try:
        response = requests.get(url_pdf, timeout=30)
        
        if response.status_code != 200:
            print(f"Código {response.status_code}: Arquivo não encontrado.")
            enviar_email(data_exibicao, url_pdf, localizado=False, tem_dados=False)
            return

        # Salva PDF temporário
        pdf_local = "temp_diario.pdf"
        with open(pdf_local, "wb") as f:
            f.write(response.content)

        # Leitura seletiva: envia para a IA apenas as páginas que contêm o termo chave
        doc = fitz.open(pdf_local)
        texto_relevante = ""
        for pagina in doc:
            texto_pag = pagina.get_text()
            if "CONCURSO DE REMOÇÃO" in texto_pag.upper():
                texto_relevante += texto_pag + "\n"
        doc.close()

        if not texto_relevante:
            print("Termo 'CONCURSO DE REMOÇÃO' não encontrado no texto do PDF.")
            enviar_email(data_exibicao, url_pdf, localizado=True, tem_dados=False, arquivo_pdf=pdf_local)
            return

        # Processamento pela IA
        print("Extraindo dados com Gemini...")
        dados = extrair_dados_com_ia(texto_relevante)

        if dados:
            print(f"Sucesso! {len(dados)} vagas encontradas.")
            excel_local = "Vagas_Extraidas.xlsx"
            formatar_excel(dados, excel_local, data_exibicao)
            enviar_email(data_exibicao, url_pdf, localizado=True, tem_dados=True, arquivo_excel=excel_local, arquivo_pdf=pdf_local)
        else:
            print("A IA analisou as páginas mas não encontrou a tabela de vagas.")
            enviar_email(data_exibicao, url_pdf, localizado=True, tem_dados=False, arquivo_pdf=pdf_local)
            
    except Exception as e:
        print(f"Erro crítico na rotina: {e}")

if __name__ == "__main__":
    rodar()
