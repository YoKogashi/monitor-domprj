import requests
import pandas as pd
import smtplib
import os
import fitz  # PyMuPDF
import re    # Biblioteca para expressões regulares (o nosso "bisturi")
from google import genai
from email.message import EmailMessage
from datetime import datetime, timedelta
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" 
SENHA_APP = "saty tgmz rzrz yrai" 
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

def extrair_dados_com_ia(texto_recortado):
    """
    Agora a IA recebe apenas o 'filé mignon' do texto, 
    já recortado e focado apenas na remoção.
    """
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        
        prompt = f"""
        Você é um assistente especialista em estruturação de dados.
        Abaixo está um trecho de um Diário Oficial que trata de vagas de REMOÇÃO.

        SUA TAREFA:
        Extraia a lista de vagas. Cada item de vaga costuma seguir este padrão:
        [Número do Item] - [Nome da Promotoria] - [Origem da vaga (ex: promoção, vacância)] - [Critério (Antiguidade/Merecimento)]

        Crie uma tabela com exatidão usando PONTO E VÍRGULA (;) como separador.

        SAÍDA OBRIGATÓRIA E ÚNICA:
        Item;Órgão;Critério;Origem da Vaga

        Exemplo de resposta esperada:
        4.1;2ª Promotoria Cível;Antiguidade;Promoção de Sérgio Bumaschny
        4.2;1ª Promotoria Criminal;Merecimento;Remoção de Marcos

        IMPORTANTE: 
        - Responda APENAS com os dados formatados, linha por linha.
        - Não invente nada. Se não houver itens descrevendo Promotorias e critérios, responda: VAZIO.

        TEXTO PARA EXTRAÇÃO:
        {texto_recortado}
        """
        
        response = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
        res = response.text.strip()
        
        if "VAZIO" in res or ";" not in res:
            return []
            
        linhas = [l.strip() for l in res.split('\n') if ';' in l]
        return [linha.split(';') for linha in linhas]
    except Exception as e:
        print(f"Erro na IA: {e}")
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
        print("Relatório enviado com sucesso.")
    except Exception as e:
        print(f"Erro no envio do e-mail: {e}")

def rodar():
    # --- DATA DE TESTE (08/05/2026) ---
    data_alvo = "08.05.2026"
    data_exibicao = "08/05/2026"
    url_pdf = f"https://www.mprj.mp.br/documents/20184/8887328/{data_alvo}.pdf"
    
    try:
        print(f"Buscando PDF: {url_pdf}")
        response = requests.get(url_pdf, timeout=30)
        if response.status_code != 200:
            print("Arquivo não encontrado no site.")
            enviar_email(data_exibicao, url_pdf, False, False)
            return

        pdf_local = "temp_diario.pdf"
        with open(pdf_local, "wb") as f:
            f.write(response.content)

        print("Extraindo texto do PDF...")
        doc = fitz.open(pdf_local)
        texto_completo = ""
        for pagina in doc:
            texto_completo += pagina.get_text("text") + "\n"
        doc.close()

        # O BISTURI: Usando Expressões Regulares para encontrar a seção exata
        print("Procurando seção de Concurso de Remoção...")
        
        # Procura a frase exata (ignorando quebras de linha estranhas que o PDF possa ter)
        padrao_busca = re.compile(r"CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA(.*?)(\n[A-Z0-9.\s]+:|\n\n\n|\Z)", re.DOTALL | re.IGNORECASE)
        match = padrao_busca.search(texto_completo)

        texto_para_ia = ""
        if match:
            # Pega o texto encontrado a partir da frase chave
            texto_encontrado = match.group(0)
            
            # Limita a 5000 caracteres (suficiente para as vagas, mas elimina o resto do diário)
            texto_para_ia = texto_encontrado[:5000]
            print(f"Seção encontrada! Recortei {len(texto_para_ia)} caracteres.")
        else:
            print("Aviso: Expressão regular falhou. Tentando busca manual por palavras...")
            # Plano B: Busca simples de posição
            indice_inicio = texto_completo.upper().find("CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA")
            if indice_inicio != -1:
                texto_para_ia = texto_completo[indice_inicio : indice_inicio + 5000]
                print(f"Seção encontrada manualmente! Recortei {len(texto_para_ia)} caracteres.")

        if not texto_para_ia:
            print("Não encontrei a seção de remoção de promotores no texto.")
            enviar_email(data_exibicao, url_pdf, True, False, arquivo_pdf=pdf_local)
            return

        print("Enviando recorte do texto para o Gemini...")
        dados = extrair_dados_com_ia(texto_para_ia)

        if dados:
            print(f"Sucesso! A IA extraiu {len(dados)} vagas.")
            excel_local = "Vagas_Encontradas.xlsx"
            formatar_excel(dados, excel_local, data_exibicao)
            enviar_email(data_exibicao, url_pdf, True, True, excel_local, pdf_local)
        else:
            print("A IA recebeu o recorte correto, mas retornou vazio (falha de formatação).")
            enviar_email(data_exibicao, url_pdf, True, False, arquivo_pdf=pdf_local)
            
    except Exception as e:
        print(f"Erro crítico: {e}")

if __name__ == "__main__":
    rodar()
