import requests
import pandas as pd
import smtplib
import os
from google import genai
from email.message import EmailMessage
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
EMAIL_DESTINO = "renan.barros@mprj.mp.br"
EMAIL_REMETENTE = "renan.help@gmail.com" 
SENHA_APP = "saty tgmz rzrz yrai" 
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

def extrair_dados_com_ia(caminho_pdf):
    """
    Faz o upload do PDF diretamente para a API do Gemini.
    A IA lê o documento nativamente, preservando o layout visual.
    """
    client = genai.Client(api_key=GEMINI_KEY)
    arquivo_gemini = None
    
    try:
        print("Fazendo upload do PDF para o Gemini...")
        # 1. Envia o arquivo físico para a API
        arquivo_gemini = client.files.upload(file=caminho_pdf)
        
        prompt = """
        Você é um analista de dados especialista em Diários Oficiais.
        Abaixo está um documento PDF completo. Sua missão é localizar e extrair as vagas do "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA".

        ESTRUTURA DE BUSCA:
        - Navegue pelo documento até encontrar a seção que fala sobre Concurso de Remoção para Promotor.
        - Identifique os itens numerados que descrevam a abertura de vaga em uma Promotoria de Justiça.
        - Identifique a origem (ex: vaga decorrente da promoção de...) e o critério (Antiguidade/Merecimento).

        SAÍDA OBRIGATÓRIA (Separada por ponto e vírgula):
        Item;Órgão;Critério;Origem da Vaga

        Exemplo de resposta esperada:
        4.1;2ª Promotoria Cível da Capital;Antiguidade;Promoção de Sérgio Bumaschny
        4.2;1ª Promotoria Criminal;Merecimento;Remoção de Marcos

        Importante: Retorne APENAS as linhas de dados extraídas, sem cabeçalhos ou introduções.
        Se não houver vagas de remoção no documento, responda apenas: VAZIO.
        """
        
        print("Analisando o PDF nativamente...")
        # 2. Passa o arquivo e o prompt para o modelo Flash
        response = client.models.generate_content(
            model="gemini-1.5-flash", 
            contents=[arquivo_gemini, prompt]
        )
        res = response.text.strip()
        
        if "VAZIO" in res or ";" not in res:
            return []
            
        linhas = [l.strip() for l in res.split('\n') if ';' in l]
        return [linha.split(';') for linha in linhas]
        
    except Exception as e:
        print(f"Erro na IA: {e}")
        return []
    finally:
        # 3. Limpeza: Deleta o arquivo dos servidores do Google após a análise
        if arquivo_gemini:
            try:
                client.files.delete(name=arquivo_gemini.name)
                print("Arquivo temporário apagado dos servidores do Gemini.")
            except:
                pass

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

        # Salva o PDF localmente
        pdf_local = "temp_diario.pdf"
        with open(pdf_local, "wb") as f:
            f.write(response.content)

        # Envia o CAMINHO DO ARQUIVO para a IA, não o texto.
        dados = extrair_dados_com_ia(pdf_local)

        if dados:
            print(f"Sucesso! A IA extraiu {len(dados)} vagas.")
            excel_local = "Vagas_Encontradas.xlsx"
            formatar_excel(dados, excel_local, data_exibicao)
            enviar_email(data_exibicao, url_pdf, True, True, excel_local, pdf_local)
        else:
            print("A IA analisou o PDF, mas retornou vazio.")
            enviar_email(data_exibicao, url_pdf, True, False, arquivo_pdf=pdf_local)
            
    except Exception as e:
        print(f"Erro crítico: {e}")

if __name__ == "__main__":
    rodar()
