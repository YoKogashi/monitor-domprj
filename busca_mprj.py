import requests
import pandas as pd
import smtplib
import os
import time
import fitz  # PyMuPDF
from email.message import EmailMessage
from datetime import datetime, timedelta
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURAÇÕES ---
# Lista de e-mails que receberão o relatório
EMAILS_DESTINO = ["renan.barros@mprj.mp.br", "sandro.silva@mprj.mp.br"]
EMAIL_REMETENTE = "renan.help@gmail.com" 
SENHA_APP = "saty tgmz rzrz yrai" 
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

def extrair_dados_com_ia(caminho_pdf):
    tempo_processamento = 0
    status_ia = "Nao iniciado"
    sessao_info = "Nao identificada"
    validade_info = "Nao identificada"
    
    try:
        print("Lendo PDF localmente com PyMuPDF (Estrategia Sniper V4 - Contexto Expandido)...")
        doc = fitz.open(caminho_pdf)
        paginas_alvo = set()
        
        for i, pagina in enumerate(doc):
            texto_pag = pagina.get_text("text")
            
            if "CONCURSO DE REMOÇÃO" in texto_pag.upper() and "PROMOTOR" in texto_pag.upper():
                # Captura a página atual, 1 à frente e 2 para trás (Garante captura do cabeçalho da Ata)
                paginas_alvo.add(max(0, i - 2)) 
                paginas_alvo.add(max(0, i - 1))
                paginas_alvo.add(i)
                if i + 1 < len(doc):
                    paginas_alvo.add(i + 1)
                    
        doc.close()

        paginas_alvo = sorted(list(paginas_alvo))

        if not paginas_alvo:
            return [], "", "", "Falha: Secao nao encontrada em nenhuma pagina do PDF", 0

        texto_alvo = ""
        doc = fitz.open(caminho_pdf)
        for i in paginas_alvo:
            texto_alvo += doc[i].get_text("text") + "\n"
        doc.close()

        print(f"Busca total concluida! {len(paginas_alvo)} paginas capturadas. Enviando para a IA...")
        
        prompt = f"""
        Você é um analista de dados especialista em Diários Oficiais do Ministério Público.
        Abaixo estão trechos do documento contendo a publicação de vagas para "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA", juntamente com as páginas anteriores para contexto (como Atas do Conselho Superior).

        Sua missão é extrair informações gerais e a lista de vagas.

        1. IDENTIFICAÇÃO DA SESSÃO: Procure pelo cabeçalho ou descrição da Ata do Conselho Superior e identifique qual foi a sessão (ex: 1ª, 2ª, 54ª Sessão Ordinária/Extraordinária) e a data de realização da sessão.
        2. VALIDADE: Procure pela data de validade da remoção (ex: "com validade a contar de DD/MM/AAAA").
        3. LISTA DE VAGAS:
           - Identifique os itens numerados (podem começar com 1, 2, 3.1, 4.1, etc).
           - Identifique o Órgão (Nome da Promotoria).
           - Identifique o Critério (Antiguidade ou Merecimento).
           - Identifique a Origem da vaga (ex: decorrente da promoção de Fulano).

        SAÍDA OBRIGATÓRIA (Siga estritamente este formato de blocos):
        SESSAO: [Número e Data da Sessão, ou "Não identificada"]
        VALIDADE: [Data de validade, ou "Não identificada"]
        VAGAS:
        Item;Órgão;Critério;Origem da Vaga

        Importante: Abaixo do cabeçalho de VAGAS, retorne APENAS as linhas separadas por ponto e vírgula. Se não houver vagas de remoção, coloque apenas a palavra VAZIO.

        TEXTO PARA ANÁLISE:
        {texto_alvo}
        """
        
        headers = {'Content-Type': 'application/json'}
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.1}
        }
        
        url_api = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_KEY}"
        
        inicio_ia = time.time()
        response = requests.post(url_api, headers=headers, json=payload)
        fim_ia = time.time()
        tempo_processamento = round(fim_ia - inicio_ia, 2)
        
        if response.status_code == 200:
            status_ia = "Sucesso na comunicacao"
            dados_json = response.json()
            try:
                res = dados_json['candidates'][0]['content']['parts'][0]['text'].strip()
            except KeyError:
                return [], "", "", "Erro ao interpretar resposta da API", tempo_processamento
            
            lista_vagas = []
            is_vagas = False
            
            # Processa o formato de blocos retornado pela IA
            for linha in res.split('\n'):
                linha_limpa = linha.strip()
                if not linha_limpa:
                    continue
                    
                if linha_limpa.startswith("SESSAO:"):
                    sessao_info = linha_limpa.replace("SESSAO:", "").strip()
                elif linha_limpa.startswith("VALIDADE:"):
                    validade_info = linha_limpa.replace("VALIDADE:", "").strip()
                elif linha_limpa.startswith("VAGAS:"):
                    is_vagas = True
                elif is_vagas:
                    if linha_limpa == "VAZIO":
                        break
                    # Ignora a linha de cabeçalho e garante que é uma linha de dados
                    if ';' in linha_limpa and "Item;Órgão" not in linha_limpa: 
                        lista_vagas.append(linha_limpa.split(';'))
                        
            return lista_vagas, sessao_info, validade_info, status_ia, tempo_processamento
        else:
            erro_msg = f"Erro {response.status_code}: {response.text}"
            print(erro_msg)
            return [], "", "", f"Erro na API ({response.status_code})", tempo_processamento
        
    except Exception as e:
        print(f"Erro no processamento da IA: {e}")
        return [], "", "", f"Erro critico: {str(e)}", tempo_processamento

def formatar_excel(dados, arquivo, data_do):
    df = pd.DataFrame(dados, columns=["Item", "Órgão", "Critério", "Origem da Vaga (Decorrente de)"])
    with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='Vagas')
        ws = writer.sheets['Vagas']
        
        ws.merge_cells('A1:D1')
        ws['A1'] = f"Resultados encontrados no DOeMPRJ de {data_do}"
        ws['A1'].font = Font(size=14, bold=True, color="2F5597")
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        for cell in ws[3]:
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for row in ws.iter_rows(min_row=3, max_row=len(dados)+3):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical='center')

        for row_idx in range(1, len(dados) + 4):
            ws.row_dimensions[row_idx].height = 25
            
        for col in ws.iter_cols(min_row=3, max_row=len(dados)+3, min_col=1, max_col=4):
            max_length = 0
            col_letter = col[0].column_letter
            
            for cell in col:
                if cell.value:
                    tamanho_texto = max([len(linha) for linha in str(cell.value).split('\n')])
                    if tamanho_texto > max_length:
                        max_length = tamanho_texto
            
            ws.column_dimensions[col_letter].width = max_length + 3

def enviar_email(data_do, url_pdf, localizado, status_dl, status_ia, tem_dados, qtd_vagas=0, tempo_ia=0, tamanho_kb=0, sessao_info="", validade_info="", arquivo_excel=None, arquivo_pdf=None):
    msg = EmailMessage()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = ", ".join(EMAILS_DESTINO)
    msg['Subject'] = f"Monitoramento DOeMPRJ - {data_do}"
    
    status_arquivo = "Localizado" if localizado else "Nao localizado"
    endereco_url = url_pdf if localizado else "Nao localizado"
    
    if tem_dados:
        resultado_texto = (
            f"Sucesso. {qtd_vagas} vagas de remocao extraidas e informadas no arquivo em anexo.\n\n"
            f"--- INFORMACOES DA PUBLICACAO ---\n"
            f"Sessao do Conselho: {sessao_info}\n"
            f"Validade da Remocao: {validade_info}\n"
        )
    else:
        resultado_texto = "Dados de remocao nao encontrados."

    corpo = (
        f"Pesquisa realizada para o Diario Oficial de {data_do}.\n"
        f"{resultado_texto}\n\n"
        f"-------------------------------------------------\n"
        f"Relatorio de Execucao - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"-------------------------------------------------\n\n"
        f"Arquivo DOe: {status_arquivo}\n"
        f"Endereco URL: {endereco_url}\n"
        f"Status do Download: {status_dl} (Tamanho: {tamanho_kb} KB)\n"
        f"Comunicacao com IA: {status_ia}\n"
        f"Tempo de Leitura da IA: {tempo_ia} segundos\n"
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
        print("E-mail formatado enviado com sucesso para todos os destinatarios.")
    except Exception as e:
        print(f"Erro no envio do e-mail: {e}")

def rodar():
    ontem = datetime.now() - timedelta(days=1)
    data_alvo = ontem.strftime("%d.%m.%Y")       
    data_exibicao = ontem.strftime("%d/%m/%Y")   
    
    url_pdf = f"https://www.mprj.mp.br/documents/20184/8887328/{data_alvo}.pdf"
    
    localizado = False
    status_download = "Nao iniciado"
    status_ia = "Nao iniciado"
    tem_dados = False
    tamanho_pdf_kb = 0
    
    try:
        print(f"Buscando PDF para a data {data_exibicao}: {url_pdf}")
        response = requests.get(url_pdf, timeout=30)
        
        if response.status_code == 200:
            localizado = True
            status_download = "Bem sucedido"
            
            pdf_local = "temp_diario.pdf"
            with open(pdf_local, "wb") as f:
                f.write(response.content)

            tamanho_pdf_kb = round(os.path.getsize(pdf_local) / 1024, 2)

            dados, sessao_info, validade_info, status_ia, tempo_processamento = extrair_dados_com_ia(pdf_local)

            if dados:
                tem_dados = True
                qtd_vagas = len(dados)
                excel_local = "Vagas_Encontradas.xlsx"
                formatar_excel(dados, excel_local, data_exibicao)
                
                enviar_email(data_exibicao, url_pdf, localizado, status_download, status_ia, tem_dados, 
                             qtd_vagas=qtd_vagas, tempo_ia=tempo_processamento, tamanho_kb=tamanho_pdf_kb, 
                             sessao_info=sessao_info, validade_info=validade_info,
                             arquivo_excel=excel_local, arquivo_pdf=pdf_local)
            else:
                enviar_email(data_exibicao, url_pdf, localizado, status_download, status_ia, tem_dados, 
                             tempo_ia=tempo_processamento, tamanho_kb=tamanho_pdf_kb, arquivo_pdf=pdf_local)
        
        else:
            status_download = f"Mal sucedido (Erro {response.status_code})"
            enviar_email(data_exibicao, url_pdf, localizado, status_download, status_ia, tem_dados)
            
    except Exception as e:
        status_download = f"Mal sucedido ({str(e)})"
        print(f"Erro critico: {e}")
        enviar_email(data_exibicao, url_pdf, localizado, status_download, status_ia, tem_dados)

if __name__ == "__main__":
    rodar()
