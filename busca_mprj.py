import json
import os
import re
import smtplib
import time
from datetime import datetime, timedelta
from email.message import EmailMessage

import fitz  # PyMuPDF
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import requests

# --- CONFIGURAÇÕES E SEGURANÇA ---
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "renan.help@gmail.com")
SENHA_APP = os.getenv("EMAIL_SENHA_APP")
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

EMAILS_SEM_RESULTADO = ["renan.barros@mprj.mp.br"]
EMAILS_COM_RESULTADO = [
    "renan.barros@mprj.mp.br",
    "sandro.silva@mprj.mp.br",    
]

# Termo da 2ª busca
TERMO_VIOLENCIA_DOMESTICA = (
    "Promotoria de Justiça junto aos II e IV Juizados de Violência Doméstica"
)

# Schema estruturado para a extração do Concurso de Remoção (Busca 1)
JSON_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "sessao": {"type": "STRING"},
        "validade": {"type": "STRING"},
        "vagas": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "item": {"type": "STRING"},
                    "orgao": {"type": "STRING"},
                    "criterio": {"type": "STRING"},
                    "origem": {"type": "STRING"},
                },
                "required": ["item", "orgao", "criterio", "origem"],
            },
        },
    },
    "required": ["sessao", "validade", "vagas"],
}


# ==========================================
# BUSCA 1: REMOÇÃO COM IA
# ==========================================
def extrair_dados_com_ia(caminho_pdf):
    tempo_processamento = 0
    status_ia = "Não iniciado"

    try:
        doc = fitz.open(caminho_pdf)
        paginas_alvo = set()

        for i, pagina in enumerate(doc):
            texto_pag = pagina.get_text("text").upper()
            if (
                "CONCURSO DE REMOÇÃO" in texto_pag
                and "PROMOTOR DE JUSTIÇA" in texto_pag
            ):
                paginas_alvo.update(
                    [max(0, i - 2), max(0, i - 1), i, min(len(doc) - 1, i + 1)]
                )

        if not paginas_alvo:
            doc.close()
            return (
                [],
                "",
                "",
                "Seção de remoção não encontrada",
                tempo_processamento,
            )

        texto_alvo = ""
        for i in sorted(paginas_alvo):
            texto_alvo += doc[i].get_text("text") + "\n--- QUEBRA DE PÁGINA ---\n"
        doc.close()

        prompt = f"""
        Você é um analista de dados especialista em Diários Oficiais do Ministério Público.
        Analise os trechos abaixo e extraia os dados do "CONCURSO DE REMOÇÃO PARA PROMOTOR DE JUSTIÇA":
        1. SESSÃO: Número e data da Sessão do Conselho Superior (ex: '54ª Sessão Ordinária de 10/05/2026' ou 'Não identificada').
        2. VALIDADE: Data de validade da remoção (ex: '01/06/2026' ou 'Não identificada').
        3. LISTA DE VAGAS: Extraia item numerado, Órgão/Promotoria, Critério (Antiguidade/Merecimento) e Origem da Vaga (ex: Decorrente da promoção de...). Se não houver vagas, retorne lista vazia.

        TEXTO:
        {texto_alvo}
        """

        url_api = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_KEY}"
        headers = {"Content-Type": "application/json"}
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.0,
                "responseMimeType": "application/json",
                "responseSchema": JSON_SCHEMA,
            },
        }

        inicio_ia = time.time()
        response = requests.post(
            url_api, headers=headers, json=payload, timeout=60
        )
        tempo_processamento = round(time.time() - inicio_ia, 2)

        if response.status_code == 200:
            res_json = response.json()
            raw_text = res_json["candidates"][0]["content"]["parts"][0]["text"]
            resultado = json.loads(raw_text)

            sessao_info = resultado.get("sessao", "Não identificada")
            validade_info = resultado.get("validade", "Não identificada")
            lista_vagas = [
                [v["item"], v["orgao"], v["criterio"], v["origem"]]
                for v in resultado.get("vagas", [])
            ]

            return (
                lista_vagas,
                sessao_info,
                validade_info,
                "Sucesso na comunicação",
                tempo_processamento,
            )
        else:
            return (
                [],
                "",
                "",
                f"Erro na API ({response.status_code})",
                tempo_processamento,
            )

    except Exception as e:
        return (
            [],
            "",
            "",
            f"Erro crítico: {str(e)}",
            round(time.time() - inicio_ia, 2) if "inicio_ia" in locals() else 0,
        )


# ==========================================
# BUSCA 2: TEXTO EXATO (IPSIS LITTERIS)
# ==========================================
def buscar_paragrafos_exatos(caminho_pdf, termo_busca):
    """Varre o documento e retorna os parágrafos/blocos onde o termo ocorre, com número de página."""
    achados = []
    # Normaliza quebras de linha e múltiplos espaços para busca flexível
    termo_padronizado = re.sub(r"\s+", " ", termo_busca.strip()).lower()

    try:
        doc = fitz.open(caminho_pdf)
        for num_pag, pagina in enumerate(doc):
            # get_text("blocks") retorna blocos no formato: (x0, y0, x1, y1, texto, bloco_num, tipo)
            blocos = pagina.get_text("blocks")
            for b in blocos:
                texto_bloco = b[4]
                texto_bloco_normalizado = re.sub(
                    r"\s+", " ", texto_bloco
                ).lower()

                if termo_padronizado in texto_bloco_normalizado:
                    achados.append(
                        {
                            "pagina": num_pag + 1,
                            "paragrafo": texto_bloco.strip(),
                        }
                    )
        doc.close()
    except Exception as e:
        print(f"Erro ao buscar termo exato: {e}")

    return achados


# ==========================================
# EXPORTAÇÃO EXCEL
# ==========================================
def formatar_excel(dados_vagas, achados_termo, arquivo, data_do):
    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        header_fill = PatternFill(
            start_color="2F5597", end_color="2F5597", fill_type="solid"
        )

        # Aba 1: Vagas de Remoção (se houver)
        if dados_vagas:
            df_vagas = pd.DataFrame(
                dados_vagas,
                columns=[
                    "Item",
                    "Órgão",
                    "Critério",
                    "Origem da Vaga (Decorrente de)",
                ],
            )
            df_vagas.to_excel(writer, index=False, startrow=2, sheet_name="Remoção")
            ws_rem = writer.sheets["Remoção"]

            ws_rem.merge_cells("A1:D1")
            ws_rem["A1"] = (
                f"Vagas de Remoção encontradas no DOeMPRJ de {data_do}"
            )
            ws_rem["A1"].font = Font(size=13, bold=True, color="2F5597")
            ws_rem["A1"].alignment = Alignment(
                horizontal="center", vertical="center"
            )

            for cell in ws_rem[3]:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            for row in ws_rem.iter_rows(
                min_row=3, max_row=len(dados_vagas) + 3
            ):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(
                        wrap_text=True, vertical="center"
                    )

            ws_rem.column_dimensions["A"].width = 10
            ws_rem.column_dimensions["B"].width = 45
            ws_rem.column_dimensions["C"].width = 18
            ws_rem.column_dimensions["D"].width = 45

        # Aba 2: Ocorrências de Violência Doméstica
        if achados_termo:
            df_termo = pd.DataFrame(achados_termo)
            df_termo.rename(
                columns={
                    "pagina": "Página no DOe",
                    "paragrafo": "Parágrafo Completo (Ipsis Litteris)",
                },
                inplace=True,
            )
            df_termo.to_excel(
                writer,
                index=False,
                startrow=2,
                sheet_name="Violência Doméstica",
            )
            ws_termo = writer.sheets["Violência Doméstica"]

            ws_termo.merge_cells("A1:B1")
            ws_termo["A1"] = (
                f"Menções à Promotoria - DOeMPRJ de {data_do}"
            )
            ws_termo["A1"].font = Font(size=13, bold=True, color="2F5597")
            ws_termo["A1"].alignment = Alignment(
                horizontal="center", vertical="center"
            )

            for cell in ws_termo[3]:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            for row in ws_termo.iter_rows(
                min_row=3, max_row=len(achados_termo) + 3
            ):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(
                        wrap_text=True, vertical="center"
                    )

            ws_termo.column_dimensions["A"].width = 15
            ws_termo.column_dimensions["B"].width = 95


# ==========================================
# ENVIO DE E-MAIL
# ==========================================
def enviar_email(
    emails_destino,
    data_do,
    url_pdf,
    localizado,
    status_dl,
    status_ia,
    tem_vagas,
    qtd_vagas=0,
    tempo_ia=0,
    tamanho_kb=0,
    sessao_info="",
    validade_info="",
    achados_termo=None,
    arquivo_excel=None,
    arquivo_pdf=None,
):
    if not SENHA_APP:
        print("ALERTA: EMAIL_SENHA_APP não configurada. E-mail não enviado.")
        return

    msg = EmailMessage()
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = ", ".join(emails_destino)
    msg["Subject"] = f"Monitoramento DOeMPRJ - {data_do}"

    status_arquivo = "Localizado" if localizado else "Não localizado"
    endereco_url = url_pdf if localizado else "Não localizado"

    # Bloco 1: Concurso de Remoção
    if tem_vagas:
        resultado_remocao = (
            f"• Concurso de Remoção: {qtd_vagas} vaga(s) identificada(s).\n"
            f"  - Sessão do Conselho: {sessao_info}\n"
            f"  - Validade da Remoção: {validade_info}\n"
        )
    else:
        resultado_remocao = (
            "• Concurso de Remoção: Nenhuma vaga identificada nesta edição.\n"
        )

    # Bloco 2: Termo Específico (Violência Doméstica)
    if achados_termo:
        resultado_termo = (
            f"• Menção à Promotoria (II e IV JVD): {len(achados_termo)} ocorrência(s) encontrada(s)!\n\n"
            f"--- PARÁGRAFO(S) ENCONTRADO(S) (IPSIS LITTERIS) ---\n"
        )
        for idx, item in enumerate(achados_termo, start=1):
            resultado_termo += (
                f"[{idx}] Página {item['pagina']}:\n\"{item['paragrafo']}\"\n\n"
            )
    else:
        resultado_termo = (
            "• Menção à Promotoria (II e IV JVD): Nenhuma ocorrência encontrada.\n"
        )

    corpo = (
        f"Pesquisa realizada para o Diário Oficial de {data_do}.\n\n"
        f"--- RESULTADOS DAS BUSCAS ---\n"
        f"{resultado_remocao}\n"
        f"{resultado_termo}"
        f"-------------------------------------------------\n"
        f"Relatório de Execução - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"-------------------------------------------------\n"
        f"Arquivo DOe: {status_arquivo}\n"
        f"Endereço URL: {endereco_url}\n"
        f"Status do Download: {status_dl} (Tamanho: {tamanho_kb} KB)\n"
        f"Comunicação com IA: {status_ia}\n"
        f"Tempo de Leitura da IA: {tempo_ia} segundos\n"
    )
    msg.set_content(corpo)

    if arquivo_excel and os.path.exists(arquivo_excel):
        with open(arquivo_excel, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="xlsx",
                filename=os.path.basename(arquivo_excel),
            )

    if arquivo_pdf and os.path.exists(arquivo_pdf):
        with open(arquivo_pdf, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename=f"DO_MPRJ_{data_do.replace('/', '-')}.pdf",
            )

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_REMETENTE, SENHA_APP)
            smtp.send_message(msg)
        print("E-mail enviado com sucesso.")
    except Exception as e:
        print(f"Erro no envio do e-mail: {e}")


# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
def rodar():
    ontem = datetime.now() - timedelta(days=1)
    data_alvo = ontem.strftime("%d.%m.%Y")
    data_exibicao = ontem.strftime("%d/%m/%Y")

    url_pdf = f"https://www.mprj.mp.br/documents/20184/8887328/{data_alvo}.pdf"
    pdf_local = "temp_diario.pdf"
    excel_local = "Resultados_DOeMPRJ.xlsx"

    localizado = False
    status_download = "Não iniciado"
    status_ia = "Não iniciado"
    tamanho_pdf_kb = 0

    try:
        print(f"Buscando PDF para {data_exibicao}: {url_pdf}")
        response = requests.get(
            url_pdf,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                )
            },
            timeout=45,
        )

        if response.status_code == 200:
            localizado = True
            status_download = "Bem sucedido"

            with open(pdf_local, "wb") as f:
                f.write(response.content)

            tamanho_pdf_kb = round(os.path.getsize(pdf_local) / 1024, 2)

            # Executa Busca 1: Remoção (com Gemini)
            vagas, sessao, validade, status_ia, tempo_proc = (
                extrair_dados_com_ia(pdf_local)
            )
            tem_vagas = bool(vagas)

            # Executa Busca 2: Violência Doméstica (exata no PyMuPDF)
            achados_vd = buscar_paragrafos_exatos(
                pdf_local, TERMO_VIOLENCIA_DOMESTICA
            )
            tem_vd = bool(achados_vd)

            tem_qualquer_resultado = tem_vagas or tem_vd

            destinatarios = (
                EMAILS_COM_RESULTADO
                if tem_qualquer_resultado
                else EMAILS_SEM_RESULTADO
            )
            arquivo_anexo_excel = None

            if tem_qualquer_resultado:
                formatar_excel(vagas, achados_vd, excel_local, data_exibicao)
                arquivo_anexo_excel = excel_local

            enviar_email(
                emails_destino=destinatarios,
                data_do=data_exibicao,
                url_pdf=url_pdf,
                localizado=localizado,
                status_dl=status_download,
                status_ia=status_ia,
                tem_vagas=tem_vagas,
                qtd_vagas=len(vagas),
                tempo_ia=tempo_proc,
                tamanho_kb=tamanho_pdf_kb,
                sessao_info=sessao,
                validade_info=validade,
                achados_termo=achados_vd,
                arquivo_excel=arquivo_anexo_excel,
                arquivo_pdf=pdf_local,
            )

        else:
            status_download = f"Mal sucedido (HTTP {response.status_code})"
            enviar_email(
                emails_destino=EMAILS_SEM_RESULTADO,
                data_do=data_exibicao,
                url_pdf=url_pdf,
                localizado=localizado,
                status_dl=status_download,
                status_ia=status_ia,
                tem_vagas=False,
            )

    except Exception as e:
        print(f"Erro crítico no fluxo: {e}")
        enviar_email(
            emails_destino=EMAILS_SEM_RESULTADO,
            data_do=data_exibicao,
            url_pdf=url_pdf,
            localizado=localizado,
            status_dl=f"Erro: {str(e)}",
            status_ia=status_ia,
            tem_vagas=False,
        )

    finally:
        for f in [pdf_local, excel_local]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except OSError:
                    pass


if __name__ == "__main__":
    rodar()
