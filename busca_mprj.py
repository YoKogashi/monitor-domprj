import requests
from bs4 import BeautifulSoup

# O link que você me passou
URL_BUSCA = "https://www.mprj.mp.br/busca?p_p_id=br_mp_mprj_internet_busca_web_BuscaPortlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_count=1&_br_mp_mprj_internet_busca_web_BuscaPortlet_jspPage=%2Fhtml%2Fview.jsp&_br_mp_mprj_internet_busca_web_BuscaPortlet_exibicao_param=card&_br_mp_mprj_internet_busca_web_BuscaPortlet_filtro_param=doerj&_br_mp_mprj_internet_busca_web_BuscaPortlet_delta=15&_br_mp_mprj_internet_busca_web_BuscaPortlet_keywords=&_br_mp_mprj_internet_busca_web_BuscaPortlet_advancedSearch=false&_br_mp_mprj_internet_busca_web_BuscaPortlet_andOperator=true&_br_mp_mprj_internet_busca_web_BuscaPortlet_resetCur=false&_br_mp_mprj_internet_busca_web_BuscaPortlet_cur=1"

def buscar():
    print("Iniciando busca diária no MPRJ...")
    resposta = requests.get(URL_BUSCA)
    soup = BeautifulSoup(resposta.text, 'html.parser')
    
    # Busca por links de PDF (geralmente terminam em .pdf ou contêm 'download')
    links = soup.find_all('a', href=True)
    pdf_links = [l['href'] for l in links if 'pdf' in l['href'].lower()][:2]
    
    if not pdf_links:
        print("Nenhum Diário Oficial novo encontrado hoje.")
        return

    print(f"Foram encontrados {len(pdf_links)} editais recentes.")
    for link in pdf_links:
        print(f"Link para conferir: {link}")

if __name__ == "__main__":
    buscar()
