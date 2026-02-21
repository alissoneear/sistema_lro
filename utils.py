import os
import sys
import glob
import re
import unicodedata

from config import Config, Cor

try:
    import pdfplumber
    PLUMBER_ENABLED = True
except ImportError:
    PLUMBER_ENABLED = False

try:
    from pypdf import PdfReader
    PYPDF_ENABLED = True
except ImportError:
    PYPDF_ENABLED = False

try:
    import msvcrt
    MSVCRT_ENABLED = True
except ImportError:
    MSVCRT_ENABLED = False

def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def pedir_confirmacao(mensagem):
    print(mensagem, end='', flush=True)
    if os.name == 'nt' and MSVCRT_ENABLED:
        while True:
            tecla = msvcrt.getch()
            if tecla in [b's', b'S']:
                print("S")
                return True
            elif tecla == b'\r':
                print("Enter")
                return True
            elif tecla == b'\x1b':
                print("ESC")
                return False
            elif tecla == b'\x03':
                print("^C")
                raise KeyboardInterrupt
    else:
        try:
            resp = input().strip().upper()
            if resp == 'ESC': return False
            return True
        except EOFError:
            raise KeyboardInterrupt

def abrir_arquivo(caminho):
    try:
        if os.name == 'nt': os.startfile(caminho)
        elif sys.platform == 'darwin': os.system(f'open "{caminho}"')
        else: os.system(f'xdg-open "{caminho}"')
    except Exception as e:
        print(f"{Cor.RED}[Erro ao abrir]: {e}{Cor.RESET}")


def normalizar_texto(texto):
    if not texto: return "---"
    texto = str(texto).upper()
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('utf-8')
    
    # Uso de Regex (\b) para garantir que só troca "IS" se for palavra solta (não no meio de REGIS)
    correcoes = {
        r"\bJACKSON\b": "JAKSON",
        r"\bJACSON\b": "JAKSON",
        r"\bGIOVANI\b": "GIOVANNI",
        r"\bIS\b": "1S",  
        r"\bI S\b": "1S",
        r"\bCP\b": "CAP",
        r"\bD\.?\s*CASTRO\b": "DE CASTRO"
    }
    for errado, certo in correcoes.items():
        texto = re.sub(errado, certo, texto)
        
    return texto.strip()

def extrair_nome_base(nome_guerra):
    nome_norm = normalizar_texto(nome_guerra)
    partes = nome_norm.split()
    patentes = ['1S', '2S', '3S', 'SO', 'TEN', '1T', '2T', 'CAP', 'MAJ', 'CEL', 'SGT']
    if len(partes) > 1 and partes[0] in patentes:
        return " ".join(partes[1:])
    return nome_norm

def encontrar_legenda(nome_extraido, dicionario_mapa):
    nome_ext = normalizar_texto(nome_extraido)
    if nome_ext in ["---", ""]: return "---"
    
    for nome_guerra, dados in dicionario_mapa.items():
        nome_base = extrair_nome_base(nome_guerra)
        if nome_base in nome_ext or nome_ext in nome_base:
            return dados["legenda"]
    return "???"

def encontrar_legenda_fallback(texto_alvo, dicionario_mapa):
    texto_norm = normalizar_texto(texto_alvo)
    for nome_guerra, dados in dicionario_mapa.items():
        nome_base = extrair_nome_base(nome_guerra)
        if nome_base in texto_norm:
            return dados["legenda"]
    return "???"

def verificar_assinatura_estrutural(caminho_pdf):
    if not PYPDF_ENABLED: return False
    try:
        reader = PdfReader(caminho_pdf)
        if reader.get_fields():
            for field in reader.get_fields().values():
                if field.field_type == "/Sig": return True
        for page in reader.pages:
            if "/Annots" in page:
                for annot in page["/Annots"]:
                    obj = annot.get_object()
                    if obj.get("/FT") == "/Sig": return True
    except Exception: pass
    return False

def busca_inteligente_equipe(texto_bloco, mapa_efetivo):
    """Varre um bloco de texto procurando militares do dicionário de forma robusta."""
    texto_norm = normalizar_texto(texto_bloco)
    
    # Criar lista de nomes base (ex: 'GIOVANNI') ordenada pelo tamanho (maiores primeiro)
    nomes_base = [(extrair_nome_base(ng), ng) for ng in mapa_efetivo.keys()]
    nomes_base.sort(key=lambda x: len(x[0]), reverse=True)
    
    for nome_base, nome_completo in nomes_base:
        if re.search(rf'\b{re.escape(nome_base)}\b', texto_norm):
            # Retorna apenas o Posto e Nome (ex: '1S GIOVANNI')
            partes = nome_completo.split('-')
            return partes[1].strip() if len(partes) > 1 else nome_completo
    return "---"

def extrair_dados_texto(texto_linear, dados, mes=None, ano_curto=None):
    # Cabeçalho
    match_data = re.search(r"(\d+º)\s*turno.*?do dia\s*(\d+)\s*de\s*([a-zA-Zç]+)\s*de\s*(\d{4})", texto_linear, re.IGNORECASE)
    if match_data:
        dados["cabecalho"] = f"Dia {match_data.group(2)} de {match_data.group(3)} de {match_data.group(1)} turno {match_data.group(4)}"
    
    # Recebeu/Passou
    match_bloco_recebeu = re.search(r"RECEBIMENTO DO SERVI[CÇ]O:(.*?)(?:2\.\s*EQUIPE|EQUIPE DE SERVI[CÇ]O)", texto_linear, re.IGNORECASE)
    raw_recebeu = match_bloco_recebeu.group(1).strip() if match_bloco_recebeu else "---"
    
    match_bloco_passou = re.search(r"PASSAGEM DE SERVI[CÇ]O:(.*?)(?:gov\.br|Documento assinado|Asas que|$)", texto_linear, re.IGNORECASE)
    raw_passou = match_bloco_passou.group(1).strip() if match_bloco_passou else "---"

    dados["recebeu"] = re.sub(r",?\s*(?:ciente|cientificando).*$", "", raw_recebeu, flags=re.IGNORECASE).strip()
    dados["passou"] = re.sub(r",?\s*(?:ciente|cientificando).*$", "", raw_passou, flags=re.IGNORECASE).strip()


    # RESPONSÁVEL PELA ASSINATURA (Busca e Padronização de Nome Completo)
    match_gov_limpo = re.search(r"assinado digitalmente\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ a-z]+?)\s+Data", texto_linear, re.IGNORECASE)
    match_gov_link = re.search(r"validar\.iti\.gov\.br\.?\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ a-z]+?)\s*(?:-|\d[ST]|CAP|MAJ|SO|TEN)", texto_linear, re.IGNORECASE)
    match_txt = re.search(r"ordens em vigor\.?\s*([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ a-z]{5,50}?)\s*(?:-|\d[ST]|CAP|MAJ|SO|TEN)", texto_linear, re.IGNORECASE)

    responsavel_bruto = None
    if match_gov_limpo and len(match_gov_limpo.group(1)) < 60:
        responsavel_bruto = match_gov_limpo.group(1).strip().upper()
    elif match_gov_link and len(match_gov_link.group(1)) < 60:
        responsavel_bruto = match_gov_link.group(1).strip().upper()
    elif match_txt and not re.search(r"gov\.br|assinado", match_txt.group(1), re.IGNORECASE):
        responsavel_bruto = match_txt.group(1).strip().upper()

    # --- Lógica de Padronização Oficial ---
    from config import DadosEfetivo
    chave = f"{mes}{ano_curto}" if mes and ano_curto else None
    dados_hist = DadosEfetivo.HISTORICO.get(chave) if chave else None
    
    nomes_oea = dados_hist["oea"] if dados_hist else DadosEfetivo.nomes_oea
    nomes_bct = dados_hist["bct"] if dados_hist else DadosEfetivo.nomes_bct
    nomes_smc = DadosEfetivo.nomes_smc
    todas_linhas_efetivo = nomes_oea + nomes_bct + nomes_smc

    def obter_nome_oficial(texto_alvo):
        """Pesquisa o nome base no dicionário e retorna a string oficial completa."""
        if not texto_alvo: return None
        texto_norm = normalizar_texto(texto_alvo)
        
        lista_buscas = []
        for linha in todas_linhas_efetivo:
            partes = [p.strip() for p in linha.split('-')]
            nome_oficial = partes[0] # Ex: "2S BCO KEVIN SOUZA DE OLIVEIRA"
            nome_guerra = partes[1] if len(partes) > 1 else nome_oficial
            nome_base = extrair_nome_base(nome_guerra) # Ex: "KEVIN"
            lista_buscas.append((nome_base, nome_oficial))
        
        # Ordena para garantir que não haja falsos positivos (nomes contidos em outros)
        lista_buscas.sort(key=lambda x: len(x[0]), reverse=True)
        
        for nome_base, nome_oficial in lista_buscas:
            if re.search(rf'\b{re.escape(nome_base)}\b', texto_norm):
                return nome_oficial
        return None

    # Aplicação da padronização
    if responsavel_bruto:
        nome_encontrado = obter_nome_oficial(responsavel_bruto)
        dados["responsavel"] = nome_encontrado if nome_encontrado else responsavel_bruto
    else:
        # Fallback Inteligente: varre o final do documento e devolve o nome completo
        texto_final = normalizar_texto(texto_linear[-300:])
        nome_encontrado = obter_nome_oficial(texto_final)
        dados["responsavel"] = nome_encontrado if nome_encontrado else "---"


# Equipe (Agora totalmente automatizada e inteligente)
    match_bloco_eq = re.search(r"EQUIPE DE SERVI[CÇ]O:(.*?)3\.", texto_linear, re.IGNORECASE)
    if match_bloco_eq:
        txt_eq = match_bloco_eq.group(1)
        dados["texto_equipe"] = txt_eq 
        
        from config import DadosEfetivo
        # Passar o mês e ano para buscar o histórico correto!
        m_smc, m_bct, m_oea = DadosEfetivo.mapear_efetivo(mes, ano_curto)
        
        dados["equipe"]["smc"] = busca_inteligente_equipe(txt_eq, m_smc)
        dados["equipe"]["bct"] = busca_inteligente_equipe(txt_eq, m_bct)
        dados["equipe"]["oea"] = busca_inteligente_equipe(txt_eq, m_oea)

def analisar_conteudo_lro(caminho_pdf, mes=None, ano_curto=None):
    dados = {
        "cabecalho": "---", "responsavel": "---", "recebeu": "---", "passou": "---", 
        "equipe": {"smc": "---", "bct": "---", "oea": "---"},
        "assinatura": verificar_assinatura_estrutural(caminho_pdf),
        "texto_completo": "",
        "texto_equipe": "" 
    }
    if PLUMBER_ENABLED:
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                texto_completo = "".join([pagina.extract_text() or "" for pagina in pdf.pages])
                texto_linear = texto_completo.replace('\n', ' ')
                dados["texto_completo"] = texto_linear
                if not dados["assinatura"] and "validar.iti.gov.br" in texto_linear.lower():
                    dados["assinatura"] = True
                # Repassar mes e ano aqui:
                extrair_dados_texto(texto_linear, dados, mes, ano_curto)
        except Exception: pass
    return dados

def buscar_arquivos_flexivel(pasta, data_string, turno):
    padroes = [
        f"{data_string}_{turno}TURNO*", f"{data_string}-{turno}TURNO*",
        f"{data_string}_{turno} TURNO*", f"{data_string}-{turno} TURNO*",
        f"{data_string} {turno}TURNO*"
    ]
    encontrados = []
    for p in padroes: encontrados.extend(glob.glob(os.path.join(pasta, p)))
    return list(set(encontrados))

def calcular_turnos_validos(dia, mes_str, dia_atual, mes_atual, hora_atual):
    turnos = [1, 2, 3]
    if dia == dia_atual and mes_str == mes_atual:
        if hora_atual < 14: turnos = []
        elif hora_atual < 21: turnos = [1]
        else: turnos = [1, 2]
    return turnos