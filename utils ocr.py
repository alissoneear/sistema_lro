import os
import sys
import glob
import re
import unicodedata
import numpy as np
import pdfplumber
from config import Config, Cor

# --- DEPENDÊNCIAS OPCIONAIS ---
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

# --- LÓGICA EASYOCR ---
try:
    import easyocr
    from pdf2image import convert_from_path
    EASYOCR_ENABLED = True
    READER = None 
except ImportError:
    EASYOCR_ENABLED = False

def obter_reader():
    """Inicializa o EasyOCR apenas uma vez para ganhar velocidade."""
    global READER
    if READER is None and EASYOCR_ENABLED:
        # gpu=False para garantir funcionamento em PCs sem placa de vídeo
        READER = easyocr.Reader(['pt'], gpu=False) 
    return READER

def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def pedir_confirmacao(mensagem):
    print(mensagem, end='', flush=True)
    if os.name == 'nt' and MSVCRT_ENABLED:
        while True:
            tecla = msvcrt.getch()
            if tecla in [b's', b'S', b'\r']: return True
            elif tecla in [b'\x1b', b'n', b'N']: return False
            elif tecla == b'\x03': raise KeyboardInterrupt
    else:
        try:
            resp = input().strip().upper()
            return resp != 'ESC'
        except (EOFError, KeyboardInterrupt): raise KeyboardInterrupt

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
    
    correcoes = {
        r"\bIS\b": "1S", r"\bI S\b": "1S", r"\bCP\b": "CAP",
        r"\bJACKSON\b": "JAKSON", r"\bGIOVANI\b": "GIOVANNI",
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
    texto_norm = normalizar_texto(texto_bloco)
    nomes_base = [(extrair_nome_base(ng), ng) for ng in mapa_efetivo.keys()]
    nomes_base.sort(key=lambda x: len(x[0]), reverse=True)
    
    for nome_base, nome_completo in nomes_base:
        if re.search(rf'\b{re.escape(nome_base)}\b', texto_norm):
            partes = nome_completo.split('-')
            return partes[1].strip() if len(partes) > 1 else nome_completo
    return "---"

def extrair_dados_texto_flexivel(texto_linear, dados, mes=None, ano_curto=None):
    """Nova lógica robusta para os LROs de 2026."""
    from config import DadosEfetivo
    m_smc, m_bct, m_oea = DadosEfetivo.mapear_efetivo(mes, ano_curto)

    # 1. CABEÇALHO (Regex Elástico)
    m_data = re.search(r"(\d+º)\s*turno.*?dia\s*(\d+)\s*de\s*([a-zA-Zç]+)\s*de\s*(\d{4})", texto_linear, re.IGNORECASE)
    if m_data:
        dados["cabecalho"] = f"Dia {m_data.group(2)} de {m_data.group(3)} ({m_data.group(1)} turno)"

    # 2. EQUIPE (Captura tudo entre o item 2. e o item 3.)
    m_bloco_eq = re.search(r"2\.\s*(.*?)\s*3\.", texto_linear, re.IGNORECASE | re.DOTALL)
    if m_bloco_eq:
        bloco = m_bloco_eq.group(1)
        dados["equipe"]["smc"] = busca_inteligente_equipe(bloco, m_smc)
        dados["equipe"]["bct"] = busca_inteligente_equipe(bloco, m_bct)
        dados["equipe"]["oea"] = busca_inteligente_equipe(bloco, m_oea)

    # 3. RESPONSÁVEL (Nome antes da Data de assinatura)
    m_resp = re.search(r"([A-ZÃÁÉÍÓÚ ]+?)\s+Data: \d{2}/\d{2}/\d{4}", texto_linear)
    if m_resp:
        dados["responsavel"] = m_resp.group(1).strip()

    # 4. RECEBIMENTO/PASSAGEM
    m_rec = re.search(r"RECEBIMENTO.*?do\s+([^,.-]+)", texto_linear, re.IGNORECASE)
    m_pas = re.search(r"PASSAGEM.*?ao\s+([^,.-]+)", texto_linear, re.IGNORECASE)
    dados["recebeu"] = m_rec.group(1).strip() if m_rec else "---"
    dados["passou"] = m_pas.group(1).strip() if m_pas else "---"

def analisar_conteudo_lro(caminho_pdf, mes=None, ano_curto=None):
    dados = {
        "cabecalho": "---", "responsavel": "---", "recebeu": "---", "passou": "---", 
        "equipe": {"smc": "---", "bct": "---", "oea": "---"},
        "assinatura": verificar_assinatura_estrutural(caminho_pdf),
        "texto_completo": ""
    }

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = "".join([p.extract_text() or "" for p in pdf.pages])
            
        if not texto.strip() and EASYOCR_ENABLED:
            print(f"{Cor.YELLOW}   [!] PDF de imagem detectado. Iniciando EasyOCR...{Cor.RESET}")
            reader = obter_reader()
            paginas_img = convert_from_path(caminho_pdf)
            texto = ""
            for img in paginas_img:
                texto += " ".join(reader.readtext(np.array(img), detail=0)) + " "

        texto_linear = texto.replace('\n', ' ')
        dados["texto_completo"] = texto_linear
        
        if texto_linear.strip():
            if "validar.iti.gov.br" in texto_linear.lower(): dados["assinatura"] = True
            extrair_dados_texto_flexivel(texto_linear, dados, mes, ano_curto)

    except Exception as e:
        print(f"{Cor.RED}Erro na análise: {e}{Cor.RESET}")
    return dados

def buscar_arquivos_flexivel(pasta, data_string, turno):
    padroes = [f"{data_string}_{turno}TURNO*", f"{data_string}-{turno}TURNO*", f"{data_string} {turno}TURNO*"]
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