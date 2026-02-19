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
    
    correcoes = {
        "JACKSON": "JAKSON",
        "JACSON": "JAKSON",
        "GIOVANI": "GIOVANNI",
        "IS ": "1S ",  # Corrige erro de digitação de I no lugar de 1
        "I S ": "1S ",
        "CP ": "CAP "  # Corrige CP FÁBIO CÉSAR
    }
    for errado, certo in correcoes.items():
        texto = texto.replace(errado, certo)
        
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

def extrair_dados_texto(texto_linear, dados):
    match_data = re.search(r"(\d+º)\s*turno.*?do dia\s*(\d+)\s*de\s*([a-zA-Zç]+)\s*de\s*(\d{4})", texto_linear, re.IGNORECASE)
    if match_data:
        dados["cabecalho"] = f"dia {match_data.group(2)} de {match_data.group(3)} de {match_data.group(1)} turno {match_data.group(4)}"
    
    match_recebi = re.search(r"(Recebi-o\s+(?:aos|às|as).*?)((?:,|\.|ciente))", texto_linear, re.IGNORECASE)
    if match_recebi: dados["recebeu"] = match_recebi.group(1).strip()

    match_passei = re.search(r"(Passei-o\s+(?:aos|às|as).*?)((?:,|\.|cientificando))", texto_linear, re.IGNORECASE)
    if match_passei: dados["passou"] = match_passei.group(1).strip()

    match_resp_gov = re.search(r"validar\.iti\.gov\.br\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ a-z]+?)\s*-", texto_linear, re.IGNORECASE)
    if match_resp_gov: dados["responsavel"] = match_resp_gov.group(1).strip().upper()
    else:
        match_resp_txt = re.search(r"ordens em vigor\.\s*([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ a-z]+?)\s*-", texto_linear, re.IGNORECASE)
        if match_resp_txt: dados["responsavel"] = match_resp_txt.group(1).strip().upper()

    match_bloco = re.search(r"EQUIPE DE SERVI[CÇ]O:(.*?)3\.", texto_linear, re.IGNORECASE)
    if match_bloco:
        txt_eq = match_bloco.group(1)
        dados["texto_equipe"] = txt_eq # Guardar APENAS este bloco para o fallback
        
        # Regex atualizada: Aceita [;,] (Vírgula ou Ponto-e-vírgula) e não quebra com espaços
        m_smc = re.search(r"(?:^|[;,])\s*(.*?)\s*(?:-|)\s*SMC", txt_eq, re.IGNORECASE)
        if m_smc: dados["equipe"]["smc"] = m_smc.group(1).strip()
        
        m_bct = re.search(r"[;,]\s*(.*?)\s*(?:-|)\s*Controlador", txt_eq, re.IGNORECASE)
        if m_bct: dados["equipe"]["bct"] = m_bct.group(1).strip()
        
        m_oea = re.search(r"[;,]\s*(.*?)\s*(?:-|)\s*Operador", txt_eq, re.IGNORECASE)
        if m_oea: dados["equipe"]["oea"] = m_oea.group(1).strip()

def analisar_conteudo_lro(caminho_pdf):
    dados = {
        "cabecalho": "---", "responsavel": "---", "recebeu": "---", "passou": "---", 
        "equipe": {"smc": "---", "bct": "---", "oea": "---"},
        "assinatura": verificar_assinatura_estrutural(caminho_pdf),
        "texto_completo": "",
        "texto_equipe": "" # Campo novo!
    }
    if PLUMBER_ENABLED:
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                texto_completo = "".join([pagina.extract_text() or "" for pagina in pdf.pages])
                texto_linear = texto_completo.replace('\n', ' ')
                
                dados["texto_completo"] = texto_linear
                
                if not dados["assinatura"] and "validar.iti.gov.br" in texto_linear.lower():
                    dados["assinatura"] = True
                extrair_dados_texto(texto_linear, dados)
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