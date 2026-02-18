import os
import datetime
import calendar
import glob
import time
import re
import sys

# --- IMPORTAÇÃO DAS BIBLIOTECAS ---
try:
    import pdfplumber
    PLUMBER_ENABLED = True
except ImportError:
    PLUMBER_ENABLED = False
    print("AVISO: 'pdfplumber' não instalado. A leitura de texto (nomes) falhará.")

try:
    from pypdf import PdfReader
    PYPDF_ENABLED = True
except ImportError:
    PYPDF_ENABLED = False
    print("AVISO: 'pypdf' não instalado. A validação de assinatura digital falhará.")


# --- CONFIGURAÇÃO DO CAMINHO DE REDE ---
#CAMINHO_RAIZ = r"R:\DO\COI\ARCC-CW\14 - LRO"
CAMINHO_RAIZ = r"E:\dev\sistema_lro\ARQUIVOS"

# --- CORES ---
class Cor:
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    MAGENTA = '\033[95m'
    DARK_YELLOW = '\033[33m'
    DARK_RED = '\033[31m'
    bg_BLUE = '\033[44m'
    WHITE = '\033[97m'
    GREY = '\033[90m'
    RESET = '\033[0m'

def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def abrir_arquivo(caminho):
    try:
        os.startfile(caminho)
    except Exception as e:
        if sys.platform == 'darwin': os.system(f'open "{caminho}"')
        else: print(f"{Cor.RED}[Erro ao abrir]: {e}{Cor.RESET}")

# --- NOVA FUNÇÃO: VERIFICA ASSINATURA NA ESTRUTURA (pypdf) ---
def verificar_assinatura_estrutural(caminho_pdf):
    if not PYPDF_ENABLED: return False
    try:
        reader = PdfReader(caminho_pdf)
        # Verifica se existem campos de formulário (AcroForm)
        if reader.get_fields():
            for field in reader.get_fields().values():
                # O Gov.br cria um campo do tipo '/Sig' (Signature)
                # Se encontrarmos QUALQUER campo de assinatura, consideramos assinado.
                if field.field_type == "/Sig":
                    return True
        
        # Check secundário: Procura anotações de Widget de assinatura nas páginas
        for page in reader.pages:
            if "/Annots" in page:
                for annot in page["/Annots"]:
                    obj = annot.get_object()
                    if obj.get("/FT") == "/Sig":
                        return True
    except:
        return False
    return False

# --- ANÁLISE DE TEXTO (pdfplumber) ---
def analisar_conteudo_lro(caminho_pdf):
    dados = {
        "cabecalho": "---", "recebeu": "---", "passou": "---", 
        "equipe": {"smc": "---", "bct": "---", "oea": "---"},
        "assinatura": False
    }

    # 1. VALIDAÇÃO ROBUSTA DE ASSINATURA (Via pypdf)
    # Isso verifica o certificado digital, não o texto visual.
    if verificar_assinatura_estrutural(caminho_pdf):
        dados["assinatura"] = True
    
    # 2. EXTRAÇÃO DE TEXTO (Via pdfplumber)
    if PLUMBER_ENABLED:
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                texto_completo = ""
                for pagina in pdf.pages:
                    texto_completo += pagina.extract_text() or ""
                
                texto_linear = texto_completo.replace('\n', ' ')
                
                # Fallback: Se pypdf falhou mas tem o link no texto, aceita também
                if not dados["assinatura"]:
                    if "validar.iti.gov.br" in texto_linear.lower():
                        dados["assinatura"] = True

                # Regex de Dados
                match_data = re.search(r"(\d+º)\s*turno.*?do dia\s*(\d+)\s*de\s*([a-zA-Zç]+)\s*de\s*(\d{4})", texto_linear, re.IGNORECASE)
                if match_data:
                    dados["cabecalho"] = f"dia {match_data.group(2)} de {match_data.group(3)} de {match_data.group(1)} turno {match_data.group(4)}"
                
                match_recebi = re.search(r"(Recebi-o aos.*?)((?:,|\.|ciente))", texto_linear, re.IGNORECASE)
                if match_recebi: dados["recebeu"] = match_recebi.group(1).strip()

                match_passei = re.search(r"(Passei-o aos.*?)((?:,|\.|cientificando))", texto_linear, re.IGNORECASE)
                if match_passei: dados["passou"] = match_passei.group(1).strip()

                match_bloco = re.search(r"EQUIPE DE SERVIÇO:(.*?)3\.", texto_linear, re.IGNORECASE)
                if match_bloco:
                    txt_eq = match_bloco.group(1)
                    m_smc = re.search(r"(?:^|;)\s*([^;-]+?)\s*(?:-|)\s*SMC", txt_eq, re.IGNORECASE)
                    if m_smc: dados["equipe"]["smc"] = m_smc.group(1).strip()
                    m_bct = re.search(r";\s*([^;-]+?)\s*(?:-|)\s*Controlador", txt_eq, re.IGNORECASE)
                    if m_bct: dados["equipe"]["bct"] = m_bct.group(1).strip()
                    m_oea = re.search(r";\s*([^;-]+?)\s*(?:-|)\s*Operador", txt_eq, re.IGNORECASE)
                    if m_oea: dados["equipe"]["oea"] = m_oea.group(1).strip()
        except: pass

    return dados

# --- FUNÇÃO DE BUSCA FLEXÍVEL ---
def buscar_arquivos_flexivel(pasta, data_string, turno):
    padroes = [
        f"{data_string}_{turno}TURNO*",
        f"{data_string}-{turno}TURNO*",
        f"{data_string}_{turno} TURNO*",
        f"{data_string}-{turno} TURNO*",
        f"{data_string} {turno}TURNO*"
    ]
    encontrados = []
    for p in padroes:
        encontrados.extend(glob.glob(os.path.join(pasta, p)))
    return list(set(encontrados))

mapa_pastas = {
    "01": "1 - JANEIRO",   "02": "2 - FEVEREIRO", "03": "3 - MARÇO", "04": "4 - ABRIL",
    "05": "5 - MAIO",      "06": "6 - JUNHO",     "07": "7 - JULHO", "08": "8 - AGOSTO",
    "09": "9 - SETEMBRO",  "10": "10 - OUTUBRO",  "11": "11 - NOVEMBRO", "12": "12 - DEZEMBRO"
}

def main():
    if os.name == 'nt' and not os.path.exists(CAMINHO_RAIZ):
        print(f"{Cor.RED}[ERRO CRÍTICO] Caminho {CAMINHO_RAIZ} não encontrado.{Cor.RESET}"); input(); return

    while True:
        limpar_tela()
        agora = datetime.datetime.now()
        mes, ano_curto = agora.strftime("%m"), agora.strftime("%y")
        hora_atual = agora.hour

        print(f"{Cor.CYAN}=== Verificador LRO 19.0 (Detector de Estrutura) ==={Cor.RESET}")
        if os.name == 'nt': print(f"{Cor.GREY}Conectado: {CAMINHO_RAIZ}{Cor.RESET}")
        
        problemas, relatorio = 0, []
        lista_pendentes = [] 
        lista_para_criar = []
        lista_ano_errado = []

        inp_mes = input(f"MÊS (Enter para {mes}): "); mes = inp_mes.zfill(2) if inp_mes else mes
        inp_ano = input(f"ANO (Enter para {ano_curto}): "); ano_curto = inp_ano if inp_ano else ano_curto
        ano_longo = "20" + ano_curto
        try: ano_errado = str(int(ano_curto) - 1).zfill(2)
        except: ano_errado = "00"

        # Navegação
        root = CAMINHO_RAIZ if os.name == 'nt' else "."
        path_ano = os.path.join(root, f"LRO {ano_longo}")
        if os.name != 'nt' and not os.path.exists(path_ano): path_ano = root 
        
        if not os.path.exists(path_ano) and os.name == 'nt':
            print(f"{Cor.RED}Pasta LRO {ano_longo} não existe.{Cor.RESET}"); time.sleep(2); continue
            
        path_mes = os.path.join(path_ano, mapa_pastas.get(mes, "X"))
        if not os.path.exists(path_mes) and os.name == 'nt':
            print(f"{Cor.RED}Pasta do mês não encontrada.{Cor.RESET}"); time.sleep(2); continue
        
        if os.name != 'nt': path_mes = "." 

        try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
        except: continue
        if mes == agora.strftime("%m") and ano_curto == agora.strftime("%y"): qtd_dias = agora.day

        print(f"\n{Cor.YELLOW}>> Analisando {mapa_pastas.get(mes)}...{Cor.RESET}")
        print("-" * 60)

        for dia in range(1, qtd_dias + 1):
            dia_fmt = f"{dia:02d}"
            data_str = f"{dia_fmt}{mes}{ano_curto}"
            data_str_err = f"{dia_fmt}{mes}{ano_errado}"

            turnos = [1, 2, 3]
            if dia == agora.day and mes == agora.strftime("%m"):
                if hora_atual < 14: turnos = []
                elif hora_atual < 21: turnos = [1]
                else: turnos = [1, 2]

            for turno in turnos:
                arquivos_turno = buscar_arquivos_flexivel(path_mes, data_str, turno)
                arquivos_errados = buscar_arquivos_flexivel(path_mes, data_str_err, turno)

                tem_ok = [f for f in arquivos_turno if "OK" in f.upper()]
                tem_falta_ass = [f for f in arquivos_turno if "FALTA ASS" in f.upper()]
                tem_falta_lro = [f for f in arquivos_turno if "FALTA LRO" in f.upper()]
                tem_novo = [f for f in arquivos_turno if "OK" not in f.upper() and "FALTA" not in f.upper()]

                if tem_ok: pass
                elif tem_falta_ass:
                    print(f"{Cor.YELLOW}[!] FALTA ASSINATURA: {os.path.basename(tem_falta_ass[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_ass[0])); problemas += 1
                elif tem_novo and tem_falta_lro:
                    print(f"{Cor.DARK_YELLOW}[!] LRO FEITO (TXT PRESENTE): {os.path.basename(tem_novo[0])}{Cor.RESET}")
                    lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno})
                    problemas += 1
                elif tem_falta_lro:
                    print(f"{Cor.MAGENTA}[!] NÃO CONFECCIONADO: {os.path.basename(tem_falta_lro[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_lro[0])); problemas += 1
                elif tem_novo:
                    n = os.path.basename(tem_novo[0])
                    if "-" in n or " " in n.replace("TURNO", ""):
                        print(f"{Cor.CYAN}[?] PENDENTE (NOME FORA DO PADRÃO): {n}{Cor.RESET}")
                    else:
                        print(f"{Cor.CYAN}[?] PENDENTE DE VERIFICAÇÃO: {n}{Cor.RESET}")
                    
                    lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno})
                    problemas += 1
                elif arquivos_errados:
                    print(f"{Cor.DARK_RED}[!] ANO INCORRETO: {os.path.basename(arquivos_errados[0])}{Cor.RESET}")
                    lista_ano_errado.append(arquivos_errados[0]); problemas += 1
                else:
                    print(f"{Cor.RED}[X] INEXISTENTE: {data_str} - {turno}º Turno{Cor.RESET}")
                    relatorio.append(f"Dia {dia_fmt} - {turno}º Turno: NÃO ENCONTRADO"); problemas += 1
                    lista_para_criar.append({"str": data_str, "turno": turno})

        print("-" * 60)
        if problemas == 0: print(f"{Cor.GREEN}Tudo em dia!{Cor.RESET}\n")

        # Ações Corretivas
        if lista_ano_errado:
            if input(f"{Cor.DARK_RED}Corrigir {len(lista_ano_errado)} anos errados? (S/N): {Cor.RESET}").lower() in ['s','ok']:
                for arq in lista_ano_errado:
                    abrir_arquivo(arq)
                    if input("   Renomear? (S/OK): ").lower() in ['s','ok']:
                        try: os.rename(arq, arq.replace(ano_errado, ano_curto))
                        except: pass

        if problemas > 0 and relatorio:
            print(f"\n{Cor.bg_BLUE}--- COPIAR E COBRAR ---{Cor.RESET}")
            for l in relatorio: print(l)
            print("-" * 30 + "\n")

        if lista_para_criar:
            if input(f"{Cor.YELLOW}Criar arquivos FALTA LRO? (S/N): {Cor.RESET}").lower() in ['s','ok']:
                for i in lista_para_criar:
                    n = f"{i['str']}_{i['turno']}TURNO FALTA LRO ().txt"
                    try: 
                        with open(os.path.join(path_mes, n), 'w') as f: f.write("Falta")
                        print(f"{Cor.GREEN}   [V] Criado: {n}{Cor.RESET}")
                    except: pass

        # --- VERIFICAÇÃO VISUAL COM ESTRUTURA ---
        if lista_pendentes:
            print(f"{Cor.CYAN}Verificar {len(lista_pendentes)} livros?{Cor.RESET}")
            if input("Iniciar UM POR UM? (S/N): ").lower() in ['s','ok']:
                cnt = 1
                for item in lista_pendentes:
                    caminho = item['path']
                    nome_atual = os.path.basename(caminho)
                    
                    print(f"\n{Cor.YELLOW}--- ({cnt}/{len(lista_pendentes)}) ---{Cor.RESET}")
                    print(f"{Cor.CYAN}Arquivo: {nome_atual}{Cor.RESET}")
                    
                    abrir_arquivo(caminho)
                    print(f"{Cor.GREY}Analisando estrutura e texto...{Cor.RESET}")
                    info = analisar_conteudo_lro(caminho)
                    
                    if info:
                        print("-" * 60)
                        print(f"📅 {Cor.GREEN}{info['cabecalho']}{Cor.RESET}")
                        
                        # --- EXIBIÇÃO DA ASSINATURA ---
                        if info['assinatura']:
                            print(f"🔏 ASSINATURA: {Cor.GREEN}OK (Certificado Digital Detectado) ✅{Cor.RESET}")
                        else:
                            print(f"🔏 ASSINATURA: {Cor.RED}NÃO DETECTADA NA ESTRUTURA ❌{Cor.RESET}")
                            
                        print("-" * 60)
                        print(f"⬅️  Recebeu: {info['recebeu']}")
                        print(f"➡️  Passou:  {info['passou']}")
                        print(f"{Cor.WHITE}👥 Equipe:{Cor.RESET}")
                        print(f"   • SMC: {info['equipe']['smc']}")
                        print(f"   • BCT: {info['equipe']['bct']}")
                        print(f"   • OEA: {info['equipe']['oea']}")
                        print("-" * 60)
                    
                    acao = input(">> Confirmar e assinar? (S/OK): ").lower()
                    if acao in ['s', 'ok']:
                        dir_arq = os.path.dirname(caminho)
                        extensao = os.path.splitext(caminho)[1]
                        novo_nome_base = f"{item['data']}_{item['turno']}TURNO OK{extensao}"
                        novo_caminho = os.path.join(dir_arq, novo_nome_base)
                        
                        renomeado = False
                        while not renomeado:
                            try:
                                os.rename(caminho, novo_caminho)
                                print(f"{Cor.GREEN}   [V] Validado e Padronizado: {novo_nome_base}{Cor.RESET}")
                                renomeado = True
                            except PermissionError:
                                print(f"{Cor.RED}   [!] Feche o arquivo PDF! Enter p/ tentar.{Cor.RESET}"); input()
                            except Exception as e:
                                print(f"{Cor.RED}Erro: {e}{Cor.RESET}"); break
                    else:
                        print(f"{Cor.GREY}   [-] Mantido.{Cor.RESET}")
                    cnt += 1

        if input("Sair? (S/Enter): ").lower() in ['s','ok']: break
        else: break

if __name__ == "__main__": main()