import os
import sys
import datetime
import calendar
import glob
import time
import re

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

try:
    import msvcrt
    MSVCRT_ENABLED = True
except ImportError:
    MSVCRT_ENABLED = False

# --- CONFIGURAÇÃO ---
class Config:
    CAMINHO_RAIZ = r"E:\dev\sistema_lro\ARQUIVOS" if os.name == 'nt' else "."
    
    MAPA_PASTAS = {
        "01": "1 - JANEIRO",   "02": "2 - FEVEREIRO", "03": "3 - MARÇO", "04": "4 - ABRIL",
        "05": "5 - MAIO",      "06": "6 - JUNHO",     "07": "7 - JULHO", "08": "8 - AGOSTO",
        "09": "9 - SETEMBRO",  "10": "10 - OUTUBRO",  "11": "11 - NOVEMBRO", "12": "12 - DEZEMBRO"
    }
    
    MAPA_SEMANA = {0: 'SEG', 1: 'TER', 2: 'QUA', 3: 'QUI', 4: 'SEX', 5: 'SÁB', 6: 'DOM'}

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
    ORANGE = '\033[38;5;208m'

# --- DADOS DO EFETIVO ---
class DadosEfetivo:
    legendas_oea = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    legendas_bct = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
    legendas_smc = ['GEA', 'EVT', 'FCS', 'MAL', 'BAR', 'ESC', 'CPL', 'JAK', 'CER']

    nomes_oea = [
        '1S QSS BCO LUIZ FILIPE TERBECK - 1S TERBECK - 6032338', 
        '1S BCO RUI ANTONIO DOS SANTOS JUNIOR - 1S RUI - 6032311', 
        '2S BCO SAULO JULIO SANTOS GONÇALVES - 2S SAULO - 6134262', 
        '2S BCO KEVIN SOUZA DE OLIVEIRA - 2S KEVIN - 6157912', 
        '2S BCO ALAN SITORSKI - 2S SITORSKI - 6156568', 
        '2S BCO LUCIANO FERNANDES DOS SANTOS - 2S FERNANDES - 6156878', 
        'SO BCO FRANCISCO TORRES SOARES - SO SOARES - 3502201', 
        '2S BCO ELIAS EDVANEIDSON DE ARAUJO CARDOSO - 2S ELIAS - 6328512', 
        '2S BCO ALISSON LOURENÇO DA SILVA - 2S LOURENÇO - 6780245'
    ]

    nomes_bct = [
        'SO BCT MARCELO DE SOUZA PEREIRA - SO SOUZA - 3988554', 
        '1S BCT DIOGO FAVARO WUNDERLICH - 1S FAVARO - 4202317', 
        '1S BCT RÉGIS FERRARI - 1S RÉGIS - 4378717', 
        '1S BCT CLEON FRAGA DOS SANTOS - 1S CLEON - 4378660', 
        '1S BCT MARCELO ZANOTTO GONÇALVES - 1S ZANOTTO - 6100775', 
        '1S BCT GIOVANNI COCCO ALVES - 1S GIOVANNI - 6158080', 
        '2S BCT YURI DE PAIVA WERNECK - 2S WERNECK - 6301371', 
        '2S BCT EDSON DE SOUZA NUNES - 2S EDSON - 6750451', 
        '1S BCT GUSTAVO FRACAO LAGO - 1S GUSTAVO - 4378482', 
        '1S BCT CRISTIANO PAZ PRATES - 1S PRATES - 4378768', 
        '1S BCT EVANDRO MACHADO BITTENCOURT - 1S BITTENCOURT - 6087841', 
        '2S BCT LIS CECI LYRA FONTES - 2S LIS - 6301118', 
        '2S BCT EDUARDO OLIVEIRA DE CASTRO - 2S DE CASTRO - 6576044'
    ]

    nomes_smc = [
        'MAJ AV THIAGO GEALH DE CAMPOS - MAJ GEALH', 
        'MAJ QOECTA EVERTON RIBEIRO DA COSTA - MAJ EVERTON', 
        'CAP QOECTA FÁBIO CÉSAR SILVA DE OLIVEIRA - CAP FÁBIO CÉSAR', 
        'CAP QOECTA JOSÉ GUILHERME MALTA - CAP MALTA', 
        'CAP QOECTA BARBARA PACHECO LINS - CAP BARBARA', 
        'CAP EDGAR HENRIQUE ESCOBAR DOS SANTOS - CAP ESCOBAR', 
        '1T AV DANIEL BUERY DE MELO CAMPELO - 1T CAMPELO', 
        '1T QOECTA JAKSON DA SILVA - 1T JAKSON', 
        '1T QOECTA JULIMAR CERUTTI DA SILVA - 1T CERUTTI'
    ]

    @staticmethod
    def mapear_efetivo():
        mapa_smc, mapa_bct, mapa_oea = {}, {}, {}
        for i, linha in enumerate(DadosEfetivo.nomes_oea):
            partes = [p.strip() for p in linha.split('-')]
            if len(partes) >= 3: mapa_oea[partes[1].upper()] = {"legenda": DadosEfetivo.legendas_oea[i], "saram": partes[2]}
        for i, linha in enumerate(DadosEfetivo.nomes_bct):
            partes = [p.strip() for p in linha.split('-')]
            if len(partes) >= 3: mapa_bct[partes[1].upper()] = {"legenda": DadosEfetivo.legendas_bct[i], "saram": partes[2]}
        for i, linha in enumerate(DadosEfetivo.nomes_smc):
            partes = [p.strip() for p in linha.split('-')]
            if len(partes) >= 2: mapa_smc[partes[1].upper()] = {"legenda": DadosEfetivo.legendas_smc[i], "saram": None}
        return mapa_smc, mapa_bct, mapa_oea


# --- UTILITÁRIOS ---
def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def pedir_confirmacao(mensagem):
    """Lê a tecla 'S', 'Enter' ou 'ESC' instantaneamente (no Windows)."""
    print(mensagem, end='', flush=True)
    if os.name == 'nt' and MSVCRT_ENABLED:
        while True:
            tecla = msvcrt.getch()
            if tecla in [b's', b'S']:
                print("S")
                return True
            elif tecla == b'\r':  # Tecla Enter
                print("Enter")
                return True
            elif tecla == b'\x1b': # Tecla ESC
                print("ESC")
                return False
    else:
        # Fallback caso não seja Windows
        resp = input().strip().upper()
        if resp == 'ESC': return False
        return True

def abrir_arquivo(caminho):
    try:
        if os.name == 'nt': os.startfile(caminho)
        elif sys.platform == 'darwin': os.system(f'open "{caminho}"')
        else: os.system(f'xdg-open "{caminho}"')
    except Exception as e:
        print(f"{Cor.RED}[Erro ao abrir]: {e}{Cor.RESET}")

def encontrar_legenda(nome_extraido, dicionario_mapa):
    nome_ext = nome_extraido.upper().strip()
    if nome_ext in ["---", ""]: return "---"
    for nome_guerra, dados in dicionario_mapa.items():
        if nome_guerra in nome_ext or nome_ext in nome_guerra:
            return dados["legenda"]
    return "???"


# --- PROCESSAMENTO DE PDF ---
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

def analisar_conteudo_lro(caminho_pdf):
    dados = {
        "cabecalho": "---", "responsavel": "---", "recebeu": "---", "passou": "---", 
        "equipe": {"smc": "---", "bct": "---", "oea": "---"},
        "assinatura": verificar_assinatura_estrutural(caminho_pdf)
    }
    if PLUMBER_ENABLED:
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                texto_completo = "".join([pagina.extract_text() or "" for pagina in pdf.pages])
                texto_linear = texto_completo.replace('\n', ' ')
                
                if not dados["assinatura"] and "validar.iti.gov.br" in texto_linear.lower():
                    dados["assinatura"] = True
                extrair_dados_texto(texto_linear, dados)
        except Exception: pass
    return dados

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

    match_bloco = re.search(r"EQUIPE DE SERVIÇO:(.*?)3\.", texto_linear, re.IGNORECASE)
    if match_bloco:
        txt_eq = match_bloco.group(1)
        m_smc = re.search(r"(?:^|;)\s*([^;-]+?)\s*(?:-|)\s*SMC", txt_eq, re.IGNORECASE)
        if m_smc: dados["equipe"]["smc"] = m_smc.group(1).strip()
        m_bct = re.search(r";\s*([^;-]+?)\s*(?:-|)\s*Controlador", txt_eq, re.IGNORECASE)
        if m_bct: dados["equipe"]["bct"] = m_bct.group(1).strip()
        m_oea = re.search(r";\s*([^;-]+?)\s*(?:-|)\s*Operador", txt_eq, re.IGNORECASE)
        if m_oea: dados["equipe"]["oea"] = m_oea.group(1).strip()

# --- LÓGICA DE FICHEIROS ---
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

def corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado):
    if not lista_ano_errado: return
    if pedir_confirmacao(f"{Cor.DARK_RED}Corrigir {len(lista_ano_errado)} anos errados? (S/Enter p/ Sim, ESC p/ Pular): {Cor.RESET}"):
        for arq in lista_ano_errado:
            abrir_arquivo(arq)
            if pedir_confirmacao(f"   Renomear {os.path.basename(arq)}? (S/Enter p/ Sim, ESC p/ Não): "):
                try: 
                    os.rename(arq, arq.replace(ano_errado, ano_curto))
                    print(f"{Cor.GREEN}   Renomeado com sucesso.{Cor.RESET}")
                except Exception as e: print(f"{Cor.RED}   Erro ao renomear: {e}{Cor.RESET}")

def processo_verificacao_visual(lista_pendentes):
    if not lista_pendentes: return
    print(f"{Cor.CYAN}Verificar {len(lista_pendentes)} livros pendentes?{Cor.RESET}")
    if not pedir_confirmacao("Iniciar UM POR UM? (S/Enter p/ Sim, ESC p/ Pular): "): return

    for cnt, item in enumerate(lista_pendentes, start=1):
        caminho, data_str, turno = item['path'], item['data'], item['turno']
        nome_atual = os.path.basename(caminho)
        print(f"\n{Cor.YELLOW}--- ({cnt}/{len(lista_pendentes)}) ---{Cor.RESET}")
        print(f"{Cor.CYAN}Arquivo: {nome_atual}{Cor.RESET}")
        abrir_arquivo(caminho)
        print(f"{Cor.GREY}A analisar estrutura e texto...{Cor.RESET}")
        info = analisar_conteudo_lro(caminho)
        exibir_dados_analise(info)
        
        if pedir_confirmacao(">> Confirmar e assinar? (S/Enter p/ Sim, ESC p/ Não): "):
            dir_arq = os.path.dirname(caminho)
            extensao = os.path.splitext(caminho)[1]
            novo_nome_base = f"{data_str}_{turno}TURNO OK{extensao}"
            novo_caminho = os.path.join(dir_arq, novo_nome_base)
            while True:
                try:
                    os.rename(caminho, novo_caminho)
                    print(f"{Cor.GREEN}   [V] Validado e Padronizado: {novo_nome_base}{Cor.RESET}")
                    break
                except PermissionError: 
                    if not pedir_confirmacao(f"{Cor.RED}   [!] Feche o arquivo PDF! S/Enter para tentar novamente, ESC para pular.{Cor.RESET}"):
                        break
                except Exception as e: print(f"{Cor.RED}Erro: {e}{Cor.RESET}"); break
        else: print(f"{Cor.GREY}   [-] Mantido.{Cor.RESET}")

def exibir_dados_analise(info):
    if not info: return
    print("-" * 60)
    print(f"📅 {Cor.GREEN}{info['cabecalho']}{Cor.RESET}")
    print(f"👤 RESPONSÁVEL: {Cor.CYAN}{info['responsavel']}{Cor.RESET}")
    if info['assinatura']: print(f"🔏 ASSINATURA: {Cor.GREEN}OK (Certificado Digital Detectado) ✅{Cor.RESET}")
    else: print(f"🔏 ASSINATURA: {Cor.RED}NÃO DETECTADA NA ESTRUTURA ❌{Cor.RESET}")
    print("-" * 60)
    
    # Espaçamento corrigido conforme solicitado!
    print(f"⬅️ Recebeu: {info['recebeu']}")
    print(f"➡️ Passou:  {info['passou']}")
    
    print(f"{Cor.WHITE}👥 Equipe:{Cor.RESET}")
    print(f"   • SMC: {info['equipe']['smc']}")
    print(f"   • BCT: {info['equipe']['bct']}")
    print(f"   • OEA: {info['equipe']['oea']}")
    print("-" * 60)


# ==========================================
#         MÓDULO 1: VERIFICADOR LRO
# ==========================================
def modulo_verificador_lro():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    while True:
        limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        print(f"{Cor.CYAN}=== Verificador LRO ==={Cor.RESET}")
        if os.name == 'nt': print(f"{Cor.GREY}Conectado: {Config.CAMINHO_RAIZ}{Cor.RESET}")
        
        inp_mes = input(f"MÊS (Enter para {mes_atual}): ")
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        inp_ano = input(f"ANO (Enter para {ano_atual_curto}): ")
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        
        ano_longo = "20" + ano_curto
        try: ano_errado = str(int(ano_curto) - 1).zfill(2)
        except ValueError: ano_errado = "00"

        path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
        if os.name != 'nt' and not os.path.exists(path_ano): path_ano = Config.CAMINHO_RAIZ 
            
        if os.name == 'nt' and not os.path.exists(path_ano):
            print(f"{Cor.RED}Pasta LRO {ano_longo} não existe.{Cor.RESET}"); time.sleep(2); continue
            
        path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
        if os.name == 'nt' and not os.path.exists(path_mes):
            print(f"{Cor.RED}Pasta do mês não encontrada.{Cor.RESET}"); time.sleep(2); continue
        
        if os.name != 'nt': path_mes = "." 

        try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
        except Exception: continue
            
        if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

        print(f"\n{Cor.YELLOW}>> Analisando {Config.MAPA_PASTAS.get(mes)}...{Cor.RESET}")
        print("-" * 60)

        problemas = 0
        relatorio, lista_pendentes, lista_para_criar, lista_ano_errado = [], [], [], []

        for dia in range(1, qtd_dias + 1):
            dia_fmt = f"{dia:02d}"
            data_str = f"{dia_fmt}{mes}{ano_curto}"
            data_str_err = f"{dia_fmt}{mes}{ano_errado}"
            turnos = calcular_turnos_validos(dia, mes, agora.day, mes_atual, agora.hour)

            for turno in turnos:
                arquivos_turno = buscar_arquivos_flexivel(path_mes, data_str, turno)
                arquivos_errados = buscar_arquivos_flexivel(path_mes, data_str_err, turno)
                tem_ok = [f for f in arquivos_turno if "OK" in f.upper()]
                tem_falta_ass = [f for f in arquivos_turno if "FALTA ASS" in f.upper()]
                tem_falta_lro = [f for f in arquivos_turno if "FALTA LRO" in f.upper()]
                tem_novo = [f for f in arquivos_turno if "OK" not in f.upper() and "FALTA" not in f.upper()]

                if tem_ok: continue
                elif tem_falta_ass:
                    print(f"{Cor.YELLOW}[!] FALTA ASSINATURA: {os.path.basename(tem_falta_ass[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_ass[0])); problemas += 1
                elif tem_novo and tem_falta_lro:
                    print(f"{Cor.DARK_YELLOW}[!] LRO FEITO (TXT PRESENTE): {os.path.basename(tem_novo[0])}{Cor.RESET}")
                    lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno}); problemas += 1
                elif tem_falta_lro:
                    print(f"{Cor.MAGENTA}[!] NÃO CONFECCIONADO: {os.path.basename(tem_falta_lro[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_lro[0])); problemas += 1
                elif tem_novo:
                    n = os.path.basename(tem_novo[0])
                    msg_padrao = "NOME FORA DO PADRÃO" if ("-" in n or " " in n.replace("TURNO", "")) else "PENDENTE DE VERIFICAÇÃO"
                    print(f"{Cor.CYAN}[?] {msg_padrao}: {n}{Cor.RESET}")
                    lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno}); problemas += 1
                elif arquivos_errados:
                    print(f"{Cor.DARK_RED}[!] ANO INCORRETO: {os.path.basename(arquivos_errados[0])}{Cor.RESET}")
                    lista_ano_errado.append(arquivos_errados[0]); problemas += 1
                else:
                    print(f"{Cor.RED}[X] INEXISTENTE: {data_str} - {turno}º Turno{Cor.RESET}")
                    relatorio.append(f"Dia {dia_fmt} - {turno}º Turno: NÃO ENCONTRADO"); problemas += 1
                    lista_para_criar.append({"str": data_str, "turno": turno})

        print("-" * 60)
        if problemas == 0: print(f"{Cor.GREEN}Tudo em dia!{Cor.RESET}\n")

        corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado)

        if problemas > 0 and relatorio:
            print(f"\n{Cor.bg_BLUE}--- COPIAR E COBRAR ---{Cor.RESET}")
            for l in relatorio: print(l)
            print("-" * 30 + "\n")

        if lista_para_criar:
            if pedir_confirmacao(f"{Cor.YELLOW}Criar arquivos FALTA LRO? (S/Enter p/ Sim, ESC p/ Não): {Cor.RESET}"):
                for i in lista_para_criar:
                    n = f"{i['str']}_{i['turno']}TURNO FALTA LRO ().txt"
                    try: 
                        with open(os.path.join(path_mes, n), 'w') as f: f.write("Falta")
                        print(f"{Cor.GREEN}   [V] Criado: {n}{Cor.RESET}")
                    except Exception as e: print(f"{Cor.RED}   Erro ao criar {n}: {e}{Cor.RESET}")

        processo_verificacao_visual(lista_pendentes)

        if not pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outro mês? (S/Enter p/ Sim, ESC p/ Voltar ao menu): {Cor.RESET}"): 
            break


# ==========================================
#         MÓDULO 2: ESCALA CUMPRIDA
# ==========================================
def modulo_escala_cumprida():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    mapa_smc, mapa_bct, mapa_oea = DadosEfetivo.mapear_efetivo()

    while True:
        limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        print(f"{Cor.ORANGE}=== SISTEMA LRO - Escala Cumprida ==={Cor.RESET}")
        
        inp_mes = input(f"MÊS (Enter para {mes_atual}): ")
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        inp_ano = input(f"ANO (Enter para {ano_atual_curto}): ")
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        ano_longo = "20" + ano_curto

        # Loop interno para permitir verificar outras especialidades do mesmo mês
        while True:
            limpar_tela()
            print(f"{Cor.ORANGE}=== SISTEMA LRO - Escala Cumprida ({mes}/{ano_curto}) ==={Cor.RESET}")
            print(f"\n{Cor.CYAN}Qual escala deseja gerar?{Cor.RESET}")
            print("  [1] SMC")
            print("  [2] BCT")
            print("  [3] OEA")
            print("  [0] Voltar ao Menu")
            opcao_escala = input("\nOpção: ")
            
            if opcao_escala == '0':
                return # Sai completamente para o menu principal
                
            if opcao_escala not in ['1', '2', '3']:
                print(f"{Cor.RED}Opção inválida! Tente novamente.{Cor.RESET}")
                time.sleep(1.5)
                continue

            path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
            if os.name != 'nt' and not os.path.exists(path_ano): path_ano = Config.CAMINHO_RAIZ 
                
            path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
            if os.name != 'nt': path_mes = "." 

            if not os.path.exists(path_mes) and os.name == 'nt':
                print(f"{Cor.RED}Pasta do mês não encontrada ({path_mes}).{Cor.RESET}"); time.sleep(2); break

            try: 
                qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
            except Exception: 
                break
                
            if mes == mes_atual and ano_curto == ano_atual_curto: 
                qtd_dias = agora.day

            print(f"\n{Cor.GREY}Extraindo dados de forma invisível... Aguarde.{Cor.RESET}\n")

            if opcao_escala == '1':
                print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM |  SMC  {Cor.RESET}")
                tracos_separador = 19
            elif opcao_escala in ['2', '3']:
                print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM | 1º TURNO | 2º TURNO | 3º TURNO {Cor.RESET}")
                tracos_separador = 45

            for dia in range(1, qtd_dias + 1):
                dia_fmt = f"{dia:02d}"
                data_str = f"{dia_fmt}{mes}{ano_curto}"
                
                data_dt = datetime.date(int(ano_longo), int(mes), dia)
                sigla_sem = Config.MAPA_SEMANA[data_dt.weekday()]
                
                turnos = calcular_turnos_validos(dia, mes, agora.day, mes_atual, agora.hour)

                dia_dados = {
                    'smc': '---',
                    'bct': {1: '---', 2: '---', 3: '---'},
                    'oea': {1: '---', 2: '---', 3: '---'}
                }

                for turno in turnos:
                    arquivos = buscar_arquivos_flexivel(path_mes, data_str, turno)
                    if not arquivos: continue

                    arquivos_ok = [f for f in arquivos if "OK" in f.upper()]
                    arquivo_alvo = arquivos_ok[0] if arquivos_ok else arquivos[0]
                    
                    if "FALTA LRO" in arquivo_alvo.upper() and arquivo_alvo.endswith('.txt'):
                        dia_dados['bct'][turno] = 'PND'
                        dia_dados['oea'][turno] = 'PND'
                        continue

                    info = analisar_conteudo_lro(arquivo_alvo)
                    if info:
                        leg_smc = encontrar_legenda(info['equipe']['smc'], mapa_smc)
                        if dia_dados['smc'] == '---' and leg_smc not in ['---', '???']:
                            dia_dados['smc'] = leg_smc
                        elif dia_dados['smc'] == '---':
                            dia_dados['smc'] = leg_smc
                            
                        dia_dados['bct'][turno] = encontrar_legenda(info['equipe']['bct'], mapa_bct)
                        dia_dados['oea'][turno] = encontrar_legenda(info['equipe']['oea'], mapa_oea)
                    else:
                        dia_dados['bct'][turno] = 'ERR'
                        dia_dados['oea'][turno] = 'ERR'

                if opcao_escala == '1':
                    smc = dia_dados['smc']
                    print(f" {dia_fmt}  | {sigla_sem} |  {smc:^3}  ")
                elif opcao_escala == '2':
                    b1, b2, b3 = dia_dados['bct'][1], dia_dados['bct'][2], dia_dados['bct'][3]
                    print(f" {dia_fmt}  | {sigla_sem} |   {b1:^4}   |   {b2:^4}   |   {b3:^4}   ")
                elif opcao_escala == '3':
                    o1, o2, o3 = dia_dados['oea'][1], dia_dados['oea'][2], dia_dados['oea'][3]
                    print(f" {dia_fmt}  | {sigla_sem} |   {o1:^4}   |   {o2:^4}   |   {o3:^4}   ")

            print("-" * tracos_separador)
            
            # Pergunta instantânea: S / Enter continua neste mês, ESC volta ao menu.
            if not pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outra especialidade deste mês? (S/Enter p/ Sim, ESC p/ Voltar ao menu): {Cor.RESET}"):
                return # Volta ao menu principal


# ==========================================
#              MENU PRINCIPAL
# ==========================================
def menu_principal():
    while True:
        limpar_tela()
        print(f"{Cor.ORANGE}")
        print("╔══════════════════════════════════════╗")
        print("║                                      ║")
        print("║             SISTEMA LRO              ║")
        print("║                                      ║")
        print("╚══════════════════════════════════════╝")
        print(f"{Cor.RESET}")
        
        print("Escolha uma funcionalidade:\n")
        print(f"  {Cor.CYAN}[1]{Cor.RESET} Verificador LRO")
        print(f"  {Cor.CYAN}[2]{Cor.RESET} Escala Cumprida")
        print(f"  {Cor.RED}[0]{Cor.RESET} Sair\n")
        
        opcao = input("Opção: ")
        
        if opcao == '1':
            modulo_verificador_lro()
        elif opcao == '2':
            modulo_escala_cumprida()
        elif opcao == '0':
            limpar_tela()
            print(f"{Cor.GREEN}A sair do SISTEMA LRO... Até à próxima!{Cor.RESET}\n")
            break
        else:
            print(f"\n{Cor.RED}[!] Opção inválida! Tente novamente.{Cor.RESET}")
            time.sleep(1.5)

if __name__ == "__main__":
    menu_principal()