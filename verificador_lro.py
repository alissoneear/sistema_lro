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


# --- CONFIGURAÇÃO ---
class Config:
    # Ajustado para manter a compatibilidade com o ambiente Windows mapeado em rede
    CAMINHO_RAIZ = r"E:\dev\sistema_lro\ARQUIVOS" if os.name == 'nt' else "."
    
    MAPA_PASTAS = {
        "01": "1 - JANEIRO",   "02": "2 - FEVEREIRO", "03": "3 - MARÇO", "04": "4 - ABRIL",
        "05": "5 - MAIO",      "06": "6 - JUNHO",     "07": "7 - JULHO", "08": "8 - AGOSTO",
        "09": "9 - SETEMBRO",  "10": "10 - OUTUBRO",  "11": "11 - NOVEMBRO", "12": "12 - DEZEMBRO"
    }

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

# --- UTILITÁRIOS ---
def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def abrir_arquivo(caminho):
    try:
        if os.name == 'nt':
            os.startfile(caminho)
        elif sys.platform == 'darwin': 
            os.system(f'open "{caminho}"')
        else:
            os.system(f'xdg-open "{caminho}"')
    except Exception as e:
        print(f"{Cor.RED}[Erro ao abrir]: {e}{Cor.RESET}")

# --- PROCESSAMENTO DE PDF ---
def verificar_assinatura_estrutural(caminho_pdf):
    if not PYPDF_ENABLED: 
        return False
    try:
        reader = PdfReader(caminho_pdf)
        if reader.get_fields():
            for field in reader.get_fields().values():
                if field.field_type == "/Sig":
                    return True
        
        for page in reader.pages:
            if "/Annots" in page:
                for annot in page["/Annots"]:
                    obj = annot.get_object()
                    if obj.get("/FT") == "/Sig":
                        return True
    except Exception as e:
        # Silenciar ou logar o erro se o PDF estiver corrompido
        pass
    return False

def analisar_conteudo_lro(caminho_pdf):
    dados = {
        "cabecalho": "---", "recebeu": "---", "passou": "---", 
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
        except Exception:
            pass

    return dados

def extrair_dados_texto(texto_linear, dados):
    """Extrai informações via Regex do texto plano do PDF."""
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

# --- LÓGICA DE FICHEIROS ---
def buscar_arquivos_flexivel(pasta, data_string, turno):
    padroes = [
        f"{data_string}_{turno}TURNO*", f"{data_string}-{turno}TURNO*",
        f"{data_string}_{turno} TURNO*", f"{data_string}-{turno} TURNO*",
        f"{data_string} {turno}TURNO*"
    ]
    encontrados = []
    for p in padroes:
        encontrados.extend(glob.glob(os.path.join(pasta, p)))
    return list(set(encontrados))

def calcular_turnos_validos(dia, mes_str, dia_atual, mes_atual, hora_atual):
    turnos = [1, 2, 3]
    if dia == dia_atual and mes_str == mes_atual:
        if hora_atual < 14: turnos = []
        elif hora_atual < 21: turnos = [1]
        else: turnos = [1, 2]
    return turnos

# --- FLUXO PRINCIPAL E INTERFACE ---
def corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado):
    if not lista_ano_errado: return
    
    resposta = input(f"{Cor.DARK_RED}Corrigir {len(lista_ano_errado)} anos errados? (S/N): {Cor.RESET}").lower()
    if resposta in ['s', 'ok']:
        for arq in lista_ano_errado:
            abrir_arquivo(arq)
            if input(f"   Renomear {os.path.basename(arq)}? (S/OK): ").lower() in ['s', 'ok']:
                try: 
                    os.rename(arq, arq.replace(ano_errado, ano_curto))
                    print(f"{Cor.GREEN}   Renomeado com sucesso.{Cor.RESET}")
                except Exception as e: 
                    print(f"{Cor.RED}   Erro ao renomear: {e}{Cor.RESET}")

def processo_verificacao_visual(lista_pendentes):
    if not lista_pendentes: return
    
    print(f"{Cor.CYAN}Verificar {len(lista_pendentes)} livros pendentes?{Cor.RESET}")
    if input("Iniciar UM POR UM? (S/N): ").lower() not in ['s', 'ok']: return

    for cnt, item in enumerate(lista_pendentes, start=1):
        caminho, data_str, turno = item['path'], item['data'], item['turno']
        nome_atual = os.path.basename(caminho)
        
        print(f"\n{Cor.YELLOW}--- ({cnt}/{len(lista_pendentes)}) ---{Cor.RESET}")
        print(f"{Cor.CYAN}Arquivo: {nome_atual}{Cor.RESET}")
        
        abrir_arquivo(caminho)
        print(f"{Cor.GREY}A analisar estrutura e texto...{Cor.RESET}")
        info = analisar_conteudo_lro(caminho)
        
        exibir_dados_analise(info)
        
        if input(">> Confirmar e assinar? (S/OK): ").lower() in ['s', 'ok']:
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
                    input(f"{Cor.RED}   [!] Feche o ficheiro PDF! Pressione Enter para tentar novamente.{Cor.RESET}")
                except Exception as e:
                    print(f"{Cor.RED}Erro: {e}{Cor.RESET}")
                    break
        else:
            print(f"{Cor.GREY}   [-] Mantido.{Cor.RESET}")

def exibir_dados_analise(info):
    if not info: return
    print("-" * 60)
    print(f"📅 {Cor.GREEN}{info['cabecalho']}{Cor.RESET}")
    if info['assinatura']:
        print(f"🔏 ASSINATURA: {Cor.GREEN}OK (Certificado Digital Detectado) ✅{Cor.RESET}")
    else:
        print(f"🔏 ASSINATURA: {Cor.RED}NÃO DETETADA NA ESTRUTURA ❌{Cor.RESET}")
    print("-" * 60)
    print(f"⬅️  Recebeu: {info['recebeu']}")
    print(f"➡️  Passou:  {info['passou']}")
    print(f"{Cor.WHITE}👥 Equipa:{Cor.RESET}")
    print(f"   • SMC: {info['equipe']['smc']}")
    print(f"   • BCT: {info['equipe']['bct']}")
    print(f"   • OEA: {info['equipe']['oea']}")
    print("-" * 60)

def main():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    while True:
        limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        print(f"{Cor.CYAN}=== Verificador LRO 19.0 (Refatorado) ==={Cor.RESET}")
        if os.name == 'nt': print(f"{Cor.GREY}Conectado: {Config.CAMINHO_RAIZ}{Cor.RESET}")
        
        inp_mes = input(f"MÊS (Enter para {mes_atual}): ")
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        inp_ano = input(f"ANO (Enter para {ano_atual_curto}): ")
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        
        ano_longo = "20" + ano_curto
        try: 
            ano_errado = str(int(ano_curto) - 1).zfill(2)
        except ValueError: 
            ano_errado = "00"

        # Resolução de Caminhos
        path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
        if os.name != 'nt' and not os.path.exists(path_ano): 
            path_ano = Config.CAMINHO_RAIZ 
            
        if os.name == 'nt' and not os.path.exists(path_ano):
            print(f"{Cor.RED}Pasta LRO {ano_longo} não existe.{Cor.RESET}")
            time.sleep(2); continue
            
        path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
        if os.name == 'nt' and not os.path.exists(path_mes):
            print(f"{Cor.RED}Pasta do mês não encontrada.{Cor.RESET}")
            time.sleep(2); continue
        
        if os.name != 'nt': path_mes = "." 

        try: 
            qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
        except Exception: 
            continue
            
        if mes == mes_atual and ano_curto == ano_atual_curto: 
            qtd_dias = agora.day

        print(f"\n{Cor.YELLOW}>> A analisar {Config.MAPA_PASTAS.get(mes)}...{Cor.RESET}")
        print("-" * 60)

        # Variáveis de Estado
        problemas = 0
        relatorio = []
        lista_pendentes = [] 
        lista_para_criar = []
        lista_ano_errado = []

        # Lógica principal de varredura
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

                if tem_ok: 
                    continue
                elif tem_falta_ass:
                    print(f"{Cor.YELLOW}[!] FALTA ASSINATURA: {os.path.basename(tem_falta_ass[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_ass[0])); problemas += 1
                elif tem_novo and tem_falta_lro:
                    print(f"{Cor.DARK_YELLOW}[!] LRO FEITO (TXT PRESENTE): {os.path.basename(tem_novo[0])}{Cor.RESET}")
                    lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno})
                    problemas += 1
                elif tem_falta_lro:
                    print(f"{Cor.MAGENTA}[!] NÃO CONFECIONADO: {os.path.basename(tem_falta_lro[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_lro[0])); problemas += 1
                elif tem_novo:
                    n = os.path.basename(tem_novo[0])
                    msg_padrao = "NOME FORA DO PADRÃO" if ("-" in n or " " in n.replace("TURNO", "")) else "PENDENTE DE VERIFICAÇÃO"
                    print(f"{Cor.CYAN}[?] {msg_padrao}: {n}{Cor.RESET}")
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
        if problemas == 0: 
            print(f"{Cor.GREEN}Tudo em dia!{Cor.RESET}\n")

        # --- Ações Corretivas ---
        corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado)

        if problemas > 0 and relatorio:
            print(f"\n{Cor.bg_BLUE}--- COPIAR E COBRAR ---{Cor.RESET}")
            for l in relatorio: print(l)
            print("-" * 30 + "\n")

        if lista_para_criar:
            if input(f"{Cor.YELLOW}Criar ficheiros FALTA LRO? (S/N): {Cor.RESET}").lower() in ['s', 'ok']:
                for i in lista_para_criar:
                    n = f"{i['str']}_{i['turno']}TURNO FALTA LRO ().txt"
                    try: 
                        with open(os.path.join(path_mes, n), 'w') as f: f.write("Falta")
                        print(f"{Cor.GREEN}   [V] Criado: {n}{Cor.RESET}")
                    except Exception as e: 
                        print(f"{Cor.RED}   Erro ao criar {n}: {e}{Cor.RESET}")

        processo_verificacao_visual(lista_pendentes)

        if input("\nSair? (S/Enter): ").lower() in ['s', 'ok']: 
            break

if __name__ == "__main__":
    main()