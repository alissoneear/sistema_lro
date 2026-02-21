import os
import time
import datetime
import calendar
import re

# Adicionado o DadosEfetivo aqui nas importações!
from config import Config, Cor, DadosEfetivo 
import utils

def corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado):
    if not lista_ano_errado: return
    if utils.pedir_confirmacao(f"{Cor.DARK_RED}Corrigir {len(lista_ano_errado)} anos errados? (S/Enter p/ Sim, ESC p/ Pular): {Cor.RESET}"):
        for arq in lista_ano_errado:
            utils.abrir_arquivo(arq)
            if utils.pedir_confirmacao(f"   Renomear {os.path.basename(arq)}? (S/Enter p/ Sim, ESC p/ Não): "):
                try: 
                    os.rename(arq, arq.replace(ano_errado, ano_curto))
                    print(f"{Cor.GREEN}   Renomeado com sucesso.{Cor.RESET}")
                except Exception as e: print(f"{Cor.RED}   Erro ao renomear: {e}{Cor.RESET}")

def exibir_dados_analise(info):
    if not info: return
    print("-" * 60)
    print(f"📅 {Cor.GREEN}{info['cabecalho']}{Cor.RESET}")
    print(f"👤 RESPONSÁVEL: {Cor.CYAN}{info['responsavel']}{Cor.RESET}")
    if info['assinatura']: print(f"🔏 ASSINATURA: {Cor.GREEN}OK (Certificado Digital Detectado) ✅{Cor.RESET}")
    else: print(f"🔏 ASSINATURA: {Cor.RED}NÃO DETECTADA NA ESTRUTURA ❌{Cor.RESET}")
    print("-" * 60)
    print(f"⬅️ Recebeu: {info['recebeu']}")
    print(f"➡️ Passou:  {info['passou']}")
    
    print(f"{Cor.WHITE}👥 Equipe:{Cor.RESET}") 
    print(f"   • SMC: {info['equipe']['smc']}")
    print(f"   • BCT: {info['equipe']['bct']}")
    print(f"   • OEA: {info['equipe']['oea']}")
    print("-" * 60)

def renomear_arquivo(caminho_atual, novo_caminho):
    while True:
        try:
            os.rename(caminho_atual, novo_caminho)
            print(f"{Cor.GREEN}   [V] Validado e Padronizado: {os.path.basename(novo_caminho)}{Cor.RESET}")
            break
        except PermissionError: 
            if not utils.pedir_confirmacao(f"{Cor.RED}   [!] Feche o ficheiro PDF! S/Enter p/ tentar novamente, ESC p/ pular.{Cor.RESET}"):
                break
        except Exception as e: 
            print(f"{Cor.RED}Erro: {e}{Cor.RESET}")
            break

def processo_verificacao_visual(lista_pendentes, mes, ano_curto):
    if not lista_pendentes: return
    print(f"{Cor.ORANGE}Verificar {len(lista_pendentes)} livros pendentes?{Cor.RESET}")
    if not utils.pedir_confirmacao("Iniciar UM POR UM? (S/Enter p/ Sim, ESC p/ Pular): "): return

    print("\n")
    abrir_pdfs = utils.pedir_confirmacao(f"{Cor.YELLOW}Deseja ABRIR os PDFs para acompanhamento? (S/Enter p/ Sim, ESC p/ Modo Rápido): {Cor.RESET}")

    for cnt, item in enumerate(lista_pendentes, start=1):
        caminho, data_str, turno = item['path'], item['data'], item['turno']
        nome_atual = os.path.basename(caminho)
        
        # --- NOVO ESPAÇAMENTO E DESTAQUE VISUAL ---
        print("\n\n") # Duas quebras de linha para criar uma área de "respiro"
        print(f"{Cor.bg_ORANGE}{Cor.WHITE}  ▶ LIVRO ({cnt}/{len(lista_pendentes)}) - Arquivo: {nome_atual}  {Cor.RESET}")
        
        if abrir_pdfs:
            utils.abrir_arquivo(caminho)
            
        print(f"{Cor.GREY}A analisar estrutura e texto...{Cor.RESET}")
        info = utils.analisar_conteudo_lro(caminho, mes, ano_curto)
        exibir_dados_analise(info)
        
        dir_arq = os.path.dirname(caminho)
        extensao = os.path.splitext(caminho)[1]
        
        if info['assinatura']:
            if utils.pedir_confirmacao(">> Confirmar e assinar OK? (S/Enter p/ Sim, ESC p/ Não): "):
                novo_nome_base = f"{data_str}_{turno}TURNO OK{extensao}"
                novo_caminho = os.path.join(dir_arq, novo_nome_base)
                renomear_arquivo(caminho, novo_caminho)
            else:
                print(f"{Cor.GREY}   [-] Mantido.{Cor.RESET}")
        else:
            nome_dic = extrair_nome_relato(info.get('responsavel', ''), mes, ano_curto)
            
            if nome_dic != "???":
                sugestao_str = f" ({nome_dic})"
            else:
                resp_base = utils.extrair_nome_base(info.get('responsavel', ''))
                sugestao_str = f" ({resp_base})" if resp_base not in ["---", "???", ""] else ""
            
            msg = f">> {Cor.RED}ASSINATURA AUSENTE!{Cor.RESET} Renomear p/ FALTA ASS{sugestao_str}? (S/Enter p/ Sim, ESC p/ Não): "
            if utils.pedir_confirmacao(msg):
                novo_nome_base = f"{data_str}_{turno}TURNO FALTA ASS{sugestao_str}{extensao}"
                novo_caminho = os.path.join(dir_arq, novo_nome_base)
                renomear_arquivo(caminho, novo_caminho)
            else:
                print(f"{Cor.GREY}   [-] Mantido.{Cor.RESET}")

# --- FUNÇÃO INTELIGENTE DE EXTRAÇÃO DE NOME ---
def extrair_nome_relato(raw_str, mes=None, ano_curto=None):
    """Procura o nome oficial do militar dentro do texto sujo do PDF usando o Dicionário."""
    if not raw_str or raw_str in ["---", "???"]: return "???"
    
    # 1. Normaliza o texto (tira acentos, converte 'IS' para '1S', etc.)
    texto_norm = utils.normalizar_texto(raw_str)
    
    # 2. Carrega todos os nomes de guerra conhecidos
    mapa_smc, mapa_bct, mapa_oea = DadosEfetivo.mapear_efetivo(mes, ano_curto)
    todos_mapas = {**mapa_smc, **mapa_bct, **mapa_oea}
    
    # 3. Extrai apenas as "essências" (GIOVANNI, SAULO, RUI, etc.)
    nomes_base = [utils.extrair_nome_base(ng) for ng in todos_mapas.keys()]
    
    # 4. Ordena por tamanho (os nomes maiores primeiro) para evitar falsos positivos
    nomes_base.sort(key=len, reverse=True)
    
    # 5. Varre a frase. Usamos \b (Word Boundary) para não confundir "RUI" dentro de "ARUINADO"
    for nome_base in nomes_base:
        if re.search(rf'\b{re.escape(nome_base)}\b', texto_norm):
            return nome_base
            
    return "???"

def executar():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    while True:
        utils.limpar_tela()
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
            
        if not os.path.exists(path_ano):
            print(f"{Cor.RED}Pasta LRO {ano_longo} não existe.{Cor.RESET}")
            time.sleep(2)
            continue
            
        path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
        if not os.path.exists(path_mes):
            print(f"{Cor.RED}Pasta do mês não encontrada.{Cor.RESET}")
            time.sleep(2)
            continue 

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
            turnos = utils.calcular_turnos_validos(dia, mes, agora.day, mes_atual, agora.hour)

            for turno in turnos:
                arquivos_turno = utils.buscar_arquivos_flexivel(path_mes, data_str, turno)
                arquivos_errados = utils.buscar_arquivos_flexivel(path_mes, data_str_err, turno)
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
                    
                    lista_para_criar.append({
                        "str": data_str, "turno": turno, 
                        "dia": dia, "mes": mes, "ano": ano_longo
                    })

        print("-" * 60)
        if problemas == 0: print(f"{Cor.GREEN}Tudo em dia!{Cor.RESET}\n")

        corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado)

        if problemas > 0 and relatorio:
            print(f"\n{Cor.bg_BLUE}{Cor.WHITE} --- COPIAR E COBRAR --- {Cor.RESET}")
            for l in relatorio: print(l)
            print("-" * 30 + "\n")
            
            # === NOVO: EXPORTAÇÃO AUTOMÁTICA DO RELATÓRIO ===
            if utils.pedir_confirmacao(f">> Gerar ficheiro TXT com este relatório? (S/Enter p/ Sim, ESC p/ Não): "):
                nome_relatorio = f"Relatorio_Pendencias_{mes}_{ano_longo}.txt"
                caminho_relatorio = os.path.join(path_mes, nome_relatorio)
                try:
                    with open(caminho_relatorio, 'w', encoding='utf-8') as f:
                        f.write(f"--- RELATÓRIO DE PENDÊNCIAS LRO ({mes}/{ano_longo}) ---\n\n")
                        for l in relatorio: f.write(l + "\n")
                    print(f"{Cor.GREEN}✅ Relatório salvo na pasta do mês: {nome_relatorio}{Cor.RESET}")
                except Exception as e:
                    print(f"{Cor.RED}Erro ao salvar o ficheiro: {e}{Cor.RESET}")

        if lista_para_criar:
            print(f"\n{Cor.CYAN}--- GERAÇÃO INTELIGENTE DE ARQUIVOS 'FALTA LRO' ---{Cor.RESET}")
            for item in lista_para_criar:
                data_str = item['str']
                turno = item['turno']
                dt_atual = datetime.date(int(item['ano']), int(item['mes']), item['dia'])
                
                # Calcular Turno Anterior e Seguinte
                if turno == 1:
                    dt_prev, tr_prev = dt_atual - datetime.timedelta(days=1), 3
                    dt_next, tr_next = dt_atual, 2
                elif turno == 2:
                    dt_prev, tr_prev = dt_atual, 1
                    dt_next, tr_next = dt_atual, 3
                else:
                    dt_prev, tr_prev = dt_atual, 2
                    dt_next, tr_next = dt_atual + datetime.timedelta(days=1), 1
                
                def get_path_mes(dt):
                    path_a = os.path.join(Config.CAMINHO_RAIZ, f"LRO {dt.strftime('%Y')}")
                    return os.path.join(path_a, Config.MAPA_PASTAS.get(dt.strftime('%m'), "X"))

                # Alterado para usar a nova função de extração pelo dicionário
                def get_nome_adjacente(dt, tr, is_prev):
                    p_mes = get_path_mes(dt)
                    d_str = f"{dt.day:02d}{dt.strftime('%m')}{dt.strftime('%y')}"
                    arquivos = utils.buscar_arquivos_flexivel(p_mes, d_str, tr)
                    
                    # Filtra apenas os PDFs (para não tentar ler arquivos .txt de 'FALTA LRO')
                    pdfs = [f for f in arquivos if f.lower().endswith('.pdf')]
                    
                    if pdfs:
                        # Dá prioridade aos que já têm OK, mas se não tiver, lê o PDF pendente mesmo assim!
                        ok_files = [f for f in pdfs if "OK" in f.upper()]
                        arquivo_alvo = ok_files[0] if ok_files else pdfs[0]
                        
                        mes_dt = dt.strftime('%m')
                        ano_dt = dt.strftime('%y')
                        info = utils.analisar_conteudo_lro(arquivo_alvo, mes_dt, ano_dt)
                        raw = info.get('passou', '') if is_prev else info.get('recebeu', '')
                        return extrair_nome_relato(raw, mes_dt, ano_dt)
                    return "???"

                passou_nome = get_nome_adjacente(dt_prev, tr_prev, True)  
                recebeu_nome = get_nome_adjacente(dt_next, tr_next, False) 
                
                if passou_nome == "???" and recebeu_nome == "???": sugestao = ""
                elif passou_nome == recebeu_nome: sugestao = passou_nome
                elif passou_nome != "???" and recebeu_nome == "???": sugestao = passou_nome
                elif passou_nome == "???" and recebeu_nome != "???": sugestao = recebeu_nome
                else: sugestao = f"{passou_nome} ou {recebeu_nome}"
                    
                sugestao_str = f" ({sugestao})" if sugestao else " ()"
                novo_nome = f"{data_str}_{turno}TURNO FALTA LRO{sugestao_str}.txt"
                
                print(f"\n{Cor.YELLOW}[FALTA LRO] {data_str} - {turno}º Turno{Cor.RESET}")
                print(f" ⬅️  Turno anterior passou para: {Cor.CYAN}{passou_nome}{Cor.RESET}")
                print(f" ➡️  Turno seguinte recebeu de:  {Cor.CYAN}{recebeu_nome}{Cor.RESET}")
                
                if utils.pedir_confirmacao(f">> Criar arquivo '{novo_nome}'? (S/Enter p/ Sim, ESC p/ Pular): "):
                    try: 
                        with open(os.path.join(path_mes, novo_nome), 'w') as f: f.write("Falta")
                        print(f"{Cor.GREEN}   [V] Criado com sucesso.{Cor.RESET}")
                    except Exception as e: 
                        print(f"{Cor.RED}   Erro ao criar: {e}{Cor.RESET}")

        processo_verificacao_visual(lista_pendentes, mes, ano_curto)

        if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outro mês? (S/Enter p/ Sim, ESC p/ Voltar ao menu): {Cor.RESET}"): 
            break