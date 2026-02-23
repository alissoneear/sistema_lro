import os
import time
import datetime
import calendar
import re

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
    if info.get("inconsistencia_data"):
        print(f"{Cor.bg_RED}{Cor.WHITE}⚠️ ALERTA DE COPIAR/COLAR: {info['inconsistencia_data']} {Cor.RESET}")
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
    from rich.console import Console
    from rich.panel import Panel
    from rich.align import Align
    console = Console()
    
    if not lista_pendentes: return
    
    console.print("\n")
    painel_verificacao = Panel(
        f"[bold white]Existem [dark_orange]{len(lista_pendentes)}[/dark_orange] livros pendentes de análise visual e assinatura.[/bold white]",
        title="[bold dark_orange]🔎 INICIAR AUDITORIA VISUAL[/bold dark_orange]",
        border_style="dark_orange",
        padding=(1, 2),
        width=75
    )
    console.print(Align.center(painel_verificacao))
    console.print()
    
    if not utils.pedir_confirmacao(">> Iniciar verificação UM POR UM? (S/Enter p/ Sim, ESC p/ Pular): "): return

    print("\n")
    abrir_pdfs = utils.pedir_confirmacao(f"{Cor.YELLOW}Deseja ABRIR os PDFs para acompanhamento? (S/Enter p/ Sim, ESC p/ Modo Rápido): {Cor.RESET}")
    
    # PERGUNTA PARA INTEGRAÇÃO COM A ESCALA CUMPRIDA
    integrar_escala = utils.pedir_confirmacao(f"{Cor.YELLOW}Deseja alimentar a ESCALA CUMPRIDA durante a verificação? (S/Enter p/ Sim, ESC p/ Não): {Cor.RESET}")

    for cnt, item in enumerate(lista_pendentes, start=1):
        caminho, data_str, turno = item['path'], item['data'], item['turno']
        nome_atual = os.path.basename(caminho)
        
        print("\n\n") 
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
                
        # ALIMENTAÇÃO DO CACHE DA ESCALA CUMPRIDA
        if integrar_escala:
            import json
            m_smc, m_bct, m_oea = DadosEfetivo.mapear_efetivo(mes, ano_curto)
            
            n_smc = info['equipe'].get('smc', '---')
            n_bct = info['equipe'].get('bct', '---')
            n_oea = info['equipe'].get('oea', '---')
            
            l_smc = utils.encontrar_legenda(n_smc, m_smc) if n_smc not in ['---', '???'] else '---'
            l_bct = utils.encontrar_legenda(n_bct, m_bct) if n_bct not in ['---', '???'] else '---'
            l_oea = utils.encontrar_legenda(n_oea, m_oea) if n_oea not in ['---', '???'] else '---'
            
            def obter_nome_pela_legenda(legenda_alvo, mapa):
                for nome_guerra, dados_mapa in mapa.items():
                    if dados_mapa['legenda'] == legenda_alvo:
                        return utils.extrair_nome_base(nome_guerra)
                return "???"
            
            while True:
                print(f"\n{Cor.bg_BLUE}{Cor.WHITE} Aceitar inserção dos dados na escala cumprida? {Cor.RESET}")
                print(f"DIA {data_str[:2]} de {Config.MAPA_PASTAS.get(mes, '')} de 20{ano_curto} | {turno}º Turno")
                print(f" • SMC: {n_smc} ({Cor.CYAN}{l_smc}{Cor.RESET})")
                print(f" • BCT: {n_bct} ({Cor.CYAN}{l_bct}{Cor.RESET})")
                print(f" • OEA: {n_oea} ({Cor.CYAN}{l_oea}{Cor.RESET})")
                
                print(f"\n>> Opção [S/Enter=Sim | M=Modificar | ESC=Não]: ", end='', flush=True)
                
                if os.name == 'nt' and utils.MSVCRT_ENABLED:
                    import msvcrt
                    while True:
                        tecla = msvcrt.getch()
                        if tecla in [b's', b'S', b'\r']:
                            print("Sim")
                            acao = 'S'
                            break
                        elif tecla in [b'm', b'M']:
                            print("Modificar")
                            acao = 'M'
                            break
                        elif tecla == b'\x1b':
                            print("Não")
                            acao = 'N'
                            break
                else:
                    resp = input().strip().upper()
                    acao = 'M' if resp == 'M' else 'N' if resp == 'ESC' else 'S'
                
                if acao == 'N':
                    print(f"{Cor.GREY}   [-] Inserção ignorada.{Cor.RESET}")
                    break
                elif acao == 'M':
                    print(f"\n{Cor.YELLOW}--- MODIFICAR LEGENDAS ---{Cor.RESET}\n")
                    
                    # Função rápida para desenhar o mapa em colunas perfeitamente alinhadas
                    def exibir_mapa_colunas(nome_escala, mapa, cor_titulo):
                        print(f"{cor_titulo}■ EQUIPE {nome_escala}:{Cor.RESET}")
                        
                        # 👉 NOVO PADRÃO: [A] 1S TERBECK
                        itens = [f"[{v['legenda']}] {k.split('-')[0].strip()}" for k, v in mapa.items()]
                        
                        # Quebra a lista e imprime de 4 em 4 colunas (espaçamento ajustado para 22)
                        for i in range(0, len(itens), 4):
                            linha = itens[i:i+4]
                            print("   " + "".join(item.ljust(22) for item in linha))
                        print("") # Linha em branco para respiro
                        
                    # Desenha as 3 escalas na tela
                    exibir_mapa_colunas("SMC", m_smc, Cor.CYAN)
                    exibir_mapa_colunas("BCT", m_bct, Cor.GREEN)
                    exibir_mapa_colunas("OEA", m_oea, Cor.ORANGE)

                    print(f"{Cor.GREY}(Deixe em branco e aperte Enter para manter a legenda atual){Cor.RESET}")
                    
                    nl_smc = input(f"Nova legenda SMC [{l_smc}]: ").strip().upper()
                    nl_bct = input(f"Nova legenda BCT [{l_bct}]: ").strip().upper()
                    nl_oea = input(f"Nova legenda OEA [{l_oea}]: ").strip().upper()
                    
                    # Atualizando a legenda E o nome 
                    if nl_smc: 
                        l_smc = nl_smc
                        n_smc = obter_nome_pela_legenda(l_smc, m_smc)
                    if nl_bct: 
                        l_bct = nl_bct
                        n_bct = obter_nome_pela_legenda(l_bct, m_bct)
                    if nl_oea: 
                        l_oea = nl_oea
                        n_oea = obter_nome_pela_legenda(l_oea, m_oea)
                    continue
                    
                elif acao == 'S':
                    dia_str_int = str(int(data_str[:2])) 
                    turno_str = str(turno)
                    responsavel_ass = info.get('responsavel', '???')
                    
                    def salvar_cache(esp, dia_s, turno_s, leg, opcao_escala):
                        if leg in ['---', '???', '']: return
                        cam = os.path.join(dir_arq, f".cache_escala_{esp}.json")
                        cache = {}
                        if os.path.exists(cam):
                            try:
                                with open(cam, 'r', encoding='utf-8') as f: cache = json.load(f)
                            except: pass
                        
                        if dia_s not in cache: cache[dia_s] = {}
                        
                        if opcao_escala == '1': 
                            cache[dia_s]['smc'] = leg
                        else: 
                            # 👉 CORREÇÃO 2: Salvando o Dicionário completo para o Turbo Cache não quebrar!
                            cache[dia_s][turno_s] = {
                                'legenda': leg,
                                'assinatura_nome': responsavel_ass
                            }
                            
                        try:
                            with open(cam, 'w', encoding='utf-8') as f: json.dump(cache, f, indent=4)
                        except: pass
                        
                    salvar_cache('smc', dia_str_int, turno_str, l_smc, '1')
                    salvar_cache('bct', dia_str_int, turno_str, l_bct, '2')
                    salvar_cache('oea', dia_str_int, turno_str, l_oea, '3')
                    
                    print(f"{Cor.GREEN}   ✅ Cache da Escala Cumprida alimentado com sucesso!{Cor.RESET}")
                    break

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
    import os
    import time
    import datetime
    import calendar
    from rich.console import Console
    from rich.panel import Panel
    from rich.align import Align
    from rich.text import Text
    from rich.prompt import Prompt
    
    console = Console()

    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        console.print(Align.center(Panel(f"[bold red][ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.[/bold red]", border_style="red")))
        input()
        return

    while True:
        utils.limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        # 1. PAINEL DE TÍTULO DO MÓDULO
        titulo = Text("SISTEMA LRO\nVerificador e Assinador de Documentos", justify="center", style="bold dark_orange")
        painel_titulo = Panel(titulo, border_style="dark_orange", padding=(1, 2), width=65)
        console.print(Align.center(painel_titulo))
        console.print("\n")

        # 2. INPUT DE MÊS E ANO (Modernizado)
        if os.name == 'nt': 
            console.print(Align.center(f"[dim grey]Conectado: {Config.CAMINHO_RAIZ}[/dim grey]"))
            console.print()
            
        inp_mes = console.input(" " * 18 + f"[bold dark_orange]MÊS[/bold dark_orange] [dim white](Enter para {mes_atual}):[/dim white] ").strip()
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        
        inp_ano = console.input(" " * 18 + f"[bold dark_orange]ANO[/bold dark_orange] [dim white](Enter para {ano_atual_curto}):[/dim white] ").strip()
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        
        ano_longo = "20" + ano_curto
        try: ano_errado = str(int(ano_curto) - 1).zfill(2)
        except ValueError: ano_errado = "00"

        path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
            
        if not os.path.exists(path_ano):
            console.print(Align.center(Panel(f"[bold red]Pasta LRO {ano_longo} não existe.[/bold red]", border_style="red")))
            time.sleep(2)
            continue
            
        path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
        if not os.path.exists(path_mes):
            console.print(Align.center(Panel(f"[bold red]Pasta do mês {mes} não encontrada em {path_ano}.[/bold red]", border_style="red")))
            time.sleep(2)
            continue 

        try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
        except Exception: continue
            
        if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

        # 3. ANÁLISE SILENCIOSA E PAINEL DE STATUS
        utils.limpar_tela()
        titulo_analise = Text(f"ANALISANDO DIRETÓRIOS: {Config.MAPA_PASTAS.get(mes)} / {ano_longo}", justify="center", style="bold white on deep_sky_blue1")
        console.print(Align.center(Panel(titulo_analise, border_style="deep_sky_blue1", width=65)))
        
        problemas = 0
        relatorio, lista_pendentes, lista_para_criar, lista_ano_errado = [], [], [], []
        
        # Contadores para o painel analítico
        qtd_falta_ass, qtd_lro_feito_txt, qtd_nao_confeccionado, qtd_pendente, qtd_inexistente = 0, 0, 0, 0, 0

        # O rich.status faz uma animação giratória enquanto o código varre as pastas
        with console.status("[bold deep_sky_blue1]A varrer ficheiros e a extrair dados...", spinner="dots"):
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
                        relatorio.append(os.path.basename(tem_falta_ass[0])); problemas += 1; qtd_falta_ass += 1
                    elif tem_novo and tem_falta_lro:
                        lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno}); problemas += 1; qtd_lro_feito_txt += 1
                    elif tem_falta_lro:
                        relatorio.append(os.path.basename(tem_falta_lro[0])); problemas += 1; qtd_nao_confeccionado += 1
                    elif tem_novo:
                        lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno}); problemas += 1; qtd_pendente += 1
                    elif arquivos_errados:
                        lista_ano_errado.append(arquivos_errados[0]); problemas += 1
                    else:
                        relatorio.append(f"Dia {dia_fmt} - {turno}º Turno: NÃO ENCONTRADO"); problemas += 1; qtd_inexistente += 1
                        lista_para_criar.append({"str": data_str, "turno": turno, "dia": dia, "mes": mes, "ano": ano_longo})

        # 4. EXIBIÇÃO DO PAINEL DE STATUS
        console.print("\n")
        if problemas == 0:
            painel_status = Panel(
                "[bold green]✅ Todos os livros do mês estão OK e validados![/bold green]",
                title="[bold green]STATUS DO MÊS[/bold green]",
                border_style="green",
                padding=(1, 2),
                width=75
            )
            console.print(Align.center(painel_status))
        else:
            texto_status = Text()
            if qtd_pendente > 0: texto_status.append(f" 📄 {qtd_pendente} Livro(s) pendentes de verificação/assinatura\n", style="bold cyan")
            if qtd_lro_feito_txt > 0: texto_status.append(f" 🔄 {qtd_lro_feito_txt} Livro(s) verificados, mas o TXT de cobrança continua na pasta\n", style="bold yellow")
            if qtd_falta_ass > 0: texto_status.append(f" ✍️  {qtd_falta_ass} Livro(s) com ausência de assinatura\n", style="bold dark_orange")
            if qtd_nao_confeccionado > 0: texto_status.append(f" ⚠️  {qtd_nao_confeccionado} Livro(s) NÃO CONFECCIONADOS (Cobrança ativa)\n", style="bold magenta")
            if qtd_inexistente > 0: texto_status.append(f" ❌ {qtd_inexistente} Turno(s) INEXISTENTES (Necessário gerar cobrança)\n", style="bold red")
            if len(lista_ano_errado) > 0: texto_status.append(f" 📅 {len(lista_ano_errado)} Livro(s) com o ano incorreto no nome\n", style="bold red")
            
            painel_status = Panel(
                texto_status,
                title="[bold red]⚠️ RESUMO DE PENDÊNCIAS ENCONTRADAS[/bold red]",
                border_style="red",
                padding=(1, 2),
                width=75
            )
            console.print(Align.center(painel_status))

        corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado)

        if problemas > 0 and relatorio:

            # 5. GERAÇÃO AUTOMÁTICA DO TXT DE COBRANÇA (Invisível na tela)
            nome_relatorio = f"Relatorio_Pendencias_{mes}_{ano_longo}.txt"
            caminho_relatorio = os.path.join(path_mes, nome_relatorio)
            try:
                with open(caminho_relatorio, 'w', encoding='utf-8') as f:
                    f.write(f"--- RELATÓRIO DE PENDÊNCIAS LRO ({mes}/{ano_longo}) ---\n\n")
                    for l in relatorio: f.write(l + "\n")

                console.print(Align.center(f"[bold green]✅ Ficheiro '{nome_relatorio}' gerado e atualizado silenciosamente na pasta![/bold green]"))
            except Exception as e:
                console.print(Align.center(f"[bold red]Erro ao salvar TXT: {e}[/bold red]"))

        if lista_para_criar:
            console.print("\n")
            titulo_geracao = Text("GERAÇÃO INTELIGENTE DE ARQUIVOS 'FALTA LRO'", justify="center", style="bold white on deep_sky_blue1")
            console.print(Align.center(Panel(titulo_geracao, border_style="deep_sky_blue1", width=75)))
            
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

                def get_nome_adjacente(dt, tr, is_prev):
                    p_mes = get_path_mes(dt)
                    d_str = f"{dt.day:02d}{dt.strftime('%m')}{dt.strftime('%y')}"
                    arquivos = utils.buscar_arquivos_flexivel(p_mes, d_str, tr)
                    
                    pdfs = [f for f in arquivos if f.lower().endswith('.pdf')]
                    
                    if pdfs:
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
                
                # ========================================================
                # NOVO VISUAL PARA A GERAÇÃO DE ARQUIVO
                # ========================================================
                console.print()
                texto_falta = Text()
                texto_falta.append(f" ⬅️  Turno anterior passou para: ", style="dim white")
                texto_falta.append(f"{passou_nome}\n", style="bold cyan")
                texto_falta.append(f" ➡️  Turno seguinte recebeu de:  ", style="dim white")
                texto_falta.append(f"{recebeu_nome}", style="bold cyan")
                
                painel_falta = Panel(
                    texto_falta,
                    title=f"[bold yellow][FALTA LRO] {data_str} - {turno}º Turno[/bold yellow]",
                    border_style="yellow",
                    padding=(0, 2),
                    width=65
                )
                console.print(Align.center(painel_falta))
                
                if utils.pedir_confirmacao(f" " * 5 + f">> Criar ficheiro '{novo_nome}'? (S/Enter p/ Sim, ESC p/ Pular): "):
                    try: 
                        with open(os.path.join(path_mes, novo_nome), 'w') as f: f.write("Falta")
                        console.print(Align.center("[bold green]✅ Ficheiro criado com sucesso.[/bold green]"))
                    except Exception as e: 
                        console.print(Align.center(f"[bold red]Erro ao criar: {e}[/bold red]"))

        processo_verificacao_visual(lista_pendentes, mes, ano_curto)

        if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outro mês? (S/Enter p/ Sim, ESC p/ Voltar ao menu): {Cor.RESET}"): 
            break