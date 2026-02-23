import os
import time
import datetime
import calendar
import re
import logging
from pypdf import PdfReader, PdfWriter

# Silencia avisos de fontes do pypdf para não poluir o terminal
logging.getLogger("pypdf").setLevel(logging.ERROR)

from config import Config, Cor, DadosEfetivo
import utils

def obter_info_militar(legenda, mapa_efetivo):
    for nome_completo, dados in mapa_efetivo.items():
        if dados['legenda'] == legenda:
            partes = nome_completo.split()
            if len(partes) >= 2:
                if partes[1] in ['BCT', 'BCO', 'QSS', 'SMC']:
                    return f"{partes[0]} {partes[2]}"
                return f"{partes[0]} {partes[1]}"
            return nome_completo
    return legenda

def get_sem(ano, mes, dia):
    dt = datetime.date(int(ano), int(mes), int(dia))
    return Config.MAPA_SEMANA[dt.weekday()]

def imprimir_tabela(escala_detalhada, qtd_dias, opcao_escala, ano_longo, mes):
    """Função auxiliar para imprimir a tabela da escala formatada com Rich."""
    from rich.table import Table
    from rich.console import Console
    from rich.align import Align
    from rich import box
    import datetime
    from config import Config
    
    console = Console()

    print("\n")
    
    especialidade_nome = "SMC" if opcao_escala == '1' else "BCT" if opcao_escala == '2' else "OEA"

    tabela = Table(
        title=f"[bold dark_orange]ESCALA CUMPRIDA {especialidade_nome} - {mes}/{ano_longo}[/bold dark_orange]",
        header_style="bold dark_orange",
        border_style="grey37",
        box=box.ROUNDED,
        show_lines=False, 
        padding=(0, 2)    
    )

    tabela.add_column("DIAS", justify="center")
    tabela.add_column("SEM", justify="center")

    if opcao_escala == '1':
        tabela.add_column("SMC", justify="center")
        tracos = 19
    else:
        tabela.add_column("1ºTURNO", justify="center")
        tabela.add_column("2ºTURNO", justify="center")
        tabela.add_column("3ºTURNO", justify="center")
        tracos = 50

    for dia in range(1, qtd_dias + 1):
        dt = datetime.date(int(ano_longo), int(mes), dia)
        sigla_sem = Config.MAPA_SEMANA[dt.weekday()]
        
        dia_str = f"{dia:02d}"

        if dt.weekday() in [5, 6]:
            estilo_linha = "turquoise4"
        else:
            estilo_linha = "white"

        def formatar(leg):
            if leg in ['---', '???', 'ERR', 'PND']: 
                return f"[red3]{leg}[/red3]"
            return f"[{estilo_linha}]{leg}[/{estilo_linha}]"

        if opcao_escala == '1':
            tabela.add_row(
                f"[{estilo_linha}]{dia_str}[/{estilo_linha}]",
                f"[{estilo_linha}]{sigla_sem}[/{estilo_linha}]",
                formatar(escala_detalhada[dia]['smc'])
            )
        else:
            tabela.add_row(
                f"[{estilo_linha}]{dia_str}[/{estilo_linha}]",
                f"[{estilo_linha}]{sigla_sem}[/{estilo_linha}]",
                formatar(escala_detalhada[dia][1]['legenda']),
                formatar(escala_detalhada[dia][2]['legenda']),
                formatar(escala_detalhada[dia][3]['legenda'])
            )

    console.print(Align.center(tabela))
    return tracos

def realizar_auditoria_manual(escala_detalhada, mes, ano_curto, path_mes, opcao_escala, mapa_ativo, alertas_suspeitos=None, caminho_cache=None):
    """Procura falhas na extração OU alertas da auditoria e abre o PDF exibindo os motivos."""
    from rich.console import Console
    from rich.panel import Panel
    from rich.align import Align
    from rich.text import Text
    import os
    
    console = Console()

    if alertas_suspeitos is None: alertas_suspeitos = {}
    pendentes = []
    dias = sorted(escala_detalhada.keys())
    
    # Mapear os campos que precisam de atenção e seus motivos
    for dia in dias:
        if opcao_escala == '1':
            leg = escala_detalhada[dia]['smc']
            motivos = []
            if leg in ['---', '???', 'ERR', 'PND']:
                motivos.append("Falha na extração dos dados (ilegível ou ausente).")
            if (dia, 1) in alertas_suspeitos:
                motivos.extend(alertas_suspeitos[(dia, 1)])
                
            if motivos:
                pendentes.append((dia, 1, leg, motivos))
        else:
            for t in [1, 2, 3]:
                leg = escala_detalhada[dia][t]['legenda']
                motivos = []
                if leg in ['---', '???', 'ERR', 'PND']:
                    motivos.append("Falha na extração dos dados (ilegível ou ausente).")
                if (dia, t) in alertas_suspeitos:
                    motivos.extend(alertas_suspeitos[(dia, t)])
                    
                if motivos:
                    pendentes.append((dia, t, leg, motivos))
                    
    if not pendentes:
        return False 
        
    # ========================================================
    # 1. PAINEL DE INÍCIO DA AUDITORIA
    # ========================================================
    console.print("\n")
    painel_auditoria = Panel(
        f"[bold white]Existem [dark_orange]{len(pendentes)}[/dark_orange] turnos pendentes de revisão.[/bold white]",
        title="[bold dark_orange]🔎 INICIAR AUDITORIA MANUAL[/bold dark_orange]",
        border_style="dark_orange",
        padding=(1, 2),
        width=75
    )
    console.print(Align.center(painel_auditoria))
    console.print()

    import utils
    if not utils.pedir_confirmacao(f" " * 5 + f">> Deseja realizar a AUDITORIA MANUAL agora? (S/Enter p/ Sim, ESC p/ Pular): "):
        return False
        
    validas = [dados['legenda'] for dados in mapa_ativo.values()]
    
    # ========================================================
    # 2. CRIAÇÃO DO MAPA VISUAL DE LEGENDAS (Padrão Rich)
    # ========================================================
    nome_escala = "SMC" if opcao_escala == '1' else "BCT" if opcao_escala == '2' else "OEA"
    cor_titulo = "cyan" if opcao_escala == '1' else "green" if opcao_escala == '2' else "dark_orange"

    def gerar_texto_mapa(nome_esc, mapa, cor):
        txt = Text()
        txt.append(f"■ EQUIPE {nome_esc}:\n", style=f"bold {cor}")
        itens = [f"[{v['legenda']}] {k.split('-')[0].strip()}" for k, v in mapa.items()]
        
        linhas = []
        for i in range(0, len(itens), 4):
            linha = itens[i:i+4]
            linhas.append("   " + "".join(item.ljust(22) for item in linha))
            
        txt.append("\n".join(linhas), style="white")
        return txt

    mapa_visual_text = gerar_texto_mapa(nome_escala, mapa_ativo, cor_titulo)
    painel_mapa = Panel(
        mapa_visual_text,
        title="[bold yellow]✏️ MAPA DE LEGENDAS[/bold yellow]",
        border_style="yellow",
        padding=(1, 2)
    )
    
    modificado = False
    
    MESES_NOME = {
        "01": "janeiro", "02": "fevereiro", "03": "março",
        "04": "abril", "05": "maio", "06": "junho",
        "07": "julho", "08": "agosto", "09": "setembro",
        "10": "outubro", "11": "novembro", "12": "dezembro"
    }
    
    for dia, t, leg_atual, motivos in pendentes:
        dia_fmt = f"{dia:02d}"
        data_str = f"{dia_fmt}{mes}{ano_curto}"
        mes_limpo = MESES_NOME.get(mes, "mês")
        
        # ========================================================
        # 3. RÉGUA SEPARADORA E MOTIVOS DE AUDITORIA
        # ========================================================
        console.print("\n")
        if opcao_escala == '1':
            console.rule(f"[bold deep_sky_blue1]▶ Auditando: Dia {dia_fmt} de {mes_limpo} - SMC [/bold deep_sky_blue1]", style="deep_sky_blue1")
            arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, 1) 
        else:
            console.rule(f"[bold deep_sky_blue1]▶ Auditando: Dia {dia_fmt} de {mes_limpo} - {t}º Turno [/bold deep_sky_blue1]", style="deep_sky_blue1")
            arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, t)
            
        console.print()
        for m in motivos:
            console.print(Align.center(f"[bold red]📌 MOTIVO:[/bold red] [white]{m}[/white]"))
        console.print()
            
        pdfs = [f for f in arquivos if f.lower().endswith('.pdf')]
        arquivo_alvo = None
        if pdfs:
            ok_files = [f for f in pdfs if "OK" in f.upper()]
            arquivo_alvo = ok_files[0] if ok_files else pdfs[0]
            
        if arquivo_alvo:
            from config import Cor
            console.print(Align.center(f"[dim grey]Abrindo o documento: {os.path.basename(arquivo_alvo)}...[/dim grey]"))
            utils.abrir_arquivo(arquivo_alvo)
        else:
            console.print(Align.center("[bold red][!] Nenhum PDF encontrado para este turno.[/bold red]"))
            
        # ========================================================
        # 4. EXIBIÇÃO DO MAPA E INPUT DO UTILIZADOR
        # ========================================================
        console.print("\n")
        console.print(Align.center(painel_mapa))
        console.print(Align.center("[dim grey](Deixe em branco e aperte Enter para manter a legenda atual)[/dim grey]\n"))
        
        nova_leg = input(f"   Digite a nova legenda ou Enter p/ manter [{leg_atual}]: ").strip().upper()
        
        if nova_leg:
            if nova_leg in validas:
                if opcao_escala == '1':
                    escala_detalhada[dia]['smc'] = nova_leg
                else:
                    escala_detalhada[dia][t]['legenda'] = nova_leg
                modificado = True
                console.print(f"\n[bold green]   ✅ Legenda atualizada para: {nova_leg}[/bold green]")

                # SALVAMENTO AUTOMÁTICO NO CACHE
                if caminho_cache:
                    import json
                    try:
                        with open(caminho_cache, 'w', encoding='utf-8') as f: json.dump(escala_detalhada, f, indent=4)
                    except: pass
            else:
                console.print(f"\n[bold red]   ⚠️ Legenda '{nova_leg}' inválida. A manter '{leg_atual}'.[/bold red]")
        else:
            console.print(f"\n[dim grey]   Mantido: {leg_atual}[/dim grey]")
            
    return modificado

def verificar_e_propor_correcoes(escala_detalhada, mapa_efetivo, ano, mes):
    inconsistencias = [] 
    correcoes = []       # Não usaremos mais auto-correção, deixaremos a lista vazia
    alertas_manuais = {} # Dicionário: (dia, turno) -> [motivos]
    dias = sorted(escala_detalhada.keys())
    ignorar = ['---', '???', 'PND', 'ERR']
    
    def obter_legenda_pelo_nome(nome_guerra):
        return utils.encontrar_legenda(nome_guerra, mapa_efetivo)
        
    def add_alerta(d, t, msg):
        if (d, t) not in alertas_manuais:
            alertas_manuais[(d, t)] = []
        if msg not in alertas_manuais[(d, t)]:
            alertas_manuais[(d, t)].append(msg)

    for i, dia in enumerate(dias):
        dados_dia = escala_detalhada[dia]
        dia_sem = get_sem(ano, mes, dia)
        l1, l2, l3 = dados_dia[1]['legenda'], dados_dia[2]['legenda'], dados_dia[3]['legenda']
        
        # --- REGRA 1: DOBRA DE TURNO ---
        violacao = None
        if l1 not in ignorar and l1 == l2: violacao = (1, 2, l1)
        elif l2 not in ignorar and l2 == l3: violacao = (2, 3, l2)

        if violacao:
            t_a, t_b, leg = violacao
            nome_mil = obter_info_militar(leg, mapa_efetivo)
            
            # Motivos para a revisão visual
            motivo_dobra = f"Possível dobra de turno detetada ({t_a}º e {t_b}º) do militar {nome_mil} ({leg})."
            add_alerta(dia, t_a, motivo_dobra)
            add_alerta(dia, t_b, motivo_dobra)
            
            # Apenas relata o problema (Sem tentar adivinhar a solução)
            inconsistencias.append(f"Dia {dia:02d} ({dia_sem}): Militar {nome_mil} ({leg}) dobrou o turno ({t_a}º e {t_b}º).")

        # --- REGRA 2: FOLGA PÓS-3º TURNO ---
        if i > 0:
            dia_ant = dias[i-1]
            dia_sem_ant = get_sem(ano, mes, dia_ant)
            l3_ant = escala_detalhada[dia_ant][3]['legenda']
            
            if l3_ant not in ignorar:
                turnos_hoje_violados = [t for t in [1, 2, 3] if escala_detalhada[dia][t]['legenda'] == l3_ant]
                if turnos_hoje_violados:
                    nome_mil = obter_info_militar(l3_ant, mapa_efetivo)

                    for t_hoje in turnos_hoje_violados:
                        motivo_folga_hoje = f"Falta de folga regulamentar. O militar {nome_mil} ({l3_ant}) estava escalado no 3º Turno do dia {dia_ant:02d}."
                        motivo_folga_ontem = f"Militar {nome_mil} ({l3_ant}) escalado aqui, mas aparece sem folga no {t_hoje}º Turno do dia seguinte ({dia:02d})."
                        
                        add_alerta(dia, t_hoje, motivo_folga_hoje)
                        add_alerta(dia_ant, 3, motivo_folga_ontem)

                    # Apenas relata o problema
                    inconsistencias.append(f"Dia {dia:02d} ({dia_sem}): Militar {nome_mil} ({l3_ant}) sem folga do dia {dia_ant:02d} ({dia_sem_ant}).")
                    
    return inconsistencias, correcoes, alertas_manuais

def gerar_pdf_escala(escala_detalhada, mapa_ativo, opcao_escala, mes, ano_longo):
    """
    Busca o template, calcula horas com arredondamento e salva numa pasta de saídas dedicada.
    """
    especialidade = "bct" if opcao_escala == '2' else "oea"
    
    MESES = {
        "01": ("jan", "JANEIRO"), "02": ("fev", "FEVEREIRO"), "03": ("mar", "MARÇO"),
        "04": ("abr", "ABRIL"), "05": ("mai", "MAIO"), "06": ("jun", "JUNHO"),
        "07": ("jul", "JULHO"), "08": ("ago", "AGOSTO"), "09": ("set", "SETEMBRO"),
        "10": ("out", "OUTUBRO"), "11": ("nov", "NOVEMBRO"), "12": ("dez", "DEZEMBRO")
    }
    mes_abrev, mes_longo = MESES.get(mes, ("xxx", "XXX"))
    
    caminho_template = os.path.join(
        "templates", especialidade, ano_longo, mes_abrev, 
        f"{especialidade}_{ano_longo}_{mes_abrev}_temp.pdf"
    )
    
    pasta_saida = os.path.join("SAIDAS_PDF", especialidade, ano_longo, mes_abrev)
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)
    
    if not os.path.exists(caminho_template):
        print(f"{Cor.RED}[!] Operação cancelada: Template não encontrado em: {caminho_template}{Cor.RESET}")
        return
        
    dados_pdf = {}
    dados_pdf["mes_ano"] = f"CUMPRIDA {mes_longo}/{ano_longo}"
    # Formatação com zfill(2) para garantir sempre 2 dígitos
    dados_pdf["ef_total"] = str(len(mapa_ativo)).zfill(2)
    
    minutos_por_turno = {1: 435, 2: 435, 3: 615}
    horas_militares = {}
    
    for dia, turnos_dia in escala_detalhada.items():
        # NOVO: Preenche o dia da semana na variável que o PDF espera (ex: d1_sem = 'SEG')
        dados_pdf[f"d{dia}_sem"] = get_sem(ano_longo, mes, dia)

        for t in [1, 2, 3]:
            leg = turnos_dia[t]['legenda']
            if leg not in ['---', 'PND', 'ERR', '???']:
                dados_pdf[f"d{dia}_t{t}"] = leg
                horas_militares[leg] = horas_militares.get(leg, 0) + minutos_por_turno[t]
            else:
                dados_pdf[f"d{dia}_t{t}"] = ""
                
    # NOVO: Se o mês tiver menos de 31 dias, limpa os campos excedentes do PDF.
    # Assim, o mesmo template com 31 linhas serve para Fevereiro (28 dias) e meses de 30 dias.
    for dia_extra in range(len(escala_detalhada) + 1, 32):
        dados_pdf[f"d{dia_extra}_sem"] = ""
        for t in [1, 2, 3]:
            dados_pdf[f"d{dia_extra}_t{t}"] = ""

    # Formatação com zfill(2) para garantir sempre 2 dígitos
    dados_pdf["ef_escala"] = str(len(horas_militares)).zfill(2)
    
    campos_horas = [f"horas_{chr(i)}" for i in range(97, 112)] 
    idx = 0
    
    for leg in sorted(horas_militares.keys()):
        if idx >= len(campos_horas): break
        
        m_totais = horas_militares[leg]
        h_inteiras = m_totais // 60
        m_restantes = m_totais % 60
        
        h_finais = h_inteiras + 1 if m_restantes >= 30 else h_inteiras
        
        dados_pdf[campos_horas[idx]] = f"{leg} - {h_finais}h"
        idx += 1
        
    while idx < len(campos_horas):
        dados_pdf[campos_horas[idx]] = ""
        idx += 1
        
    print(f"\n{Cor.GREY}A preencher o documento oficial...{Cor.RESET}")
    reader = PdfReader(caminho_template)
    writer = PdfWriter()
    writer.append(reader)
    writer.update_page_form_field_values(writer.pages[0], dados_pdf)
    
    nome_saida = f"ESCALA_CUMPRIDA_{especialidade.upper()}_{mes_abrev.upper()}_{ano_longo}.pdf"
    caminho_saida = os.path.join(pasta_saida, nome_saida)
    
    try:
        with open(caminho_saida, "wb") as f:
            writer.write(f)
        print(f"{Cor.GREEN}✅ PDF gerado com sucesso!{Cor.RESET}")
        print(f"{Cor.CYAN}Guardado em: {caminho_saida}{Cor.RESET}")
        
        print(f"\n{Cor.YELLOW}O que deseja fazer agora?{Cor.RESET}")
        print("  [1] Abrir o PDF gerado")
        print("  [2] Abrir a Pasta onde o PDF foi salvo")
        print("  [0] Voltar ao Menu")
        escolha = input(f">> Opção: ").strip()
        
        if escolha == '1':
            utils.abrir_arquivo(caminho_saida)
        elif escolha == '2':
            if os.name == 'nt': os.startfile(pasta_saida)
            elif sys.platform == 'darwin': os.system(f'open "{pasta_saida}"')
            else: os.system(f'xdg-open "{pasta_saida}"')

    except Exception as e:
        print(f"{Cor.RED}[!] Erro ao gravar o ficheiro PDF: {e}{Cor.RESET}")

def executar():
    import os
    import sys
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
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    while True:
        utils.limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        # ========================================================
        # 1. PAINEL DE TÍTULO DO MÓDULO
        # ========================================================
        titulo = Text("SISTEMA LRO\nGerador de Escala Cumprida e Auditoria", justify="center", style="bold dark_orange")
        painel_titulo = Panel(titulo, border_style="dark_orange", padding=(1, 2), width=65)
        console.print(Align.center(painel_titulo))
        console.print("\n")

        # ========================================================
        # 2. INPUT DE MÊS E ANO (Modernizado)
        # ========================================================
        inp_mes = console.input(" " * 18 + f"[bold dark_orange]MÊS[/bold dark_orange] [dim white](Enter para {mes_atual}):[/dim white] ").strip()
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        
        inp_ano = console.input(" " * 18 + f"[bold dark_orange]ANO[/bold dark_orange] [dim white](Enter para {ano_atual_curto}):[/dim white] ").strip()
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        ano_longo = "20" + ano_curto

        mapa_smc, mapa_bct, mapa_oea = DadosEfetivo.mapear_efetivo(mes, ano_curto)

        while True:
            utils.limpar_tela()
            
            titulo_menu = Text(f"PERÍODO SELECIONADO: {mes}/{ano_longo}", justify="center", style="bold white on dark_orange")
            console.print(Align.center(Panel(titulo_menu, border_style="dark_orange", width=65)))
            console.print("\n")
            
            # ========================================================
            # 3. MENU DE ESPECIALIDADES EM PAINEL
            # ========================================================
            menu_opcoes = Text()
            menu_opcoes.append("  [1] ", style="bold dark_orange")
            menu_opcoes.append("Escala Cumprida - SMC\n")
            menu_opcoes.append("  [2] ", style="bold dark_orange")
            menu_opcoes.append("Escala Cumprida - BCT\n")
            menu_opcoes.append("  [3] ", style="bold dark_orange")
            menu_opcoes.append("Escala Cumprida - OEA\n\n")
            menu_opcoes.append("  [0] ", style="bold red")
            menu_opcoes.append("Voltar ao Menu Principal")

            painel_menu = Panel(
                menu_opcoes, 
                border_style="dim white", 
                title="[bold white]Selecione a Especialidade[/bold white]", 
                title_align="left",
                width=65,
                padding=(1, 2)
            )
            console.print(Align.center(painel_menu))
            console.print()

            opcao_escala = Prompt.ask(" " * 22 + "[bold white]Opção desejada[/bold white]", choices=["0", "1", "2", "3"])
            
            if opcao_escala == '0': return 
            if opcao_escala not in ['1', '2', '3']: continue

            path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
            path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
            
            if not os.path.exists(path_mes):
                console.print(Align.center(Panel(f"[bold red]Pasta não encontrada para o período {mes}/{ano_longo}.[/bold red]", border_style="red")))
                time.sleep(2)
                break

            try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
            except: break
            if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

            # ========================================================
            # 4. DETEÇÃO DO CACHE E PERGUNTA AO USUÁRIO
            # ========================================================
            import json
            especialidade_nome = 'smc' if opcao_escala == '1' else 'bct' if opcao_escala == '2' else 'oea'
            caminho_cache = os.path.join(path_mes, f".cache_escala_{especialidade_nome}.json")
            cache_dados = {}
            
            if os.path.exists(caminho_cache):
                utils.limpar_tela()
                
                alerta_cache = Text()
                alerta_cache.append(f"Encontramos auditorias manuais feitas anteriormente para a escala de {especialidade_nome.upper()}.\n\n", style="dim white")
                alerta_cache.append("Como deseja prosseguir com a geração da tabela?", style="bold yellow")
                
                painel_cache = Panel(
                    alerta_cache,
                    title="[bold yellow]💾 PROGRESSO SALVO DETETADO![/bold yellow]",
                    border_style="yellow",
                    padding=(1, 2),
                    width=65
                )
                console.print(Align.center(painel_cache))
                console.print("\n")
                
                if utils.pedir_confirmacao(f"{Cor.CYAN}>> Deseja RETOMAR de onde parou? (S/Enter p/ Sim, ESC p/ Reiniciar do zero): {Cor.RESET}"):
                    try:
                        with open(caminho_cache, 'r', encoding='utf-8') as f:
                            cache_dados = json.load(f)
                        print(f"{Cor.GREEN}✅ Progresso carregado com sucesso!{Cor.RESET}")
                        time.sleep(1)
                    except: pass
                else:
                    os.remove(caminho_cache)
                    print(f"{Cor.GREEN}Cache apagado. Iniciando do zero!{Cor.RESET}")
                    time.sleep(1)

            print(f"\n{Cor.GREY}Processando os dados e auditando as inconsistências... Aguarde.{Cor.RESET}\n")

            escala_detalhada = {}
            mapa_ativo = mapa_bct if opcao_escala == '2' else mapa_oea 

            # CONFIGURAÇÃO DA BARRA DE PROGRESSO FLUIDA (RICH PREMIUM)
            from rich.progress import Progress, TextColumn, BarColumn, TaskProgressColumn, SpinnerColumn, ProgressColumn
            from rich.text import Text

            # Criando medidores de tempo 100% customizados e em português!
            class TempoDecorridoBR(ProgressColumn):
                def render(self, task):
                    elapsed = task.elapsed
                    if elapsed is None: return Text("0 seg", style="bold yellow")
                    mins, secs = divmod(int(elapsed), 60)
                    return Text(f"{mins} min {secs} seg" if mins > 0 else f"{secs} seg", style="bold yellow")

            class TempoRestanteBR(ProgressColumn):
                def render(self, task):
                    remaining = task.time_remaining
                    if remaining is None: return Text("--", style="bold deep_sky_blue1")
                    mins, secs = divmod(int(remaining), 60)
                    return Text(f"{mins} min {secs} seg" if mins > 0 else f"{secs} seg", style="bold deep_sky_blue1")

            total_passos = qtd_dias * 3
            
            progress = Progress(
                SpinnerColumn("dots", style="bold dark_orange"),
                TextColumn("[bold white]{task.description}"),
                BarColumn(
                    bar_width=30, # Deixei a barra um pouco mais compacta para caber os textos novos
                    style="grey37",                      
                    complete_style="bold dark_orange"    
                ),
                TaskProgressColumn(text_format="[bold cyan]{task.percentage:>3.0f}%"), 
                TextColumn("[dim grey]• Decorrido:[/dim grey]"),
                TempoDecorridoBR(),                      # 👈 Nossa coluna customizada!
                TextColumn("[dim grey]• Restante:[/dim grey]"),
                TempoRestanteBR()                        # 👈 Nossa coluna customizada!
            )
            tarefa_extracao = progress.add_task("Extraindo dados...", total=total_passos)
            
            progress.start()

            # Extração de Dados
            for dia in range(1, qtd_dias + 1):
                dia_fmt = f"{dia:02d}"
                data_str = f"{dia_fmt}{mes}{ano_curto}"
                turnos = utils.calcular_turnos_validos(dia, mes, agora.day, mes_atual, agora.hour)

                dia_dados = {'smc': '---', 'bct': {1:'---',2:'---',3:'---'}, 'oea': {1:'---',2:'---',3:'---'},
                             'meta': {1:{'assinatura_nome':'???'}, 2:{'assinatura_nome':'???'}, 3:{'assinatura_nome':'???'}}}

                for turno in [1, 2, 3]:
                    progress.update(tarefa_extracao, advance=1)
                    usou_cache = False

                    dia_str, turno_str = str(dia), str(turno)
                    
                    if cache_dados and dia_str in cache_dados and turno_str in cache_dados[dia_str]:
                        leg_salva = cache_dados[dia_str][turno_str].get('legenda', '---')
                        ass_salva = cache_dados[dia_str][turno_str].get('assinatura_nome', '???')
                        
                        if leg_salva not in ['---', '???', 'ERR', 'PND']:
                            if opcao_escala == '2': dia_dados['bct'][turno] = leg_salva
                            elif opcao_escala == '3': dia_dados['oea'][turno] = leg_salva
                            dia_dados['meta'][turno]['assinatura_nome'] = ass_salva
                            
                            if opcao_escala == '1' and 'smc' in cache_dados[dia_str]:
                                dia_dados['smc'] = cache_dados[dia_str]['smc']
                                
                            continue

                    if turno not in turnos: 
                        continue
                        
                    arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, turno)
                    # ... [O resto do código para baixo continua igual]
                    if not arquivos: continue
                    pdfs = [f for f in arquivos if f.lower().endswith('.pdf')]
                    if not pdfs:
                        if any("FALTA LRO" in f.upper() for f in arquivos):
                            dia_dados['bct'][turno] = 'PND'; dia_dados['oea'][turno] = 'PND'
                            dia_dados['meta'][turno]['assinatura_nome'] = 'PND'
                        continue

                    arquivo_alvo = [f for f in pdfs if "OK" in f.upper()][0] if [f for f in pdfs if "OK" in f.upper()] else pdfs[0]
                    info = utils.analisar_conteudo_lro(arquivo_alvo, mes, ano_curto)
                    if info:
                        dia_dados['smc'] = utils.encontrar_legenda(info['equipe']['smc'], mapa_smc)
                        dia_dados['bct'][turno] = utils.encontrar_legenda(info['equipe']['bct'], mapa_bct)
                        dia_dados['oea'][turno] = utils.encontrar_legenda(info['equipe']['oea'], mapa_oea)
                        
                        resp_base = utils.extrair_nome_base(info.get('responsavel', ''))
                        dia_dados['meta'][turno]['assinatura_nome'] = resp_base
                    else:
                        dia_dados['bct'][turno] = 'ERR'; dia_dados['oea'][turno] = 'ERR'

                escala_detalhada[dia] = {}
                for t in [1, 2, 3]:
                    leg = dia_dados['bct'][t] if opcao_escala == '2' else dia_dados['oea'][t] if opcao_escala == '3' else '---'
                    escala_detalhada[dia][t] = {'legenda': leg, 'assinatura_nome': dia_dados['meta'][t]['assinatura_nome']}
                    if opcao_escala == '1': escala_detalhada[dia]['smc'] = dia_dados['smc']
            
            progress.stop() # Finaliza a animação da barra do Rich
            print("\n") # <- FIM DA BARRA DE PROGRESSO AQUI

            # SALVA O ESTADO INICIAL NO CACHE
            if caminho_cache:
                import json
                try:
                    with open(caminho_cache, 'w', encoding='utf-8') as f: json.dump(escala_detalhada, f, indent=4)
                except: pass

            # --- 1. ALERTA DE SEGURANÇA OPERACIONAL (RADAR) COM RICH ---
            if opcao_escala in ['2', '3']:
                inconsistencias, _, _ = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)
                
                from rich.panel import Panel
                from rich.text import Text
                from rich.align import Align
                from rich.console import Console
                console_rich = Console()
                
                if inconsistencias:
                    texto_alerta = Text()
                    for inc in inconsistencias: 
                        if "sem folga" in inc:
                            texto_alerta.append(f" {inc}\n", style="bold red")
                        elif "dobrou" in inc:
                            texto_alerta.append(f" {inc}\n", style="bold yellow")
                            
                    texto_alerta.append("\nℹ️  Nota: O sistema identificou quebra de descanso. Recomendada revisão na Auditoria Manual.", style="dim white")
                    
                    painel = Panel(
                        texto_alerta, 
                        title="[bold white on red] 🚨 ANÁLISE PRÉVIA DE CONSISTÊNCIA [/bold white on red]", 
                        border_style="red", 
                        padding=(1, 2)
                    )
                    console_rich.print(Align.center(painel))
                else:
                    painel = Panel(
                        "[bold green]✅ Nenhuma quebra de descanso (folga/dobra) identificada na escala.[/bold green]", 
                        title="[bold green] 🔍 ANÁLISE PRÉVIA DE CONSISTÊNCIA [/bold green]", 
                        border_style="green", 
                        padding=(1, 2)
                    )
                    console_rich.print(Align.center(painel))

            # --- 2. LOOP DE AUDITORIA MANUAL INFINITO ---
            while True:
                alertas_ativos = {}
                if opcao_escala in ['2', '3']:
                    _, _, alertas_ativos = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)

                teve_correcao_manual = realizar_auditoria_manual(
                    escala_detalhada, mes, ano_curto, path_mes, opcao_escala, mapa_ativo, alertas_ativos, caminho_cache
                )
                
                if teve_correcao_manual:
                    utils.limpar_tela()
                    print(f"{Cor.ORANGE}=== SISTEMA LRO - Escala Cumprida ({mes}/{ano_curto}) ==={Cor.RESET}")
                    print(f"\n{Cor.GREEN}✅ Tabela atualizada após Auditoria Manual:{Cor.RESET}")
                    tracos = imprimir_tabela(escala_detalhada, qtd_dias, opcao_escala, ano_longo, mes)
                else:
                    break

            # --- 3. RESUMO FINAL E GERAÇÃO DO PDF COM RICH ---
            if opcao_escala in ['2', '3']:
                inc_final, _, _ = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)
                
                from rich.panel import Panel
                from rich.text import Text
                from rich.align import Align
                from rich.console import Console
                console_rich = Console()
                
                print("\n")
                if inc_final:
                    texto_final = Text()
                    for inc in inc_final: 
                        texto_final.append(f" {inc}\n", style="bold red")
                        
                    painel_final = Panel(
                        texto_final, 
                        title="[bold red]RESUMO DA AUDITORIA OPERACIONAL[/bold red]", 
                        border_style="red", 
                        padding=(1, 2)
                    )
                else:
                    painel_final = Panel(
                        "[bold green]✅ Escala validada e consistente com as normas de folga.[/bold green]", 
                        title="[bold green]RESUMO DA AUDITORIA OPERACIONAL[/bold green]",
                        border_style="green", 
                        padding=(1, 2)
                    )
                    
                console_rich.print(Align.center(painel_final))
                
                if utils.pedir_confirmacao(f"\n{Cor.CYAN}>> Deseja GERAR O PDF OFICIAL desta Escala Cumprida? (S/Enter p/ Sim, ESC p/ Pular): {Cor.RESET}"):
                    gerar_pdf_escala(escala_detalhada, mapa_ativo, opcao_escala, mes, ano_longo)

            if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outra especialidade? (S/Enter p/ Sim, ESC p/ Voltar): {Cor.RESET}"): return