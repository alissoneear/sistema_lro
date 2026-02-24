import os
import time
import datetime
import calendar

from config import Config, Cor
import utils
from modulos import verificador 

def executar():
    from rich.console import Console
    from rich.panel import Panel
    from rich.align import Align
    from rich.text import Text
    
    console = Console()

    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        console.print(Align.center(Panel(f"[bold red][ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.[/bold red]", border_style="red")))
        input()
        return

    while True:
        utils.limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        # ========================================================
        # 1. PAINEL DE TÍTULO DO MÓDULO
        # ========================================================
        titulo = Text("SISTEMA LRO\nConferência Rápida (Auxílio de Escala)", justify="center", style="bold dark_orange")
        painel_titulo = Panel(titulo, border_style="dark_orange", padding=(1, 2), width=65)
        console.print(Align.center(painel_titulo))
        console.print("\n")

        # ========================================================
        # 2. INPUT DE MÊS E ANO (Padrão Dashboard)
        # ========================================================
        if os.name == 'nt': 
            console.print(Align.center(f"[dim grey]Conectado: {Config.CAMINHO_RAIZ}[/dim grey]"))
            console.print()

        inp_mes = console.input(" " * 18 + f"[bold dark_orange]MÊS[/bold dark_orange] [dim white](Enter para {mes_atual}):[/dim white] ").strip()
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        
        inp_ano = console.input(" " * 18 + f"[bold dark_orange]ANO[/bold dark_orange] [dim white](Enter para {ano_atual_curto}):[/dim white] ").strip()
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        ano_longo = "20" + ano_curto

        path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
        path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
      
        if not os.path.exists(path_mes):
            console.print(Align.center(Panel(f"[bold red]Pasta do mês não encontrada.[/bold red]", border_style="red")))
            time.sleep(2)
            continue 

        try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
        except: continue
        if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

        # ========================================================
        # 3. PAINEL DE TRANSIÇÃO (Azul Cyan)
        # ========================================================
        utils.limpar_tela()
        titulo_analise = Text(f"INICIANDO CONFERÊNCIA: {Config.MAPA_PASTAS.get(mes)} / {ano_longo}", justify="center", style="bold white on deep_sky_blue1")
        console.print(Align.center(Panel(titulo_analise, border_style="deep_sky_blue1", width=65)))
        console.print(Align.center("[dim grey]Aperte S/Enter para PRÓXIMO ou ESC para SAIR.[/dim grey]\n"))
        time.sleep(1)

        contador = 0
        
        MESES_NOME = {
            "01": "janeiro", "02": "fevereiro", "03": "março",
            "04": "abril", "05": "maio", "06": "junho",
            "07": "julho", "08": "agosto", "09": "setembro",
            "10": "outubro", "11": "novembro", "12": "dezembro"
        }

        for dia in range(1, qtd_dias + 1):
            dia_fmt = f"{dia:02d}"
            data_str = f"{dia_fmt}{mes}{ano_curto}"
            
            for turno in [1, 2, 3]:
                arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, turno)
                if not arquivos: continue
                
                pdfs = [f for f in arquivos if f.lower().endswith('.pdf')]
                if not pdfs: continue
                
                contador += 1
                ok_files = [f for f in pdfs if "OK" in f.upper()]
                arquivo_alvo = ok_files[0] if ok_files else pdfs[0]
                
                mes_limpo = MESES_NOME.get(mes, "mês")
                data_formatada = f"📅 Dia {dia_fmt} de {mes_limpo} de 20{ano_curto} | {turno}º Turno"
                
                console.print("\n") 
                console.rule(f"[bold dark_orange]▶ CONFERÊNCIA - {os.path.basename(arquivo_alvo)}[/bold dark_orange]", style="dark_orange")
                
                info = utils.analisar_conteudo_lro(arquivo_alvo, mes, ano_curto)
                
                # 👈 O TRUQUE DE MESTRE! A Conferência agora pede o filtro do Verificador emprestado!
                info = verificador.enriquecer_info_lro(info, arquivo_alvo, mes, ano_curto)
                
                verificador.exibir_dados_analise(info, data_formatada)
                
                if not utils.pedir_confirmacao(f"\n>> Próximo turno? (S/Enter p/ Sim, ESC p/ Parar): "):
                    return

        console.print("\n")
        if contador == 0:
            console.print(Align.center(Panel("[bold red]Nenhum PDF encontrado para este período.[/bold red]", border_style="red")))
            time.sleep(2)
        else:
            console.print(Align.center(Panel(f"[bold green]✅ Conferência finalizada! {contador} turnos visualizados.[/bold green]", border_style="green", padding=(0,2))))
            time.sleep(2)
        
        if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outro mês? (S/Enter p/ Sim, ESC p/ Voltar): {Cor.RESET}"): 
            break