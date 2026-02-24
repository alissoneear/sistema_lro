import json
import time
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.align import Align
from rich.text import Text
from rich.prompt import Prompt
from rich import box

import utils
from config import DadosEfetivo

console = Console()
CAMINHO_JSON = "efetivo.json"

def carregar_dados():
    return DadosEfetivo.inicializar_json()

def salvar_dados(dados):
    with open(CAMINHO_JSON, 'w', encoding='utf-8') as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

def ordenar_equipe(dados, equipe):
    dados[equipe] = sorted(dados[equipe], key=lambda x: x['legenda'])

def exibir_tabelas(dados):
    for equipe in ["SMC", "BCT", "OEA"]:
        cor = "cyan" if equipe == "SMC" else "green" if equipe == "BCT" else "dark_orange"
        
        tabela = Table(
            title=f"[bold {cor}]EFETIVO OPERACIONAL - {equipe}[/bold {cor}]", 
            border_style=cor,
            box=box.ROUNDED,
            header_style="bold white" 
        )
        
        tabela.add_column("ID", justify="center")
        tabela.add_column("LEG", justify="center", style=f"bold {cor}")
        tabela.add_column("Nome de Guerra", style="bold white")
        tabela.add_column("Posto e Nome Completo", style="bold white", no_wrap=True)
        tabela.add_column("SARAM", justify="center", style="white")

        for idx, militar in enumerate(dados[equipe]):
            tabela.add_row(
                str(idx + 1).zfill(2),
                militar["legenda"],
                militar["nome_guerra"],
                militar["nome_completo"],
                militar.get("saram", "") or "---"
            )
        console.print(Align.center(tabela))
        console.print()

def executar():
    espaco = " " * 18 

    while True:
        utils.limpar_tela()
        dados = carregar_dados()

        titulo = Text("SISTEMA LRO\nGestão de Legendas e Efetivo (Banco de Dados)", justify="center", style="bold dark_orange")
        console.print(Align.center(Panel(titulo, border_style="dark_orange", padding=(1, 2), width=65)))
        console.print()

        exibir_tabelas(dados)

        menu = Text()
        menu.append("  [1] ", style="bold dark_orange")
        menu.append("Adicionar Novo Militar\n")
        menu.append("  [2] ", style="bold dark_orange")
        menu.append("Editar Militar \n")
        menu.append("  [3] ", style="bold dark_orange")
        menu.append("Remover Militar \n\n")
        menu.append("  [0] ", style="bold red")
        menu.append("Voltar ao Menu Principal")

        painel_menu = Panel(
            menu, 
            title="[bold white]Menu de Operações[/bold white]", 
            title_align="left",
            border_style="dim white", 
            width=65,
            padding=(1, 2)
        )
        console.print(Align.center(painel_menu))
        console.print()

        opcao = Prompt.ask(" " * 22 + "[bold white]Opção desejada[/bold white]", choices=["0", "1", "2", "3"])

        if opcao == "0":
            break
            
        elif opcao == "1":
            console.print("\n")
            # Restaurado o uso do Prompt.ask com a lista original em maiúsculo
            equipe = Prompt.ask(espaco + "Adicionar em qual equipe?", choices=["SMC", "BCT", "OEA"])
                
            n_comp = console.input(espaco + "Posto e Nome Completo (Ex: 2S BCO JOÃO SILVA): ").strip().upper()
            n_guer = console.input(espaco + "Nome de Guerra (Ex: 2S SILVA): ").strip().upper()
            n_leg = console.input(espaco + "Legenda (Ex: A, B, GEA): ").strip().upper()
            n_sar = console.input(espaco + "SARAM (Deixe vazio se não souber): ").strip()

            if n_comp and n_guer and n_leg:
                dados[equipe].append({
                    "nome_completo": n_comp, "nome_guerra": n_guer, "legenda": n_leg, "saram": n_sar
                })
                ordenar_equipe(dados, equipe)
                salvar_dados(dados)
                console.print(f"\n{espaco}[bold green]✅ Militar adicionado com sucesso ao Banco de Dados![/bold green]")
            else:
                console.print(f"\n{espaco}[bold red]⚠️ Operação cancelada. Nome e Legenda são obrigatórios.[/bold red]")
            time.sleep(1.5)

        elif opcao == "2":
            console.print("\n")
            # Restaurado o uso do Prompt.ask com a lista original em maiúsculo
            equipe = Prompt.ask(espaco + "Editar militar de qual equipe?", choices=["SMC", "BCT", "OEA"])
                
            cor = "cyan" if equipe == "SMC" else "green" if equipe == "BCT" else "dark_orange"
            
            max_id = len(dados[equipe])
            if max_id == 0: continue
            
            id_edit = int(console.input(espaco + f"Digite o ID do militar para editar (1 a {max_id}): ")) - 1

            if 0 <= id_edit < max_id:
                m = dados[equipe][id_edit]
                console.print(f"\n{espaco}[dim white](Deixe em branco e aperte Enter para não alterar)[/dim white]")
                
                novo_comp = console.input(espaco + f"Nome Completo \[[bold {cor}]{m['nome_completo']}[/bold {cor}]]: ").strip().upper()
                novo_guer = console.input(espaco + f"Nome de Guerra \[[bold {cor}]{m['nome_guerra']}[/bold {cor}]]: ").strip().upper()
                nova_leg = console.input(espaco + f"Legenda \[[bold {cor}]{m['legenda']}[/bold {cor}]]: ").strip().upper()
                novo_sar = console.input(espaco + f"SARAM \[[bold {cor}]{m.get('saram', '')}[/bold {cor}]]: ").strip()

                alt_comp = novo_comp if novo_comp else m['nome_completo']
                alt_guer = novo_guer if novo_guer else m['nome_guerra']
                alt_leg = nova_leg if nova_leg else m['legenda']
                alt_sar = novo_sar if novo_sar else m.get('saram', '')

                if alt_comp == m['nome_completo'] and alt_guer == m['nome_guerra'] and alt_leg == m['legenda'] and alt_sar == m.get('saram', ''):
                    console.print(f"\n{espaco}[dim white][-] Nenhuma alteração detetada.[/dim white]")
                    time.sleep(1.5)
                    continue

                conflito = None
                nova_leg_conflito = None
                
                if alt_leg != m['legenda']:
                    conflito = next((x for i, x in enumerate(dados[equipe]) if x['legenda'] == alt_leg and i != id_edit), None)
                    
                    if conflito:
                        console.print(f"\n{espaco}[bold red]⚠️ A legenda '{alt_leg}' já pertence a {conflito['nome_guerra']}![/bold red]")
                        if utils.pedir_confirmacao(espaco + ">> Deseja forçar a troca? (S/Enter p/ Sim, ESC p/ Cancelar): "):
                            nova_leg_conflito = console.input(espaco + f">> Para qual legenda [bold {cor}]{conflito['nome_guerra']}[/bold {cor}] irá?: ").strip().upper()
                            if not nova_leg_conflito:
                                console.print(f"{espaco}[bold red]❌ Alteração cancelada. A nova legenda não pode ser vazia.[/bold red]")
                                time.sleep(1.5)
                                continue
                        else:
                            console.print(f"{espaco}[bold red]❌ Alteração cancelada.[/bold red]")
                            time.sleep(1.5)
                            continue

                resumo = Text()
                
                def add_linha_resumo(label, antigo, novo, is_last=False):
                    resumo.append(f" • {label}: ", style="bold white")
                    end_char = "" if is_last else "\n"
                    if antigo != novo:
                        resumo.append(f"{antigo} ➔ ", style="dim white")
                        resumo.append(f" {novo}{end_char}", style=f"bold {cor}")
                    else:
                        resumo.append(f"{antigo} (Mantido){end_char}", style="dim white")

                add_linha_resumo("Nome Completo", m['nome_completo'], alt_comp)
                add_linha_resumo("Nome de Guerra", m['nome_guerra'], alt_guer)
                add_linha_resumo("Legenda", m['legenda'], alt_leg)
                add_linha_resumo("SARAM", m.get('saram', '') or "---", alt_sar or "---", is_last=(conflito is None))

                if conflito and nova_leg_conflito:
                    # CORREÇÃO: Aplicado estilo via parâmetro (sem tags explícitas que quebravam)
                    resumo.append("\n\n⚠️ TROCA DE LEGENDA\n", style="bold yellow")
                    resumo.append(" • O militar ", style="bold white")
                    resumo.append(f"{conflito['nome_guerra']}", style=f"bold {cor}")
                    resumo.append(" terá a legenda alterada de ", style="bold white")
                    resumo.append(f"{conflito['legenda']}", style="dim white")
                    resumo.append(" ➔ ", style="bold white")
                    resumo.append(f" {nova_leg_conflito}", style=f"bold {cor}")

                painel_resumo = Panel(
                    resumo,
                    title=f"[bold {cor}]📋 RESUMO DA ALTERAÇÃO[/bold {cor}]",
                    border_style=cor,
                    padding=(1, 2),
                    width=85
                )
                
                console.print("\n")
                console.print(Align.center(painel_resumo))
                
                if utils.pedir_confirmacao(f"\n{espaco}>> Confirmar alterações no Efetivo? (S/Enter p/ Sim, ESC p/ Cancelar): "):
                    m['nome_completo'] = alt_comp
                    m['nome_guerra'] = alt_guer
                    m['legenda'] = alt_leg
                    m['saram'] = alt_sar
                    
                    if conflito and nova_leg_conflito:
                        conflito['legenda'] = nova_leg_conflito
                        
                    ordenar_equipe(dados, equipe)
                    salvar_dados(dados)
                    console.print(f"\n{espaco}[bold green]✅ Cadastro atualizado e reordenado com sucesso![/bold green]")
                else:
                    console.print(f"\n{espaco}[bold red]❌ Alteração cancelada pelo utilizador.[/bold red]")
                    
                time.sleep(1.5)

        elif opcao == "3":
            console.print("\n")
            # Restaurado o uso do Prompt.ask com a lista original em maiúsculo
            equipe = Prompt.ask(espaco + "Remover militar de qual equipe?", choices=["SMC", "BCT", "OEA"])
                
            max_id = len(dados[equipe])
            if max_id == 0: continue
            
            id_del = int(console.input(espaco + f"Digite o ID do militar para remover (1 a {max_id}): ")) - 1

            if 0 <= id_del < max_id:
                m = dados[equipe][id_del]
                
                # CORREÇÃO: Imprime a mensagem de atenção formatada com cor ANTES da pergunta
                console.print(f"\n{espaco}[bold red]🚨 ATENÇÃO: Deseja excluir permanentemente '{m['nome_guerra']}' do sistema?[/bold red]")
                
                if utils.pedir_confirmacao(espaco + ">> Confirmar exclusão? (S/Enter p/ Sim, ESC p/ Não): "):
                    dados[equipe].pop(id_del)
                    ordenar_equipe(dados, equipe)
                    salvar_dados(dados)
                    console.print(f"\n{espaco}[bold green]✅ Militar removido e escalas futuras ajustadas![/bold green]")
                    time.sleep(1.5)