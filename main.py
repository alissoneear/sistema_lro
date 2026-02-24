import time
import sys

from rich.console import Console
from rich.panel import Panel
from rich.align import Align
from rich.text import Text
from rich.prompt import Prompt

from config import Cor
import utils

from modulos import verificador
from modulos import conferencia
from modulos import escala
from modulos import gestor_efetivo

console = Console()

def menu_principal():
    while True:
        utils.limpar_tela()
        
        # 1. PAINEL DE TÍTULO (Tema Dark Orange)
        titulo = Text("SISTEMA LRO\nCentro de Coordenação de Salvamento (ARCC-CW)", justify="center", style="bold dark_orange")
        painel_titulo = Panel(titulo, border_style="dark_orange", padding=(1, 2), width=65)
        console.print(Align.center(painel_titulo))
        
        console.print("\n")
        
        # 2. DASHBOARD
        dashboard_txt = utils.gerar_dashboard_boas_vindas()
        console.print(Align.center(Text.from_ansi(dashboard_txt)))
        
        console.print("\n")

        # 3. PAINEL DE OPÇÕES
        menu_opcoes = Text()
        menu_opcoes.append("  [1] ", style="bold dark_orange")
        menu_opcoes.append("Verificador LRO (Validar/Assinar)\n")
        menu_opcoes.append("  [2] ", style="bold dark_orange")
        menu_opcoes.append("Conferência Rápida (Auxiliar Escala)\n")
        menu_opcoes.append("  [3] ", style="bold dark_orange")
        menu_opcoes.append("Escala Cumprida (Gerar Tabela/Auditoria)\n")
        menu_opcoes.append("  [4] ", style="bold dark_orange")
        menu_opcoes.append("Gestão de Efetivo (Alterar Legendas)\n\n")
        menu_opcoes.append("  [0] ", style="bold red")
        menu_opcoes.append("Sair")

        painel_menu = Panel(
            menu_opcoes, 
            border_style="dim white", 
            title="[bold white]Menu de Operações[/bold white]", 
            title_align="left",
            width=65,
            padding=(1, 2)
        )
        console.print(Align.center(painel_menu))
        console.print()
        
        opcao = Prompt.ask(" " * 22 + "[bold white]Opção desejada[/bold white]", choices=["0", "1", "2", "3", "4"])
        
        if opcao == '1':
            verificador.executar()
        elif opcao == '2':
            conferencia.executar()
        elif opcao == '3':
            escala.executar()
        elif opcao == '4':
            gestor_efetivo.executar()
        elif opcao == '0':
            utils.limpar_tela()
            console.print(Panel("[bold green]Sessão encerrada. Bom descanso! 🚁[/bold green]", border_style="green", expand=False))
            break

if __name__ == "__main__":
    try:
        menu_principal()
    except KeyboardInterrupt:
        console.print("\n\n[bold red]Operação cancelada (CTRL+C).[/bold red]\n")
        sys.exit(0)