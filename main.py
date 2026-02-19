import time
import sys
from config import Cor
from utils import limpar_tela

# Importa os módulos
from modulos import verificador
from modulos import conferencia # Novo!
from modulos import escala

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
        print(f"  {Cor.CYAN}[1]{Cor.RESET} Verificador LRO (Validar/Assinar)")
        print(f"  {Cor.CYAN}[2]{Cor.RESET} Conferência Rápida (Auxiliar Escala)")
        print(f"  {Cor.CYAN}[3]{Cor.RESET} Escala Cumprida (Gerar Tabela/Auditoria)")
        print(f"  {Cor.RED}[0]{Cor.RESET} Sair\n")
        
        opcao = input("Opção: ")
        
        if opcao == '1':
            verificador.executar()
        elif opcao == '2':
            conferencia.executar()
        elif opcao == '3':
            escala.executar()
        elif opcao == '0':
            limpar_tela()
            print(f"{Cor.GREEN}A sair do SISTEMA LRO... Até à próxima!{Cor.RESET}\n")
            break
        else:
            print(f"\n{Cor.RED}[!] Opção inválida! Tente novamente.{Cor.RESET}")
            time.sleep(1.5)

if __name__ == "__main__":
    try:
        menu_principal()
    except KeyboardInterrupt:
        print(f"\n\n{Cor.GREEN}Sessão encerrada (CTRL+C). Até à próxima!{Cor.RESET}\n")
        sys.exit(0)