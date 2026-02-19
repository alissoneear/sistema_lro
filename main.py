import time
import sys
from config import Cor
from utils import limpar_tela

# Importa os nossos novos módulos
from modulos import verificador
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
        print(f"  {Cor.CYAN}[1]{Cor.RESET} Verificador LRO")
        print(f"  {Cor.CYAN}[2]{Cor.RESET} Escala Cumprida")
        print(f"  {Cor.RED}[0]{Cor.RESET} Sair\n")
        
        opcao = input("Opção: ")
        
        if opcao == '1':
            verificador.executar()
        elif opcao == '2':
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
        # Captura o CTRL+C em qualquer parte do sistema e sai de forma elegante
        print(f"\n\n{Cor.GREEN}Sessão encerrada (CTRL+C). Até à próxima!{Cor.RESET}\n")
        sys.exit(0)