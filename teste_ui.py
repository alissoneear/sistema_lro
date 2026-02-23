# teste_ui.py
import datetime
import utils
from config import Config
from modulos import escala

utils.limpar_tela()

# 1. Criamos dados FALSOS e perfeitos instantaneamente
escala_fake = {}
for dia in range(1, 32):
    escala_fake[dia] = {
        1: {'legenda': 'A'},
        2: {'legenda': 'B'},
        3: {'legenda': '---'} # Simulando uma pendência
    }

# 2. Chamamos a sua tabela passando os dados falsos
print("🧪 LABORATÓRIO VISUAL DA TABELA 🧪\n")

escala.imprimir_tabela(
    escala_detalhada=escala_fake, 
    qtd_dias=31, 
    opcao_escala='2', # '2' é BCT (mostra os 3 turnos)
    ano_longo='2026', 
    mes='01'
)