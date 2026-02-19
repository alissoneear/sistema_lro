import os
import time
import datetime
import calendar

from config import Config, Cor
import utils
from modulos import verificador # Usaremos a função de exibir resumo que já criamos lá

def executar():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    while True:
        utils.limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        print(f"{Cor.YELLOW}=== Conferência Rápida (Auxílio de Escala) ==={Cor.RESET}")
        
        inp_mes = input(f"MÊS (Enter para {mes_atual}): ")
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        inp_ano = input(f"ANO (Enter para {ano_atual_curto}): ")
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        ano_longo = "20" + ano_curto

        path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
        if os.name != 'nt' and not os.path.exists(path_ano): path_ano = Config.CAMINHO_RAIZ 
            
        path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
        if os.name == 'nt' and not os.path.exists(path_mes):
            print(f"{Cor.RED}Pasta do mês não encontrada.{Cor.RESET}"); time.sleep(2); continue
        
        if os.name != 'nt': path_mes = "." 

        try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
        except: continue
        if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

        print(f"\n{Cor.CYAN}Iniciando conferência de {qtd_dias} dias...{Cor.RESET}")
        print(f"{Cor.GREY}Aperte S/Enter para PRÓXIMO ou ESC para SAIR.{Cor.RESET}\n")
        time.sleep(1)

        contador = 0
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
                
                # --- ESPAÇAMENTO ADICIONADO AQUI ---
                print("\n\n") 
                print(f"{Cor.bg_ORANGE}{Cor.WHITE}  ▶ CONFERÊNCIA - Arquivo: {os.path.basename(arquivo_alvo)}  {Cor.RESET}")
                
                info = utils.analisar_conteudo_lro(arquivo_alvo)
                verificador.exibir_dados_analise(info)
                
                if not utils.pedir_confirmacao(f"{Cor.YELLOW}>> Próximo turno? (S/Enter p/ Sim, ESC p/ Parar): {Cor.RESET}"):
                    return

        if contador == 0:
            print(f"{Cor.RED}Nenhum PDF encontrado para este período.{Cor.RESET}")
            time.sleep(2)
        else:
            print(f"\n{Cor.GREEN}Conferência finalizada! {contador} turnos visualizados.{Cor.RESET}")
            time.sleep(2)
        
        if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outro mês? (S/Enter p/ Sim, ESC p/ Voltar): {Cor.RESET}"): 
            break