import os
import time
import datetime
import calendar

from config import Config, Cor, DadosEfetivo
import utils

def executar():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    mapa_smc, mapa_bct, mapa_oea = DadosEfetivo.mapear_efetivo()

    while True:
        utils.limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        print(f"{Cor.ORANGE}=== SISTEMA LRO - Escala Cumprida ==={Cor.RESET}")
        
        inp_mes = input(f"MÊS (Enter para {mes_atual}): ")
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        inp_ano = input(f"ANO (Enter para {ano_atual_curto}): ")
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        ano_longo = "20" + ano_curto

        while True:
            utils.limpar_tela()
            print(f"{Cor.ORANGE}=== SISTEMA LRO - Escala Cumprida ({mes}/{ano_curto}) ==={Cor.RESET}")
            print(f"\n{Cor.CYAN}Qual escala deseja gerar?{Cor.RESET}")
            print("  [1] SMC")
            print("  [2] BCT")
            print("  [3] OEA")
            print("  [0] Voltar ao Menu")
            opcao_escala = input("\nOpção: ")
            
            if opcao_escala == '0':
                return 
                
            if opcao_escala not in ['1', '2', '3']:
                print(f"{Cor.RED}Opção inválida! Tente novamente.{Cor.RESET}")
                time.sleep(1.5)
                continue

            path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
            if os.name != 'nt' and not os.path.exists(path_ano): path_ano = Config.CAMINHO_RAIZ 
                
            path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
            if os.name != 'nt': path_mes = "." 

            if not os.path.exists(path_mes) and os.name == 'nt':
                print(f"{Cor.RED}Pasta do mês não encontrada ({path_mes}).{Cor.RESET}"); time.sleep(2); break

            try: 
                qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
            except Exception: 
                break
                
            if mes == mes_atual and ano_curto == ano_atual_curto: 
                qtd_dias = agora.day

            print(f"\n{Cor.GREY}Extraindo dados de forma invisível... Aguarde.{Cor.RESET}\n")

            if opcao_escala == '1':
                print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM |  SMC  {Cor.RESET}")
                tracos_separador = 19
            elif opcao_escala in ['2', '3']:
                print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM | 1º TURNO | 2º TURNO | 3º TURNO {Cor.RESET}")
                tracos_separador = 45

            for dia in range(1, qtd_dias + 1):
                dia_fmt = f"{dia:02d}"
                data_str = f"{dia_fmt}{mes}{ano_curto}"
                
                data_dt = datetime.date(int(ano_longo), int(mes), dia)
                sigla_sem = Config.MAPA_SEMANA[data_dt.weekday()]
                
                turnos = utils.calcular_turnos_validos(dia, mes, agora.day, mes_atual, agora.hour)

                dia_dados = {
                    'smc': '---',
                    'bct': {1: '---', 2: '---', 3: '---'},
                    'oea': {1: '---', 2: '---', 3: '---'}
                }

                for turno in turnos:
                    arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, turno)
                    if not arquivos: continue

                    arquivos_ok = [f for f in arquivos if "OK" in f.upper()]
                    arquivo_alvo = arquivos_ok[0] if arquivos_ok else arquivos[0]
                    
                    if "FALTA LRO" in arquivo_alvo.upper() and arquivo_alvo.endswith('.txt'):
                        dia_dados['bct'][turno] = 'PND'
                        dia_dados['oea'][turno] = 'PND'
                        continue

                    info = utils.analisar_conteudo_lro(arquivo_alvo)
                    if info:
                        # 1. Tenta a Busca Normal (Expressões Regulares)
                        leg_smc = utils.encontrar_legenda(info['equipe']['smc'], mapa_smc)
                        leg_bct = utils.encontrar_legenda(info['equipe']['bct'], mapa_bct)
                        leg_oea = utils.encontrar_legenda(info['equipe']['oea'], mapa_oea)
                        
                        # 2. SE FALHAR, Ativa a Busca Agressiva (Fallback) no documento todo!
                        if leg_smc in ['---', '???']: leg_smc = utils.encontrar_legenda_fallback(info['texto_completo'], mapa_smc)
                        if leg_bct in ['---', '???']: leg_bct = utils.encontrar_legenda_fallback(info['texto_completo'], mapa_bct)
                        if leg_oea in ['---', '???']: leg_oea = utils.encontrar_legenda_fallback(info['texto_completo'], mapa_oea)

                        # Fixa o SMC no primeiro turno que encontrar válido
                        if dia_dados['smc'] == '---' and leg_smc not in ['---', '???']:
                            dia_dados['smc'] = leg_smc
                        elif dia_dados['smc'] == '---':
                            dia_dados['smc'] = leg_smc
                            
                        dia_dados['bct'][turno] = leg_bct
                        dia_dados['oea'][turno] = leg_oea
                    else:
                        dia_dados['bct'][turno] = 'ERR'
                        dia_dados['oea'][turno] = 'ERR'

                if opcao_escala == '1':
                    smc = dia_dados['smc']
                    print(f" {dia_fmt}  | {sigla_sem} |  {smc:^3}  ")
                elif opcao_escala == '2':
                    b1, b2, b3 = dia_dados['bct'][1], dia_dados['bct'][2], dia_dados['bct'][3]
                    print(f" {dia_fmt}  | {sigla_sem} |   {b1:^4}   |   {b2:^4}   |   {b3:^4}   ")
                elif opcao_escala == '3':
                    o1, o2, o3 = dia_dados['oea'][1], dia_dados['oea'][2], dia_dados['oea'][3]
                    print(f" {dia_fmt}  | {sigla_sem} |   {o1:^4}   |   {o2:^4}   |   {o3:^4}   ")

            print("-" * tracos_separador)
            
            if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outra especialidade deste mês? (S/Enter p/ Sim, ESC p/ Voltar ao menu): {Cor.RESET}"):
                return