import os
import time
import datetime
import calendar
import re

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

def verificar_e_propor_correcoes(escala_detalhada, mapa_efetivo, ano, mes):
    inconsistencias = [] 
    correcoes = []       
    dias = sorted(escala_detalhada.keys())
    ignorar = ['---', '???', 'PND', 'ERR']
    
    def obter_legenda_pelo_nome(nome_guerra):
        return utils.encontrar_legenda(nome_guerra, mapa_efetivo)

    for i, dia in enumerate(dias):
        dados_dia = escala_detalhada[dia]
        dia_sem = get_sem(ano, mes, dia)
        l1, l2, l3 = dados_dia[1]['legenda'], dados_dia[2]['legenda'], dados_dia[3]['legenda']
        
        # --- REGRA 1: DOBRA DE TURNO ---
        violacao = None
        turno_suspeito = None
        if l1 not in ignorar and l1 == l2: violacao = (1, 2, l1); turno_suspeito = 2
        elif l2 not in ignorar and l2 == l3: violacao = (2, 3, l2); turno_suspeito = 3

        if violacao:
            t_a, t_b, leg = violacao
            nome_mil = obter_info_militar(leg, mapa_efetivo)
            leg_ass = obter_legenda_pelo_nome(dados_dia[turno_suspeito]['assinatura_nome'])
            
            if leg_ass not in ['---', '???'] and leg_ass != leg:
                nome_correto = obter_info_militar(leg_ass, mapa_efetivo)
                inconsistencias.append(f"⚠️ Dia {dia:02d} ({dia_sem}): Militar {nome_mil} ({leg}) em turnos seguidos ({t_a}º/{t_b}º).\n      ↳ 🕵️‍♂️ {Cor.GREEN}SOLUÇÃO:{Cor.RESET} O {t_b}ºT foi assinado por {nome_correto} ({leg_ass}). Erro de digitação.")
                correcoes.append({'dia': dia, 'turno': turno_suspeito, 'nova_leg': leg_ass})
            else:
                inconsistencias.append(f"⚠️ Dia {dia:02d} ({dia_sem}): Militar {nome_mil} ({leg}) dobrou o turno ({t_a}º e {t_b}º).\n      ↳ ⚖️ {Cor.YELLOW}INFO:{Cor.RESET} Assinaturas confirmam que o militar realmente cumpriu ambos os turnos.")

        # --- REGRA 2: FOLGA PÓS-3º TURNO ---
        if i > 0:
            dia_ant = dias[i-1]
            dia_sem_ant = get_sem(ano, mes, dia_ant)
            l3_ant = escala_detalhada[dia_ant][3]['legenda']
            
            if l3_ant not in ignorar:
                turnos_hoje_violados = [t for t in [1, 2, 3] if escala_detalhada[dia][t]['legenda'] == l3_ant]
                if turnos_hoje_violados:
                    nome_mil = obter_info_militar(l3_ant, mapa_efetivo)
                    pistas = []
                    violation_real = True 

                    for t_hoje in turnos_hoje_violados:
                        leg_ass_hoje = obter_legenda_pelo_nome(escala_detalhada[dia][t_hoje]['assinatura_nome'])
                        if leg_ass_hoje not in ['---', '???'] and leg_ass_hoje != l3_ant:
                            nome_correto = obter_info_militar(leg_ass_hoje, mapa_efetivo)
                            pistas.append(f"O LRO do dia {dia:02d} ({t_hoje}ºT) foi assinado por {nome_correto} ({leg_ass_hoje}).")
                            correcoes.append({'dia': dia, 'turno': t_hoje, 'nova_leg': leg_ass_hoje})
                            violation_real = False

                    msg_auditoria = ""
                    if pistas: msg_auditoria = f"\n      ↳ 🕵️‍♂️ {Cor.GREEN}SOLUÇÃO:{Cor.RESET} " + " ".join(pistas) + " Provável erro de digitação."
                    elif violation_real: msg_auditoria = f"\n      ↳ ⚖️ {Cor.YELLOW}INFO:{Cor.RESET} Assinaturas confirmam que {nome_mil} cumpriu o serviço sem a folga regulamentar."

                    str_turnos = " e ".join([f"{t}º" for t in turnos_hoje_violados])
                    inconsistencias.append(f"🚨 Dia {dia:02d} ({dia_sem}): Militar {nome_mil} ({l3_ant}) sem folga do dia {dia_ant:02d} ({dia_sem_ant}).{msg_auditoria}")
                    
    return inconsistencias, correcoes

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
            
            if opcao_escala == '0': return 
            if opcao_escala not in ['1', '2', '3']: continue

            path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
            if os.name != 'nt' and not os.path.exists(path_ano): path_ano = Config.CAMINHO_RAIZ 
            path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
            if os.name != 'nt': path_mes = "." 
            
            if not os.path.exists(path_mes) and os.name == 'nt':
                print(f"{Cor.RED}Pasta não encontrada.{Cor.RESET}"); time.sleep(2); break

            try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
            except: break
            if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

            print(f"\n{Cor.GREY}A processar dados e auditar inconsistências... Aguarde.{Cor.RESET}\n")

            escala_detalhada = {}
            mapa_ativo = mapa_bct if opcao_escala == '2' else mapa_oea 

            for dia in range(1, qtd_dias + 1):
                dia_fmt = f"{dia:02d}"
                data_str = f"{dia_fmt}{mes}{ano_curto}"
                turnos = utils.calcular_turnos_validos(dia, mes, agora.day, mes_atual, agora.hour)

                dia_dados = {'smc': '---', 'bct': {1:'---',2:'---',3:'---'}, 'oea': {1:'---',2:'---',3:'---'},
                             'meta': {1:{'assinatura_nome':'???'}, 2:{'assinatura_nome':'???'}, 3:{'assinatura_nome':'???'}}}

                for turno in turnos:
                    arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, turno)
                    if not arquivos: continue
                    pdfs = [f for f in arquivos if f.lower().endswith('.pdf')]
                    if not pdfs:
                        if any("FALTA LRO" in f.upper() for f in arquivos):
                            dia_dados['bct'][turno] = 'PND'; dia_dados['oea'][turno] = 'PND'
                            dia_dados['meta'][turno]['assinatura_nome'] = 'PND'
                        continue

                    arquivo_alvo = [f for f in pdfs if "OK" in f.upper()][0] if [f for f in pdfs if "OK" in f.upper()] else pdfs[0]
                    info = utils.analisar_conteudo_lro(arquivo_alvo)
                    if info:
                        # O utils agora já traz o nome inteligente! Basta converter para legenda.
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

            # --- AUDITORIA E EXIBIÇÃO ---
            if opcao_escala in ['2', '3']:
                inconsistencias, correcoes = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)
                if inconsistencias:
                    print(f"\n{Cor.bg_BLUE}{Cor.WHITE} 🔍 ANÁLISE PRÉVIA DE CONSISTÊNCIA {Cor.RESET}")
                    for inc in inconsistencias: print(f"{Cor.RED}{inc}{Cor.RESET}")
                    if correcoes and utils.pedir_confirmacao(f"\n{Cor.YELLOW}>> Aplicar correções de digitação na tabela? (S/Enter p/ Sim, ESC p/ Não): {Cor.RESET}"):
                        for c in correcoes: escala_detalhada[c['dia']][c['turno']]['legenda'] = c['nova_leg']

            print("\n")
            if opcao_escala == '1':
                print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM |  SMC  {Cor.RESET}"); tracos = 19
            else:
                print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM | 1º TURNO | 2º TURNO | 3º TURNO {Cor.RESET}"); tracos = 45

            for dia in range(1, qtd_dias + 1):
                dt = datetime.date(int(ano_longo), int(mes), dia)
                sigla_sem = Config.MAPA_SEMANA[dt.weekday()]
                if opcao_escala == '1':
                    print(f" {dia:02d}  | {sigla_sem} |  {escala_detalhada[dia]['smc']:^3}  ")
                else:
                    l1, l2, l3 = escala_detalhada[dia][1]['legenda'], escala_detalhada[dia][2]['legenda'], escala_detalhada[dia][3]['legenda']
                    print(f" {dia:02d}  | {sigla_sem} |   {l1:^4}   |   {l2:^4}   |   {l3:^4}   ")

            print("-" * tracos)
            if opcao_escala in ['2', '3']:
                inc_final, _ = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)
                print("\n" + f"{Cor.bg_ORANGE}{Cor.WHITE} RESUMO DA AUDITORIA OPERACIONAL {Cor.RESET}".center(tracos + 10))
                if inc_final:
                    for inc in inc_final: print(f"{Cor.RED}{inc}{Cor.RESET}")
                else: print(f"{Cor.GREEN}✅ Escala consistente com as normas de folga.{Cor.RESET}")
                print("-" * tracos)

            if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outra especialidade? (S/Enter p/ Sim, ESC p/ Voltar): {Cor.RESET}"): return