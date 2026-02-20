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
    """Função auxiliar para imprimir a tabela da escala formatada."""
    print("\n")
    if opcao_escala == '1':
        print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM |  SMC  {Cor.RESET}")
        tracos = 19
    else:
        print(f"{Cor.bg_BLUE}{Cor.WHITE} DIA | SEM | 1º TURNO | 2º TURNO | 3º TURNO {Cor.RESET}")
        tracos = 45

    for dia in range(1, qtd_dias + 1):
        dt = datetime.date(int(ano_longo), int(mes), dia)
        sigla_sem = Config.MAPA_SEMANA[dt.weekday()]
        if opcao_escala == '1':
            print(f" {dia:02d}  | {sigla_sem} |  {escala_detalhada[dia]['smc']:^3}  ")
        else:
            l1 = escala_detalhada[dia][1]['legenda']
            l2 = escala_detalhada[dia][2]['legenda']
            l3 = escala_detalhada[dia][3]['legenda']
            print(f" {dia:02d}  | {sigla_sem} |   {l1:^4}   |   {l2:^4}   |   {l3:^4}   ")

    print("-" * tracos)
    return tracos

def realizar_auditoria_manual(escala_detalhada, mes, ano_curto, path_mes, opcao_escala, mapa_ativo, alertas_suspeitos=None):
    """Procura falhas na extração OU alertas da auditoria e abre o PDF exibindo os motivos."""
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
        
    print(f"\n{Cor.YELLOW}⚠️ Foram detetados {len(pendentes)} turnos pendentes de revisão.{Cor.RESET}")
    if not utils.pedir_confirmacao(f"{Cor.CYAN}>> Deseja realizar a AUDITORIA MANUAL agora? (S/Enter p/ Sim, ESC p/ Pular): {Cor.RESET}"):
        return False
        
    validas = [dados['legenda'] for dados in mapa_ativo.values()]
    validas_str = ", ".join(validas)
    modificado = False
    
    for dia, t, leg_atual, motivos in pendentes:
        dia_fmt = f"{dia:02d}"
        data_str = f"{dia_fmt}{mes}{ano_curto}"
        
        # Cabeçalho da auditoria focado no turno e nos motivos
        if opcao_escala == '1':
            print(f"\n{Cor.bg_BLUE}{Cor.WHITE} Auditando: Dia {dia_fmt} - SMC {Cor.RESET}")
            arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, 1) 
        else:
            print(f"\n{Cor.bg_BLUE}{Cor.WHITE} Auditando: Dia {dia_fmt} - {t}º Turno {Cor.RESET}")
            arquivos = utils.buscar_arquivos_flexivel(path_mes, data_str, t)
            
        for m in motivos:
            print(f"{Cor.YELLOW} ↳ Motivo:{Cor.RESET} {m}")
            
        pdfs = [f for f in arquivos if f.lower().endswith('.pdf')]
        arquivo_alvo = None
        if pdfs:
            ok_files = [f for f in pdfs if "OK" in f.upper()]
            arquivo_alvo = ok_files[0] if ok_files else pdfs[0]
            
        if arquivo_alvo:
            print(f"{Cor.GREY}A abrir o documento: {os.path.basename(arquivo_alvo)}{Cor.RESET}")
            utils.abrir_arquivo(arquivo_alvo)
        else:
            print(f"{Cor.RED}[!] Nenhum PDF encontrado para este turno.{Cor.RESET}")
            
        nova_leg = input(f"Digite a legenda correta ({validas_str}) ou Enter para manter [{leg_atual}]: ").strip().upper()
        
        if nova_leg:
            if nova_leg in validas:
                if opcao_escala == '1':
                    escala_detalhada[dia]['smc'] = nova_leg
                else:
                    escala_detalhada[dia][t]['legenda'] = nova_leg
                modificado = True
                print(f"{Cor.GREEN}✅ Legenda atualizada para: {nova_leg}{Cor.RESET}")
            else:
                print(f"{Cor.RED}⚠️ Legenda '{nova_leg}' inválida. A manter '{leg_atual}'.{Cor.RESET}")
        else:
            print(f"{Cor.GREY}Mantido: {leg_atual}{Cor.RESET}")
            
    return modificado

def verificar_e_propor_correcoes(escala_detalhada, mapa_efetivo, ano, mes):
    inconsistencias = [] 
    correcoes = []       
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
        turno_suspeito = None
        if l1 not in ignorar and l1 == l2: violacao = (1, 2, l1); turno_suspeito = 2
        elif l2 not in ignorar and l2 == l3: violacao = (2, 3, l2); turno_suspeito = 3

        if violacao:
            t_a, t_b, leg = violacao
            nome_mil = obter_info_militar(leg, mapa_efetivo)
            leg_ass = obter_legenda_pelo_nome(dados_dia[turno_suspeito]['assinatura_nome'])
            
            # Motivos para a revisão visual
            motivo_dobra = f"Possível dobra de turno detetada ({t_a}º e {t_b}º) do militar {nome_mil} ({leg})."
            add_alerta(dia, t_a, motivo_dobra)
            add_alerta(dia, t_b, motivo_dobra)
            
            if leg_ass not in ['---', '???'] and leg_ass != leg:
                nome_correto = obter_info_militar(leg_ass, mapa_efetivo)
                inconsistencias.append(f"⚠️ Dia {dia:02d} ({dia_sem}): Militar {nome_mil} ({leg}) em turnos seguidos ({t_a}º/{t_b}º).\n      ↳ 🕵️‍♂️ {Cor.GREEN}SOLUÇÃO:{Cor.RESET} O {t_b}ºT foi assinado por {nome_correto} ({leg_ass}). Erro de digitação.")
                correcoes.append({'dia': dia, 'turno': turno_suspeito, 'nova_leg': leg_ass})
            else:
                inconsistencias.append(f"⚠️ Dia {dia:02d} ({dia_sem}): Militar {nome_mil} ({leg}) dobrou o turno ({t_a}º e {t_b}º).\n      ↳ ⚖️ {Cor.YELLOW}INFO:{Cor.RESET} Assinaturas confirmam que o militar cumpriu ambos os turnos.")

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
                        # Adiciona o alerta para quem está a falhar a folga hoje e quem tirou o T3 ontem
                        motivo_folga_hoje = f"Falta de folga regulamentar. O militar {nome_mil} ({l3_ant}) estava escalado no 3º Turno do dia {dia_ant:02d}."
                        motivo_folga_ontem = f"Militar {nome_mil} ({l3_ant}) escalado aqui, mas aparece sem folga no {t_hoje}º Turno do dia seguinte ({dia:02d})."
                        
                        add_alerta(dia, t_hoje, motivo_folga_hoje)
                        add_alerta(dia_ant, 3, motivo_folga_ontem)

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
        for t in [1, 2, 3]:
            leg = turnos_dia[t]['legenda']
            if leg not in ['---', 'PND', 'ERR', '???']:
                dados_pdf[f"d{dia}_t{t}"] = leg
                horas_militares[leg] = horas_militares.get(leg, 0) + minutos_por_turno[t]
            else:
                dados_pdf[f"d{dia}_t{t}"] = ""
                
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
        
        if utils.pedir_confirmacao(f"\n{Cor.YELLOW}>> Deseja ABRIR o PDF gerado agora? (S/Enter p/ Sim, ESC p/ Não): {Cor.RESET}"):
            utils.abrir_arquivo(caminho_saida)
    except Exception as e:
        print(f"{Cor.RED}[!] Erro ao gravar o ficheiro PDF: {e}{Cor.RESET}")

def executar():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

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

        mapa_smc, mapa_bct, mapa_oea = DadosEfetivo.mapear_efetivo()

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
            path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
            
            if not os.path.exists(path_mes):
                print(f"{Cor.RED}Pasta não encontrada.{Cor.RESET}")
                time.sleep(2)
                break

            try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
            except: break
            if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

            print(f"\n{Cor.GREY}A processar dados e auditar inconsistências... Aguarde.{Cor.RESET}\n")

            escala_detalhada = {}
            mapa_ativo = mapa_bct if opcao_escala == '2' else mapa_oea 

            # Extração de Dados
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

            # --- 1. ANÁLISE PRÉVIA E PRIMEIRA EXIBIÇÃO ---
            if opcao_escala in ['2', '3']:
                inconsistencias, correcoes, _ = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)
                if inconsistencias:
                    print(f"\n{Cor.bg_BLUE}{Cor.WHITE} 🔍 ANÁLISE PRÉVIA DE CONSISTÊNCIA {Cor.RESET}")
                    for inc in inconsistencias: print(f"{Cor.RED}{inc}{Cor.RESET}")
                    if correcoes and utils.pedir_confirmacao(f"\n{Cor.YELLOW}>> Aplicar correções de digitação na tabela? (S/Enter p/ Sim, ESC p/ Não): {Cor.RESET}"):
                        for c in correcoes: escala_detalhada[c['dia']][c['turno']]['legenda'] = c['nova_leg']

            tracos = imprimir_tabela(escala_detalhada, qtd_dias, opcao_escala, ano_longo, mes)

            # --- 2. LOOP DE AUDITORIA MANUAL INFINITO ---
            while True:
                alertas_ativos = {}
                if opcao_escala in ['2', '3']:
                    _, _, alertas_ativos = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)

                teve_correcao_manual = realizar_auditoria_manual(escala_detalhada, mes, ano_curto, path_mes, opcao_escala, mapa_ativo, alertas_ativos)
                
                if teve_correcao_manual:
                    utils.limpar_tela()
                    print(f"{Cor.ORANGE}=== SISTEMA LRO - Escala Cumprida ({mes}/{ano_curto}) ==={Cor.RESET}")
                    print(f"\n{Cor.GREEN}✅ Tabela atualizada após Auditoria Manual:{Cor.RESET}")
                    tracos = imprimir_tabela(escala_detalhada, qtd_dias, opcao_escala, ano_longo, mes)
                else:
                    break

            # --- 3. RESUMO FINAL E GERAÇÃO DO PDF ---
            if opcao_escala in ['2', '3']:
                inc_final, _, _ = verificar_e_propor_correcoes(escala_detalhada, mapa_ativo, ano_longo, mes)
                print("\n" + f"{Cor.bg_ORANGE}{Cor.WHITE} RESUMO DA AUDITORIA OPERACIONAL {Cor.RESET}".center(tracos + 10))
                if inc_final:
                    for inc in inc_final: print(f"{Cor.RED}{inc}{Cor.RESET}")
                else: print(f"{Cor.GREEN}✅ Escala consistente com as normas de folga.{Cor.RESET}")
                print("-" * tracos)
                
                if utils.pedir_confirmacao(f"\n{Cor.CYAN}>> Deseja GERAR O PDF OFICIAL desta Escala Cumprida? (S/Enter p/ Sim, ESC p/ Pular): {Cor.RESET}"):
                    gerar_pdf_escala(escala_detalhada, mapa_ativo, opcao_escala, mes, ano_longo)

            if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outra especialidade? (S/Enter p/ Sim, ESC p/ Voltar): {Cor.RESET}"): return