import os
import time
import datetime
import calendar

# Importa as configurações e utilitários da pasta principal
from config import Config, Cor
import utils

def corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado):
    if not lista_ano_errado: return
    if utils.pedir_confirmacao(f"{Cor.DARK_RED}Corrigir {len(lista_ano_errado)} anos errados? (S/Enter p/ Sim, ESC p/ Pular): {Cor.RESET}"):
        for arq in lista_ano_errado:
            utils.abrir_arquivo(arq)
            if utils.pedir_confirmacao(f"   Renomear {os.path.basename(arq)}? (S/Enter p/ Sim, ESC p/ Não): "):
                try: 
                    os.rename(arq, arq.replace(ano_errado, ano_curto))
                    print(f"{Cor.GREEN}   Renomeado com sucesso.{Cor.RESET}")
                except Exception as e: print(f"{Cor.RED}   Erro ao renomear: {e}{Cor.RESET}")

def exibir_dados_analise(info):
    if not info: return
    print("-" * 60)
    print(f"📅 {Cor.GREEN}{info['cabecalho']}{Cor.RESET}")
    print(f"👤 RESPONSÁVEL: {Cor.CYAN}{info['responsavel']}{Cor.RESET}")
    if info['assinatura']: print(f"🔏 ASSINATURA: {Cor.GREEN}OK (Certificado Digital Detectado) ✅{Cor.RESET}")
    else: print(f"🔏 ASSINATURA: {Cor.RED}NÃO DETECTADA NA ESTRUTURA ❌{Cor.RESET}")
    print("-" * 60)
    print(f"⬅️ Recebeu: {info['recebeu']}")
    print(f"➡️ Passou:  {info['passou']}")
    print(f"{Cor.WHITE}👥 Equipe:{Cor.RESET}")
    print(f"   • SMC: {info['equipe']['smc']}")
    print(f"   • BCT: {info['equipe']['bct']}")
    print(f"   • OEA: {info['equipe']['oea']}")
    print("-" * 60)

def processo_verificacao_visual(lista_pendentes):
    if not lista_pendentes: return
    print(f"{Cor.CYAN}Verificar {len(lista_pendentes)} livros pendentes?{Cor.RESET}")
    if not utils.pedir_confirmacao("Iniciar UM POR UM? (S/Enter p/ Sim, ESC p/ Pular): "): return

    for cnt, item in enumerate(lista_pendentes, start=1):
        caminho, data_str, turno = item['path'], item['data'], item['turno']
        nome_atual = os.path.basename(caminho)
        print(f"\n{Cor.YELLOW}--- ({cnt}/{len(lista_pendentes)}) ---{Cor.RESET}")
        print(f"{Cor.CYAN}Arquivo: {nome_atual}{Cor.RESET}")
        utils.abrir_arquivo(caminho)
        print(f"{Cor.GREY}A analisar estrutura e texto...{Cor.RESET}")
        info = utils.analisar_conteudo_lro(caminho)
        exibir_dados_analise(info)
        
        if utils.pedir_confirmacao(">> Confirmar e assinar? (S/Enter p/ Sim, ESC p/ Não): "):
            dir_arq = os.path.dirname(caminho)
            extensao = os.path.splitext(caminho)[1]
            novo_nome_base = f"{data_str}_{turno}TURNO OK{extensao}"
            novo_caminho = os.path.join(dir_arq, novo_nome_base)
            while True:
                try:
                    os.rename(caminho, novo_caminho)
                    print(f"{Cor.GREEN}   [V] Validado e Padronizado: {novo_nome_base}{Cor.RESET}")
                    break
                except PermissionError: 
                    if not utils.pedir_confirmacao(f"{Cor.RED}   [!] Feche o ficheiro PDF! S/Enter para tentar novamente, ESC para pular.{Cor.RESET}"):
                        break
                except Exception as e: print(f"{Cor.RED}Erro: {e}{Cor.RESET}"); break
        else: print(f"{Cor.GREY}   [-] Mantido.{Cor.RESET}")

def executar():
    if os.name == 'nt' and not os.path.exists(Config.CAMINHO_RAIZ):
        input(f"{Cor.RED}[ERRO CRÍTICO] Caminho {Config.CAMINHO_RAIZ} não encontrado.{Cor.RESET}")
        return

    while True:
        utils.limpar_tela()
        agora = datetime.datetime.now()
        mes_atual, ano_atual_curto = agora.strftime("%m"), agora.strftime("%y")

        print(f"{Cor.CYAN}=== Verificador LRO ==={Cor.RESET}")
        if os.name == 'nt': print(f"{Cor.GREY}Conectado: {Config.CAMINHO_RAIZ}{Cor.RESET}")
        
        inp_mes = input(f"MÊS (Enter para {mes_atual}): ")
        mes = inp_mes.zfill(2) if inp_mes else mes_atual
        inp_ano = input(f"ANO (Enter para {ano_atual_curto}): ")
        ano_curto = inp_ano if inp_ano else ano_atual_curto
        
        ano_longo = "20" + ano_curto
        try: ano_errado = str(int(ano_curto) - 1).zfill(2)
        except ValueError: ano_errado = "00"

        path_ano = os.path.join(Config.CAMINHO_RAIZ, f"LRO {ano_longo}")
        if os.name != 'nt' and not os.path.exists(path_ano): path_ano = Config.CAMINHO_RAIZ 
            
        if os.name == 'nt' and not os.path.exists(path_ano):
            print(f"{Cor.RED}Pasta LRO {ano_longo} não existe.{Cor.RESET}"); time.sleep(2); continue
            
        path_mes = os.path.join(path_ano, Config.MAPA_PASTAS.get(mes, "X"))
        if os.name == 'nt' and not os.path.exists(path_mes):
            print(f"{Cor.RED}Pasta do mês não encontrada.{Cor.RESET}"); time.sleep(2); continue
        
        if os.name != 'nt': path_mes = "." 

        try: qtd_dias = calendar.monthrange(int(ano_longo), int(mes))[1]
        except Exception: continue
            
        if mes == mes_atual and ano_curto == ano_atual_curto: qtd_dias = agora.day

        print(f"\n{Cor.YELLOW}>> Analisando {Config.MAPA_PASTAS.get(mes)}...{Cor.RESET}")
        print("-" * 60)

        problemas = 0
        relatorio, lista_pendentes, lista_para_criar, lista_ano_errado = [], [], [], []

        for dia in range(1, qtd_dias + 1):
            dia_fmt = f"{dia:02d}"
            data_str = f"{dia_fmt}{mes}{ano_curto}"
            data_str_err = f"{dia_fmt}{mes}{ano_errado}"
            turnos = utils.calcular_turnos_validos(dia, mes, agora.day, mes_atual, agora.hour)

            for turno in turnos:
                arquivos_turno = utils.buscar_arquivos_flexivel(path_mes, data_str, turno)
                arquivos_errados = utils.buscar_arquivos_flexivel(path_mes, data_str_err, turno)
                tem_ok = [f for f in arquivos_turno if "OK" in f.upper()]
                tem_falta_ass = [f for f in arquivos_turno if "FALTA ASS" in f.upper()]
                tem_falta_lro = [f for f in arquivos_turno if "FALTA LRO" in f.upper()]
                tem_novo = [f for f in arquivos_turno if "OK" not in f.upper() and "FALTA" not in f.upper()]

                if tem_ok: continue
                elif tem_falta_ass:
                    print(f"{Cor.YELLOW}[!] FALTA ASSINATURA: {os.path.basename(tem_falta_ass[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_ass[0])); problemas += 1
                elif tem_novo and tem_falta_lro:
                    print(f"{Cor.DARK_YELLOW}[!] LRO FEITO (TXT PRESENTE): {os.path.basename(tem_novo[0])}{Cor.RESET}")
                    lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno}); problemas += 1
                elif tem_falta_lro:
                    print(f"{Cor.MAGENTA}[!] NÃO CONFECIONADO: {os.path.basename(tem_falta_lro[0])}{Cor.RESET}")
                    relatorio.append(os.path.basename(tem_falta_lro[0])); problemas += 1
                elif tem_novo:
                    n = os.path.basename(tem_novo[0])
                    msg_padrao = "NOME FORA DO PADRÃO" if ("-" in n or " " in n.replace("TURNO", "")) else "PENDENTE DE VERIFICAÇÃO"
                    print(f"{Cor.CYAN}[?] {msg_padrao}: {n}{Cor.RESET}")
                    lista_pendentes.append({'path': tem_novo[0], 'data': data_str, 'turno': turno}); problemas += 1
                elif arquivos_errados:
                    print(f"{Cor.DARK_RED}[!] ANO INCORRETO: {os.path.basename(arquivos_errados[0])}{Cor.RESET}")
                    lista_ano_errado.append(arquivos_errados[0]); problemas += 1
                else:
                    print(f"{Cor.RED}[X] INEXISTENTE: {data_str} - {turno}º Turno{Cor.RESET}")
                    relatorio.append(f"Dia {dia_fmt} - {turno}º Turno: NÃO ENCONTRADO"); problemas += 1
                    lista_para_criar.append({"str": data_str, "turno": turno})

        print("-" * 60)
        if problemas == 0: print(f"{Cor.GREEN}Tudo em dia!{Cor.RESET}\n")

        corrigir_anos_errados(lista_ano_errado, ano_curto, ano_errado)

        if problemas > 0 and relatorio:
            print(f"\n{Cor.bg_BLUE}--- COPIAR E COBRAR ---{Cor.RESET}")
            for l in relatorio: print(l)
            print("-" * 30 + "\n")

        if lista_para_criar:
            if utils.pedir_confirmacao(f"{Cor.YELLOW}Criar ficheiros FALTA LRO? (S/Enter p/ Sim, ESC p/ Não): {Cor.RESET}"):
                for i in lista_para_criar:
                    n = f"{i['str']}_{i['turno']}TURNO FALTA LRO ().txt"
                    try: 
                        with open(os.path.join(path_mes, n), 'w') as f: f.write("Falta")
                        print(f"{Cor.GREEN}   [V] Criado: {n}{Cor.RESET}")
                    except Exception as e: print(f"{Cor.RED}   Erro ao criar {n}: {e}{Cor.RESET}")

        processo_verificacao_visual(lista_pendentes)

        if not utils.pedir_confirmacao(f"\n{Cor.YELLOW}Verificar outro mês? (S/Enter p/ Sim, ESC p/ Voltar ao menu): {Cor.RESET}"): 
            break