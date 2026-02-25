import os
import sys
import json

class Config:
    # Verifica se o programa está a rodar como um Executável compilado pelo PyInstaller
    if getattr(sys, 'frozen', False):
        CAMINHO_RAIZ = r"R:\DO\COI\ARCC-CW\14 - LRO" # Produção (Lá no trabalho)
    else:
        # Se for script (.py), verifica o sistema operativo para o modo DEV
        if os.name == 'nt':
            CAMINHO_RAIZ = r"E:\dev\sistema_lro\ARQUIVOS" # Windows-Dev
        else:
            CAMINHO_RAIZ = "/Users/alissonlourenco/Dev/sistema_lro/ARQUIVOS" # Mac-Dev
    
    MAPA_PASTAS = {
        "01": "1 - JANEIRO",   "02": "2 - FEVEREIRO", "03": "3 - MARÇO", "04": "4 - ABRIL",
        "05": "5 - MAIO",      "06": "6 - JUNHO",     "07": "7 - JULHO", "08": "8 - AGOSTO",
        "09": "9 - SETEMBRO",  "10": "10 - OUTUBRO",  "11": "11 - NOVEMBRO", "12": "12 - DEZEMBRO"
    }
    
    MAPA_SEMANA = {0: 'SEG', 1: 'TER', 2: 'QUA', 3: 'QUI', 4: 'SEX', 5: 'SÁB', 6: 'DOM'}

class Cor:
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    MAGENTA = '\033[95m'
    DARK_YELLOW = '\033[33m'
    DARK_RED = '\033[31m'
    WHITE = '\033[97m'
    GREY = '\033[90m'
    RESET = '\033[0m'
    ORANGE = '\033[38;5;208m'
    bg_BLUE = '\033[44m'
    bg_ORANGE = '\033[48;5;208m'
    bg_RED = '\033[41m'

class DadosEfetivo:
    HISTORICO = {
        "1025": {
            "bct": [
                '1S BCT DIOGO FAVARO WUNDERLICH - 1S FAVARO - 4202317',            
                '1S BCT RÉGIS FERRARI - 1S RÉGIS - 4378717',                       
                '1S BCT EVANDRO MACHADO BITTENCOURT - 1S BITTENCOURT - 6087841',   
                '1S BCT MARCELO ZANOTTO GONÇALVES - 1S ZANOTTO - 6100775',         
                '2S BCT GIOVANNI COCCO ALVES - 2S GIOVANNI - 6158080',             
                '2S BCT YURI DE PAIVA WERNECK - 2S WERNECK - 6301371',             
                '2S BCT YURI WIES TRAUER - 2S TRAUER - 6378781',                   
                '2S BCT EDSON DE SOUZA NUNES - 2S EDSON - 6750451',                
                'SO BCT MARCELO DE SOUZA PEREIRA - SO SOUZA - 3988554',            
                '1S BCT GUSTAVO FRACAO LAGO - 1S GUSTAVO - 4378482',               
                '1S BCT CLEON FRAGA DOS SANTOS - 1S CLEON - 4378660',              
                '1S BCT CRISTIANO PAZ PRATES - 1S PRATES - 4378768',               
                '2S BCT LIS CECI LYRA FONTES - 2S LIS - 6301118',                  
                '2S BCT EDUARDO OLIVEIRA DE CASTRO - 2S DE CASTRO - 6576044',      
            ],
            "oea": [
                '1S BCO LUIZ FILIPE TERBECK - 1S TERBECK - 6032338',
                '1S BCO RUI ANTONIO DOS SANTOS JUNIOR - 1S RUI - 6032311',
                '2S BCO SAULO JULIO SANTOS GONÇALVES - 2S SAULO - 6134262',
                '2S BCO KEVIN SOUZA DE OLIVEIRA - 2S KEVIN - 6157912',
                '2S BCO ALAN SITORSKI - 2S SITORSKI - 6156568',
                '2S BCO LUCIANO FERNANDES DOS SANTOS - 2S FERNANDES - 6156878',
                'SO BCO FRANCISCO TORRES SOARES - SO SOARES - 3502201',
                '2S BCO ELIAS EDVANEIDSON DE ARAUJO CARDOSO - 2S ELIAS - 6328512',
                '2S BCO ALISSON LOURENÇO DA SILVA - 2S LOURENÇO - 6780245'
            ]
        },
        "1125": {
            "bct": [
                '1S BCT DIOGO FAVARO WUNDERLICH - 1S FAVARO - 4202317',            
                '1S BCT RÉGIS FERRARI - 1S RÉGIS - 4378717',                       
                '1S BCT EVANDRO MACHADO BITTENCOURT - 1S BITTENCOURT - 6087841',   
                '1S BCT MARCELO ZANOTTO GONÇALVES - 1S ZANOTTO - 6100775',         
                '1S BCT GIOVANNI COCCO ALVES - 1S GIOVANNI - 6158080',             
                '2S BCT YURI DE PAIVA WERNECK - 2S WERNECK - 6301371',             
                '2S BCT YURI WIES TRAUER - 2S TRAUER - 6378781',                   
                '2S BCT EDSON DE SOUZA NUNES - 2S EDSON - 6750451',                
                'SO BCT MARCELO DE SOUZA PEREIRA - SO SOUZA - 3988554',            
                '1S BCT GUSTAVO FRACAO LAGO - 1S GUSTAVO - 4378482',               
                '1S BCT CLEON FRAGA DOS SANTOS - 1S CLEON - 4378660',              
                '1S BCT CRISTIANO PAZ PRATES - 1S PRATES - 4378768',               
                '2S BCT LIS CECI LYRA FONTES - 2S LIS - 6301118',                  
                '2S BCT EDUARDO OLIVEIRA DE CASTRO - 2S DE CASTRO - 6576044'       
            ],
            "oea": [
                '1S BCO LUIZ FILIPE TERBECK - 1S TERBECK - 6032338',
                '1S BCO RUI ANTONIO DOS SANTOS JUNIOR - 1S RUI - 6032311',
                '2S BCO SAULO JULIO SANTOS GONÇALVES - 2S SAULO - 6134262',
                '2S BCO KEVIN SOUZA DE OLIVEIRA - 2S KEVIN - 6157912',
                '2S BCO ALAN SITORSKI - 2S SITORSKI - 6156568',
                '2S BCO LUCIANO FERNANDES DOS SANTOS - 2S FERNANDES - 6156878',
                'SO BCO FRANCISCO TORRES SOARES - SO SOARES - 3502201',
                '2S BCO ELIAS EDVANEIDSON DE ARAUJO CARDOSO - 2S ELIAS - 6328512',
                '2S BCO ALISSON LOURENÇO DA SILVA - 2S LOURENÇO - 6780245'
            ]
        },
        "1225": {
            "bct": [
                '1S BCT DIOGO FAVARO WUNDERLICH - 1S FAVARO - 4202317',            
                '1S BCT RÉGIS FERRARI - 1S RÉGIS - 4378717',                       
                '1S BCT CLEON FRAGA DOS SANTOS - 1S CLEON - 4378660',              
                '1S BCT MARCELO ZANOTTO GONÇALVES - 1S ZANOTTO - 6100775',         
                '1S BCT GIOVANNI COCCO ALVES - 1S GIOVANNI - 6158080',             
                '2S BCT YURI DE PAIVA WERNECK - 2S WERNECK - 6301371',             
                '2S BCT YURI WIES TRAUER - 2S TRAUER - 6378781',                   
                '2S BCT EDSON DE SOUZA NUNES - 2S EDSON - 6750451',                
                'SO BCT MARCELO DE SOUZA PEREIRA - SO SOUZA - 3988554',            
                '1S BCT GUSTAVO FRACAO LAGO - 1S GUSTAVO - 4378482',               
                '1S BCT EVANDRO MACHADO BITTENCOURT - 1S BITTENCOURT - 6087841',   
                '1S BCT CRISTIANO PAZ PRATES - 1S PRATES - 4378768',               
                '2S BCT LIS CECI LYRA FONTES - 2S LIS - 6301118',                  
                '2S BCT EDUARDO OLIVEIRA DE CASTRO - 2S DE CASTRO - 6576044'       
            ],
            "oea": [
                '1S BCO LUIZ FILIPE TERBECK - 1S TERBECK - 6032338',
                '1S BCO RUI ANTONIO DOS SANTOS JUNIOR - 1S RUI - 6032311',
                '2S BCO SAULO JULIO SANTOS GONÇALVES - 2S SAULO - 6134262',
                '2S BCO KEVIN SOUZA DE OLIVEIRA - 2S KEVIN - 6157912',
                '2S BCO ALAN SITORSKI - 2S SITORSKI - 6156568',
                '2S BCO LUCIANO FERNANDES DOS SANTOS - 2S FERNANDES - 6156878',
                'SO BCO FRANCISCO TORRES SOARES - SO SOARES - 3502201',
                '2S BCO ELIAS EDVANEIDSON DE ARAUJO CARDOSO - 2S ELIAS - 6328512',
                '2S BCO ALISSON LOURENÇO DA SILVA - 2S LOURENÇO - 6780245'
            ]
        },
        "0126": {
            "bct": [
                '1S BCT DIOGO FAVARO WUNDERLICH - 1S FAVARO - 4202317',            
                '1S BCT RÉGIS FERRARI - 1S RÉGIS - 4378717',                       
                '1S BCT CLEON FRAGA DOS SANTOS - 1S CLEON - 4378660',              
                '1S BCT MARCELO ZANOTTO GONÇALVES - 1S ZANOTTO - 6100775',         
                '1S BCT GIOVANNI COCCO ALVES - 1S GIOVANNI - 6158080',             
                '2S BCT YURI DE PAIVA WERNECK - 2S WERNECK - 6301371',             
                '2S BCT YURI WIES TRAUER - 2S TRAUER - 6378781',                   
                '2S BCT EDSON DE SOUZA NUNES - 2S EDSON - 6750451',                
                'SO BCT MARCELO DE SOUZA PEREIRA - SO SOUZA - 3988554',            
                '1S BCT GUSTAVO FRACAO LAGO - 1S GUSTAVO - 4378482',               
                '1S BCT EVANDRO MACHADO BITTENCOURT - 1S BITTENCOURT - 6087841',   
                '1S BCT CRISTIANO PAZ PRATES - 1S PRATES - 4378768',               
                '2S BCT LIS CECI LYRA FONTES - 2S LIS - 6301118',                  
                '2S BCT EDUARDO OLIVEIRA DE CASTRO - 2S DE CASTRO - 6576044'       
            ],
            "oea": [
                '1S BCO LUIZ FILIPE TERBECK - 1S TERBECK - 6032338',
                '1S BCO RUI ANTONIO DOS SANTOS JUNIOR - 1S RUI - 6032311',
                '2S BCO SAULO JULIO SANTOS GONÇALVES - 2S SAULO - 6134262',
                '2S BCO KEVIN SOUZA DE OLIVEIRA - 2S KEVIN - 6157912',
                '2S BCO ALAN SITORSKI - 2S SITORSKI - 6156568',
                '2S BCO LUCIANO FERNANDES DOS SANTOS - 2S FERNANDES - 6156878',
                'SO BCO FRANCISCO TORRES SOARES - SO SOARES - 3502201',
                '2S BCO ELIAS EDVANEIDSON DE ARAUJO CARDOSO - 2S ELIAS - 6328512',
                '2S BCO ALISSON LOURENÇO DA SILVA - 2S LOURENÇO - 6780245'
            ]
        }
    }

    @staticmethod
    def inicializar_json():
        """Lê o efetivo.json. Se não existir, usa a lista padrão para criá-lo automaticamente."""
        caminho_json = "efetivo.json"
        if os.path.exists(caminho_json):
            with open(caminho_json, 'r', encoding='utf-8') as f:
                return json.load(f)

        # Se o sistema nunca rodou, cria o Banco de Dados com as listas base originais!
        dados_iniciais = {"SMC": [], "BCT": [], "OEA": []}
        
        nomes_smc = [
            'MAJ AV THIAGO GEALH DE CAMPOS - MAJ GEALH', 'MAJ QOECTA EVERTON RIBEIRO DA COSTA - MAJ EVERTON',
            'CAP QOECTA FÁBIO CÉSAR SILVA DE OLIVEIRA - CAP FÁBIO CÉSAR', 'CAP QOECTA JOSÉ GUILHERME MALTA - CAP MALTA',
            'CAP QOECTA BARBARA PACHECO LINS - CAP BARBARA', 'CAP QOECTA EDGAR HENRIQUE ESCOBAR DOS SANTOS - CAP ESCOBAR',
            '1T AV DANIEL BUERY DE MELO CAMPELO - 1T CAMPELO', '1T QOECTA JAKSON DA SILVA - 1T JAKSON', '1T QOECTA JULIMAR CERUTTI DA SILVA - 1T CERUTTI'
        ]
        legendas_smc = ['GEA', 'EVT', 'FCS', 'MAL', 'BAR', 'ESC', 'CPL', 'JAK', 'CER']
        for i, linha in enumerate(nomes_smc):
            p = [x.strip() for x in linha.split('-')]
            dados_iniciais["SMC"].append({"nome_completo": p[0], "nome_guerra": p[1], "legenda": legendas_smc[i], "saram": ""})

        nomes_bct = DadosEfetivo.HISTORICO["0126"]["bct"]
        for i, linha in enumerate(nomes_bct):
            p = [x.strip() for x in linha.split('-')]
            dados_iniciais["BCT"].append({"nome_completo": p[0], "nome_guerra": p[1], "legenda": chr(65+i), "saram": p[2] if len(p)>2 else ""})

        nomes_oea = DadosEfetivo.HISTORICO["0126"]["oea"]
        for i, linha in enumerate(nomes_oea):
            p = [x.strip() for x in linha.split('-')]
            dados_iniciais["OEA"].append({"nome_completo": p[0], "nome_guerra": p[1], "legenda": chr(65+i), "saram": p[2] if len(p)>2 else ""})

        with open(caminho_json, 'w', encoding='utf-8') as f:
            json.dump(dados_iniciais, f, indent=4, ensure_ascii=False)

        return dados_iniciais

    @staticmethod
    def mapear_efetivo(mes=None, ano_curto=None):
        """Retorna os dicionários de militares prontos para o sistema usar."""
        chave = f"{mes}{ano_curto}" if mes and ano_curto else None
        dados_hist = DadosEfetivo.HISTORICO.get(chave) if chave else None

        mapa_smc, mapa_bct, mapa_oea = {}, {}, {}

        # 1. Se for um mês do passado, usa o HISTORICO cravado no código (para não quebrar auditorias antigas)
        if dados_hist:
            for i, linha in enumerate(dados_hist.get("oea", [])):
                p = [x.strip() for x in linha.split('-')]
                if len(p) >= 3: mapa_oea[p[1].upper()] = {"legenda": chr(65+i), "saram": p[2]}
            for i, linha in enumerate(dados_hist.get("bct", [])):
                p = [x.strip() for x in linha.split('-')]
                if len(p) >= 3: mapa_bct[p[1].upper()] = {"legenda": chr(65+i), "saram": p[2]}
                
            dados_json = DadosEfetivo.inicializar_json() # SMC sempre puxa do JSON
            for item in dados_json["SMC"]: mapa_smc[item["nome_guerra"].upper()] = {"legenda": item["legenda"], "saram": item["saram"]}

        # 2. Se for mês atual/futuro, SUGA os dados 100% do ficheiro JSON que você editar no CRUD!
        else:
            dados = DadosEfetivo.inicializar_json()
            for item in dados["SMC"]: mapa_smc[item["nome_guerra"].upper()] = {"legenda": item["legenda"], "saram": item["saram"]}
            for item in dados["BCT"]: mapa_bct[item["nome_guerra"].upper()] = {"legenda": item["legenda"], "saram": item["saram"]}
            for item in dados["OEA"]: mapa_oea[item["nome_guerra"].upper()] = {"legenda": item["legenda"], "saram": item["saram"]}

        return mapa_smc, mapa_bct, mapa_oea