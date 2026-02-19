import os

class Config:
    CAMINHO_RAIZ = r"E:\dev\sistema_lro\ARQUIVOS" if os.name == 'nt' else "."
    
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
    bg_BLUE = '\033[44m'
    WHITE = '\033[97m'
    GREY = '\033[90m'
    RESET = '\033[0m'
    ORANGE = '\033[38;5;208m'

class DadosEfetivo:
    legendas_oea = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    legendas_bct = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
    legendas_smc = ['GEA', 'EVT', 'FCS', 'MAL', 'BAR', 'ESC', 'CPL', 'JAK', 'CER']

    nomes_oea = [
        '1S QSS BCO LUIZ FILIPE TERBECK - 1S TERBECK - 6032338',          # A
        '1S BCO RUI ANTONIO DOS SANTOS JUNIOR - 1S RUI - 6032311',        # B
        '2S BCO SAULO JULIO SANTOS GONÇALVES - 2S SAULO - 6134262',       # C
        '2S BCO KEVIN SOUZA DE OLIVEIRA - 2S KEVIN - 6157912',            # D
        '2S BCO ALAN SITORSKI - 2S SITORSKI - 6156568',                   # E
        '2S BCO LUCIANO FERNANDES DOS SANTOS - 2S FERNANDES - 6156878',   # F
        'SO BCO FRANCISCO TORRES SOARES - SO SOARES - 3502201',           # G
        '2S BCO ELIAS EDVANEIDSON DE ARAUJO CARDOSO - 2S ELIAS - 6328512',# H
        '2S BCO ALISSON LOURENÇO DA SILVA - 2S LOURENÇO - 6780245'        # I
    ]

    nomes_bct = [
        'SO BCT MARCELO DE SOUZA PEREIRA - SO SOUZA - 3988554',           # A
        '1S BCT DIOGO FAVARO WUNDERLICH - 1S FAVARO - 4202317',           # B
        '1S BCT RÉGIS FERRARI - 1S RÉGIS - 4378717',                      # C
        '1S BCT CLEON FRAGA DOS SANTOS - 1S CLEON - 4378660',             # D
        '1S BCT MARCELO ZANOTTO GONÇALVES - 1S ZANOTTO - 6100775',        # E
        '1S BCT GIOVANNI COCCO ALVES - 1S GIOVANNI - 6158080',            # F
        '2S BCT YURI DE PAIVA WERNECK - 2S WERNECK - 6301371',            # G
        '2S BCT EDSON DE SOUZA NUNES - 2S EDSON - 6750451',               # H
        '1S BCT GUSTAVO FRACAO LAGO - 1S GUSTAVO - 4378482',              # I
        '1S BCT CRISTIANO PAZ PRATES - 1S PRATES - 4378768',              # J
        '1S BCT EVANDRO MACHADO BITTENCOURT - 1S BITTENCOURT - 6087841',  # K
        '2S BCT LIS CECI LYRA FONTES - 2S LIS - 6301118',                 # L
        '2S BCT EDUARDO OLIVEIRA DE CASTRO - 2S DE CASTRO - 6576044'      # M
    ]

    nomes_smc = [
        'MAJ AV THIAGO GEALH DE CAMPOS - MAJ GEALH',                      # GEA
        'MAJ QOECTA EVERTON RIBEIRO DA COSTA - MAJ EVERTON',              # EVT
        'CAP QOECTA FÁBIO CÉSAR SILVA DE OLIVEIRA - CAP FÁBIO CÉSAR',     # FCS
        'CAP QOECTA JOSÉ GUILHERME MALTA - CAP MALTA',                    # MAL
        'CAP QOECTA BARBARA PACHECO LINS - CAP BARBARA',                  # BAR
        'CAP EDGAR HENRIQUE ESCOBAR DOS SANTOS - CAP ESCOBAR',            # ESC
        '1T AV DANIEL BUERY DE MELO CAMPELO - 1T CAMPELO',                # CPL
        '1T QOECTA JAKSON DA SILVA - 1T JAKSON',                          # JAK
        '1T QOECTA JULIMAR CERUTTI DA SILVA - 1T CERUTTI'                 # CER
    ]

    @staticmethod
    def mapear_efetivo():
        mapa_smc, mapa_bct, mapa_oea = {}, {}, {}
        for i, linha in enumerate(DadosEfetivo.nomes_oea):
            partes = [p.strip() for p in linha.split('-')]
            if len(partes) >= 3: mapa_oea[partes[1].upper()] = {"legenda": DadosEfetivo.legendas_oea[i], "saram": partes[2]}
        for i, linha in enumerate(DadosEfetivo.nomes_bct):
            partes = [p.strip() for p in linha.split('-')]
            if len(partes) >= 3: mapa_bct[partes[1].upper()] = {"legenda": DadosEfetivo.legendas_bct[i], "saram": partes[2]}
        for i, linha in enumerate(DadosEfetivo.nomes_smc):
            partes = [p.strip() for p in linha.split('-')]
            if len(partes) >= 2: mapa_smc[partes[1].upper()] = {"legenda": DadosEfetivo.legendas_smc[i], "saram": None}
        return mapa_smc, mapa_bct, mapa_oea