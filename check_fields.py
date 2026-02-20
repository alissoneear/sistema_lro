import os
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject

def verificar_campos_formulario(caminho_pdf):
    """
    Varre o PDF em busca de campos e salva o resultado em um arquivo .txt.
    """
    if not os.path.exists(caminho_pdf):
        print(f"Erro: Arquivo '{caminho_pdf}' não encontrado.")
        return

    reader = PdfReader(caminho_pdf)
    fields = reader.get_fields()
    
    # Criamos uma lista para armazenar as linhas do relatório
    relatorio = []
    relatorio.append(f"{'='*60}")
    relatorio.append(f"ANÁLISE DE CAMPOS: {os.path.basename(caminho_pdf)}")
    relatorio.append(f"{'='*60}")

    if not fields:
        msg = "Resultado: Nenhum campo de formulário editável foi detectado."
        print(msg)
        relatorio.append(msg)
    else:
        cabecalho = f"Detectados {len(fields)} campos disponíveis:\n"
        print(cabecalho)
        relatorio.append(cabecalho)
        
        titulo_tabela = f"{'NOME DO CAMPO':<30} | {'TIPO DE DADO'}"
        separador = "-" * 60
        
        print(titulo_tabela)
        print(separador)
        relatorio.append(titulo_tabela)
        relatorio.append(separador)

        for name, field in fields.items():
            tipo = field.get('/FT', 'Desconhecido')
            linha = f"{name:<30} | {tipo}"
            print(linha)
            relatorio.append(linha)
    
    relatorio.append(f"{'='*60}\n")
    print(f"{'='*60}\n")

    # --- LÓGICA DE EXTRAÇÃO PARA .TXT ---
    # Define o nome do arquivo baseado no PDF (ex: template_escala.txt)
    nome_txt = os.path.splitext(caminho_pdf)[0] + "_campos.txt"
    
    try:
        with open(nome_txt, 'w', encoding='utf-16') as f: # Usamos utf-16 por segurança com acentos
            f.write("\n".join(relatorio))
        print(f"Relatório exportado com sucesso: {nome_txt}")
    except Exception as e:
        print(f"Erro ao gerar .txt: {e}")


def preencher_escala_teste(caminho_entrada, caminho_saida):
    """
    Preenche os campos detectados no template com dados fictícios de teste.
    """
    if not os.path.exists(caminho_entrada):
        print(f"Erro: Arquivo '{caminho_entrada}' não encontrado.")
        return

    reader = PdfReader(caminho_entrada)
    writer = PdfWriter()

    # CORREÇÃO: Usamos append() para clonar toda a estrutura do PDF (incluindo os AcroForms)
    writer.append(reader)

    # Dados fictícios solicitados
    dados_ficticios = {
        "mes_ano": "CUMPRIDA FEVEREIRO/2026",
        "d1_t1": "B",
        "d1_t2": "C",
        "d1_t3": "B"
    }

    # Preenche os campos do formulário na primeira página
    writer.update_page_form_field_values(writer.pages[0], dados_ficticios)

    # Salva o resultado
    try:
        with open(caminho_saida, "wb") as output_stream:
            writer.write(output_stream)
        print(f"Sucesso! Arquivo de teste gerado em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")

def configurar_template_fevereiro(arquivo_entrada, arquivo_saida):
    if not os.path.exists(arquivo_entrada):
        print(f"Erro: {arquivo_entrada} não encontrado.")
        return

    reader = PdfReader(arquivo_entrada)
    writer = PdfWriter()
    writer.append(reader)

    # Controle de data (iniciando no dia 7 após os campos manuais)
    dia_progresso = 7
    turno_progresso = 1

    # Acesso direto aos campos no AcroForm para renomeação de metadados
    if "/AcroForm" in writer.root_object:
        form = writer.root_object["/AcroForm"]
        if "/Fields" in form:
            fields = form["/Fields"]
            
            for ref in fields:
                field_obj = ref.get_object()
                if "/T" in field_obj:
                    nome_atual = field_obj["/T"]
                    
                    # Filtra campos genéricos text_21 até text_86
                    if nome_atual.startswith("text_"):
                        try:
                            # Extrai o número do sufixo (ex: text_21 -> 21)
                            num = int(nome_atual.split("_")[1])
                            
                            # Limita a renomeação ao intervalo de 28 dias (até d28_t3)
                            if 21 <= num <= 86:
                                novo_nome = f"d{dia_progresso}_t{turno_progresso}"
                                
                                # Atualiza a chave /T (Título do campo)
                                field_obj.update({
                                    NameObject("/T"): TextStringObject(novo_nome)
                                })
                                
                                print(f"Renomeado: {nome_atual} -> {novo_nome}")
                                
                                # Incrementa turno e dia
                                turno_progresso += 1
                                if turno_progresso > 3:
                                    turno_progresso = 1
                                    dia_progresso += 1
                        except (ValueError, IndexError):
                            continue

    with open(arquivo_saida, "wb") as f:
        writer.write(f)
    print(f"\n[OK] Template de Fevereiro (28 dias) pronto: {arquivo_saida}")

if __name__ == "__main__":
    # Caminho do seu template
    template = "bct_2026_fev_temp.pdf"
    saida = "escala_preenchida_teste.pdf"
    verificar_campos_formulario(template)
    #configurar_template_fevereiro(template, "template_fevereiro_final.pdf")
    #preencher_escala_teste(template, saida)