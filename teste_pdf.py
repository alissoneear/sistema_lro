import os
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject, TextStringObject

# ==== imports auxiliares para flatten e manipulação avançada ====
from typing import Any, Optional, Tuple
from io import BytesIO

def _resolve(obj: Any) -> Any:
    try:
        return obj.get_object()
    except Exception:
        return obj


def _iter_field_dicts(field: Any):
    """Itera recursivamente pelos dicionários de campos (/Fields e /Kids)."""
    field = _resolve(field)
    if not isinstance(field, dict):
        return

    yield field

    kids = field.get("/Kids")
    if kids:
        for kid in kids:
            yield from _iter_field_dicts(kid)


def _set_acroform_values(writer: PdfWriter, data: dict):
    """Define /V e /DV diretamente nos campos do /AcroForm (mesmo se o widget não estiver em /Annots)."""
    acro = writer._root_object.get("/AcroForm")
    if not acro:
        return

    acro = _resolve(acro)
    fields = acro.get("/Fields")
    if not fields:
        return

    for field in fields:
        for f in _iter_field_dicts(field):
            name = f.get("/T")
            if not name:
                continue
            # name pode ser objeto PDF; converte para str simples
            name_str = str(name)
            if name_str in data:
                val = data[name_str]
                # pypdf exige PdfObject; para strings use TextStringObject
                if not hasattr(val, "get_object") and not str(type(val)).endswith("PdfObject'>"):
                    val_obj = TextStringObject(str(val))
                else:
                    val_obj = val

                f[NameObject("/V")] = val_obj
                f[NameObject("/DV")] = val_obj


def _find_widget_rect(reader: PdfReader, field_name: str) -> Optional[Tuple[int, Tuple[float, float, float, float]]]:
    """Acha em qual página está o widget do campo e retorna (page_index, (llx,lly,urx,ury))."""
    for i, page in enumerate(reader.pages):
        annots = page.get("/Annots") or []
        for a in annots:
            annot = _resolve(a)
            if not isinstance(annot, dict):
                continue

            # O widget pode ter /T direto ou via /Parent
            t = annot.get("/T")
            if not t and annot.get("/Parent"):
                parent = _resolve(annot.get("/Parent"))
                if isinstance(parent, dict):
                    t = parent.get("/T")

            if t and str(t) == field_name:
                rect = annot.get("/Rect")
                if rect and len(rect) == 4:
                    llx, lly, urx, ury = [float(x) for x in rect]
                    return i, (llx, lly, urx, ury)

    return None


def _flatten_by_overlay(input_pdf: str, output_pdf: str, data: dict, font_name: str = "Helvetica", font_size: int = 12):
    """Gera um PDF 'achatado' desenhando o texto por cima do template nas posições dos campos.

    Requer `reportlab`. Se não estiver instalado, levanta um erro amigável.
    """
    try:
        from reportlab.pdfgen import canvas
    except ModuleNotFoundError as e:
        raise ModuleNotFoundError(
            "Para gerar a versão _flat.pdf, instale a dependência: pip install reportlab"
        ) from e

    src = PdfReader(input_pdf)
    out = PdfWriter()

    # Copia páginas
    out.append_pages_from_reader(src)

    for field_name, value in data.items():
        found = _find_widget_rect(src, field_name)
        if not found:
            continue

        page_index, (llx, lly, urx, ury) = found
        page = out.pages[page_index]

        # Tamanho real da página
        media = page.mediabox
        page_w = float(media.width)
        page_h = float(media.height)

        # Cria um overlay em memória
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=(page_w, page_h))
        c.setFont(font_name, font_size)

        # Desenha dentro da caixa do campo (com um pequeno padding)
        x = llx + 2
        y = (lly + ury) / 2 - (font_size * 0.35)
        c.drawString(x, y, str(value))
        c.save()

        packet.seek(0)
        overlay = PdfReader(packet)
        page.merge_page(overlay.pages[0])

    with open(output_pdf, "wb") as f:
        out.write(f)


def inspecionar_campos_pdf(caminho_modelo):
    """Lê o PDF e imprime o nome de todos os campos editáveis disponíveis."""
    reader = PdfReader(caminho_modelo)
    campos = reader.get_fields()
    
    print("--- Campos disponíveis no PDF ---")
    if campos:
        for nome_campo in campos.keys():
            print(f"- {nome_campo}")
    else:
        print("Nenhum campo de formulário encontrado. Verifique o modelo.")
    print("---------------------------------\n")

def preencher_pdf(caminho_modelo, caminho_saida, dicionario_dados):
    """Preenche os campos do modelo PDF e salva como um novo arquivo."""
    reader = PdfReader(caminho_modelo)
    writer = PdfWriter()

    # Copia todas as páginas
    writer.append_pages_from_reader(reader)

    # -----------------------------------------------------------------
    # GARANTE QUE O PdfWriter TENHA O /AcroForm (senão dá: No /AcroForm...)
    # Alguns PDFs têm campos, mas o /AcroForm não é carregado automaticamente.
    # -----------------------------------------------------------------
    root = reader.trailer.get("/Root", {})
    acroform = root.get("/AcroForm")

    if acroform is not None:
        writer._root_object.update({NameObject("/AcroForm"): acroform})

        # Faz os viewers (Adobe/Preview) regenerarem a aparência do texto
        # para que o valor apareça visualmente.
        try:
            writer._root_object["/AcroForm"].update({NameObject("/NeedAppearances"): BooleanObject(True)})
        except Exception:
            # Em alguns PDFs, /AcroForm pode ser um objeto indireto.
            # O update acima pode falhar; ainda assim o preenchimento pode funcionar.
            pass

        # Define valores diretamente no /AcroForm (mais robusto quando o widget não está em /Annots)
        _set_acroform_values(writer, dicionario_dados)

    # Preenche os campos (aplique em todas as páginas para garantir)
    for page in writer.pages:
        try:
            writer.update_page_form_field_values(page, dicionario_dados)
        except Exception:
            # Alguns PDFs têm o campo no AcroForm mas sem widget nesta página
            pass

    with open(caminho_saida, "wb") as output_stream:
        writer.write(output_stream)

    # Gera também uma versão 'achatada' (texto desenhado por cima), que aparece em qualquer viewer
    caminho_flat = caminho_saida.replace(".pdf", "_flat.pdf")
    try:
        _flatten_by_overlay(caminho_saida, caminho_flat, dicionario_dados, font_size=12)
        print(f"✅ Versão achatada gerada em: {caminho_flat}")
    except Exception as e:
        print(f"⚠️ Não consegui gerar a versão achatada: {e}")

    print(f"✅ Arquivo gerado com sucesso em: {caminho_saida}")

# ==========================================
# EXECUÇÃO DO TESTE
# ==========================================
if __name__ == "__main__":
    modelo_pdf = "template_escala.pdf" # Coloque o nome do PDF criado no Passo 1
    saida_pdf = "escala_preenchida_teste.pdf"

    # 1. Verifica se o modelo existe
    if os.path.exists(modelo_pdf):
        # 2. Descobre o nome das variáveis do PDF
        inspecionar_campos_pdf(modelo_pdf)
        
        # 3. Define os dados simulando a operação do ARCC-CW
        # ATENÇÃO: As chaves do dicionário devem ser exatamente iguais aos Nomes definidos no Passo 1
        dados_operacionais = {
            "mes_ano": "CUMPRIDA FEVEREIRO/2026"
        }
        
        # 4. Gera o arquivo final
        preencher_pdf(modelo_pdf, saida_pdf, dados_operacionais)
    else:
        print(f"❌ O arquivo '{modelo_pdf}' não foi encontrado na pasta.")