"""Renderizado real y comprobaciones geométricas; no sustituyen juicio visual."""
from pathlib import Path
import os
import shutil
import subprocess

import fitz
from .corrector_titulos import normalizar
from .lector_documentos import abrir, texto
from .organizador_capitulos import analizar, localizar_indice


def herramienta(config, name):
    configured = config['herramientas'][name]
    if configured: return configured
    found = shutil.which(name)
    if found: return found
    if name == 'soffice' and os.name == 'nt':
        for base in ('PROGRAMFILES','PROGRAMFILES(X86)'):
            path = Path(os.environ.get(base,''))/'LibreOffice/program/soffice.exe'
            if path.is_file(): return str(path)
    return None


def renderizar(docx, output, config):
    output.mkdir(parents=True, exist_ok=True)
    pdf = output/(docx.stem+'.pdf')
    if pdf.exists(): pdf.unlink()  # Solo salidas temporales de la ejecución actual.
    errors = []
    if os.name == 'nt' and config['herramientas']['word_com']:
        app = None
        try:
            import win32com.client
            app = win32com.client.DispatchEx('Word.Application')
            app.Visible = False; app.DisplayAlerts = 0; app.AutomationSecurity = 3
            app.Options.UpdateLinksAtOpen = False
            doc = app.Documents.Open(str(docx.resolve()), ReadOnly=True, AddToRecentFiles=False,
                                     ConfirmConversions=False)
            try: doc.ExportAsFixedFormat(str(pdf.resolve()), 17)
            finally: doc.Close(False)
            if pdf.is_file(): return pdf, 'Microsoft Word COM'
        except Exception as exc:
            errors.append(f'Word COM: {exc}')
        finally:
            if app is not None: app.Quit()
    tool = herramienta(config,'soffice')
    if tool:
        profile = output/'perfil_libreoffice'
        result = subprocess.run([str(tool), '-env:UserInstallation='+profile.resolve().as_uri(),
                                 '--headless','--convert-to','pdf:writer_pdf_Export',
                                 '--outdir',str(output.resolve()),str(docx.resolve())],
                                capture_output=True, text=True,
                                timeout=config['herramientas']['timeout_seg'])
        if result.returncode == 0 and pdf.is_file() and pdf.stat().st_size:
            return pdf, 'LibreOffice'
        errors.append(f'LibreOffice ({result.returncode}): {result.stdout} {result.stderr}')
    raise RuntimeError('No se pudo renderizar localmente. '+('; '.join(errors) or 'Instale LibreOffice o Word con pywin32.'))


def inspeccionar_pdf(pdf, doc, config, previews=None, diagramas=()):
    errors, metrics = [], []
    margin = config['formato']['margenes_cm']*72/2.54
    titles = [h for h in analizar(doc, [p for b in localizar_indice(doc) for p in b])
              if h.clase in ('abstract','resumen','capitulo','referencias')]
    pages_for = {}
    # Leer partes existentes: acceder a .first_page_header.paragraphs crearía partes
    # nuevas, lo que no corresponde a una operación de inspección.
    header_text = normalizar(' '.join(texto(part.element) for part in doc.part.package.parts
                            if str(part.partname).startswith(('/word/header','/word/footer'))
                            and hasattr(part,'element')))
    if previews: previews.mkdir(parents=True, exist_ok=True)
    with fitz.open(pdf) as rendered:
        for i, page in enumerate(rendered):
            content = normalizar(page.get_text())
            images = page.get_image_info()
            bounds = page.rect
            useful = fitz.Rect(margin-2,margin-2,bounds.width-margin+2,bounds.height-margin+2)
            if not content and not images and not page.get_drawings():
                errors.append(f'Página {i+1} vacía inesperada.')
            line_rects = []
            for block in page.get_text('rawdict')['blocks']:
                if block['type'] != 0: continue
                for line in block['lines']:
                    chars = [c for span in line['spans'] for c in span['chars']]
                    line_text = normalizar(''.join(c['c'] for c in chars))
                    ink = [c for c in chars if not c['c'].isspace()]
                    if not ink: continue
                    rect = fitz.Rect(ink[0]['bbox'])
                    for char in ink[1:]: rect |= fitz.Rect(char['bbox'])
                    is_header_footer = (line_text and line_text in header_text and
                                        (rect.y1 < margin or rect.y0 > bounds.height-margin))
                    if is_header_footer:
                        if not bounds.contains(rect): errors.append(f'Encabezado/pie recortado en página {i+1}.')
                        continue
                    if not useful.contains(rect):
                        errors.append(f'Texto fuera de márgenes en página {i+1}.')
                    for other in line_rects:
                        intersect = rect & other
                        if intersect.width > 2 and intersect.height > 3:
                            errors.append(f'Posible superposición de texto en página {i+1}.')
                            break
                    line_rects.append(rect)
            for im in images:
                rect = fitz.Rect(im['bbox'])
                if not useful.contains(rect): errors.append(f'Figura fuera de márgenes en página {i+1}.')
                if any((rect & t).width > 3 and (rect & t).height > 3 for t in line_rects):
                    errors.append(f'Posible superposición de figura/texto en página {i+1}.')
                for diagram in diagramas:
                    if not diagram.get('convertido'): continue
                    from PIL import Image
                    with Image.open(diagram['png']) as png:
                        match = png.width == im['width'] and png.height == im['height']
                    if match:
                        actual = diagram['fuente_final_pt'] * rect.width / (diagram['ancho_in']*72)
                        diagram['pagina_pdf'] = i+1
                        diagram['fuente_renderizada_pt'] = round(actual,2)
                        if actual < config['figuras']['tamano_texto_minimo_pt']-.1:
                            errors.append(f'Texto de diagrama ilegible en página {i+1} ({actual:.1f} pt).')
            for j, h in enumerate(titles):
                if normalizar(texto(h.p._p)) in content:
                    # La copia de capítulos en el índice de portada no es su inicio.
                    if i > 0 or h.clase in ('abstract','resumen'):
                        pages_for.setdefault(j,i+1)
            metrics.append({'pagina':i+1,'imagenes':len(images),'lineas':len(line_rects)})
            if previews:
                page.get_pixmap(matrix=fitz.Matrix(1.2,1.2)).save(previews/f'pagina_{i+1:03d}.png')
        previous = 1
        for j, h in enumerate(titles):
            at = pages_for.get(j)
            if at is None:
                errors.append(f'No se localizó el inicio renderizado: {texto(h.p._p)}')
            elif at <= previous:
                errors.append(f'Secciones sin página independiente o fuera de orden: {texto(h.p._p)} (pág. {at}).')
            else: previous = at
        for diagram in diagramas:
            if diagram.get('convertido') and 'pagina_pdf' not in diagram:
                errors.append('No se identificó la figura Mermaid en el PDF para verificar su escala final.')
        return {'paginas':len(rendered), 'inicios':[
            {'titulo':texto(h.p._p),'pagina':pages_for.get(j)} for j,h in enumerate(titles)],
            'paginas_inspeccionadas':metrics,'errores':list(dict.fromkeys(errors)),
            'revision':'Geometría y texto extraído del PDF; semántica y legibilidad de imágenes preexistentes requieren revisión humana.'}


def validar_estructura(path, auditor):
    reopened = abrir(path)
    return auditor.validar(reopened)


def reparar_saltos(doc):
    """Eliminar saltos manuales redundantes de párrafos vacíos sin eliminar párrafos."""
    from docx.oxml.ns import qn
    changed = 0
    for p in doc.paragraphs:
        if texto(p._p).strip() or p._p.xpath('.//w:drawing | .//w:object | .//w:fldChar'): continue
        for br in p._p.xpath('.//w:br[@w:type="page"]'):
            br.getparent().remove(br); changed += 1
        if p.paragraph_format.page_break_before:
            p.paragraph_format.page_break_before = False; changed += 1
    return changed
