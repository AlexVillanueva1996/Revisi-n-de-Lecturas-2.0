from io import BytesIO
from PIL import Image
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from .procesador_tablas import caption


def procesar(doc, auditor, config, issues):
    results = []
    sec = doc.sections[0]
    maxw = int(sec.page_width-sec.left_margin-sec.right_margin)
    maxh = int(sec.page_height-sec.top_margin-sec.bottom_margin) - 12700*80
    for n, drawing in enumerate(doc.element.body.xpath('.//w:drawing'), 1):
        p_el = drawing
        while p_el is not None and p_el.tag != qn('w:p'): p_el = p_el.getparent()
        if p_el is None: continue
        if p_el.getparent() != doc.element.body:
            issues.append(f'Figura {n}: imagen en tabla/cuadro, posición y título requieren revisión manual.')
            results.append({'numero':n,'revision_visual_humana':'pendiente: figura anidada conservada'})
            continue
        for inline in drawing.xpath('.//wp:inline'):
            for attr in ('distL','distR','distT','distB'): inline.set(attr,'0')
        p = Paragraph(p_el, doc._body)
        title = caption(doc, p_el, 'Figura', n, auditor, config, issues)
        p.paragraph_format.keep_together = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        extents = drawing.xpath('.//wp:extent')
        dims, dpi = None, None
        if extents:
            extent = extents[0]; w, h = int(extent.get('cx')), int(extent.get('cy'))
            if w > 0 and h > 0:
                scale = min(1, maxw/w, maxh/h)
                width, height = int(w*scale), int(h*scale)
                extent.set('cx', str(width)); extent.set('cy', str(height))
                for ext in drawing.xpath('.//a:xfrm/a:ext'):
                    ext.set('cx', str(width)); ext.set('cy', str(height))
                dims = [width/914400, height/914400]
                p.paragraph_format.first_line_indent = Pt(0)
                if width >= maxw*.98:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if scale < 1: issues.append(f'Figura {n}: tamaño ajustado a márgenes; comprobar texto interno.')
        try:
            blips = drawing.xpath('.//a:blip')
            if blips:
                rid = blips[0].get(qn('r:embed'))
                if rid:
                    blob = doc.part.related_parts[rid].blob
                    with Image.open(BytesIO(blob)) as im:
                        im.verify()
                    with Image.open(BytesIO(blob)) as im:
                        if dims: dpi = round(im.width/dims[0])
                    if dpi and dpi < 150:
                        issues.append(f'Figura {n}: resolución efectiva baja ({dpi} DPI).')
                else:
                    issues.append(f'Figura {n}: imagen enlazada externamente; no se descargó.')
        except (OSError, KeyError, ValueError) as exc:
            issues.append(f'Figura {n}: imagen no verificable con Pillow: {exc}')
        if drawing.xpath('.//wp:anchor'):
            issues.append(f'Figura {n}: objeto flotante; revisar superposición y anclaje.')
        issues.append(f'Figura {n}: legibilidad interna, recortes originales, título y fuente requieren revisión visual humana.')
        results.append({'numero':n, 'titulo':title, 'tamano_pulgadas':dims,
                        'dpi_efectivo':dpi,'revision_visual_humana':'pendiente'})
    if doc.element.body.xpath('.//w:pict | .//w:object'):
        issues.append('Hay objetos VML/OLE conservados que requieren revisión visual manual.')
    return results
