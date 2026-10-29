from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips
from docx.text.paragraph import Paragraph
from .corrector_titulos import capitalizar
from .lector_documentos import texto
import re


def caption(doc, element, kind, number, auditor, config, issues):
    prev = element.getprevious()
    found = None
    # No atravesar otros objetos ni párrafos con contenido.
    while prev is not None and prev.tag == qn('w:p'):
        s = texto(prev).strip()
        if s:
            if re.match(rf'^{kind}\s+\d+\s*[.:–-]', s, re.I):
                found = Paragraph(prev, doc._body)
            break
        if prev.xpath('.//w:drawing | .//w:object'): break
        prev = prev.getprevious()
    if found:
        old = texto(found._p)
        new = re.sub(rf'^{kind}\s+\d+\s*[.:–-]\s*', f'{kind} {number}. ', old, flags=re.I)
        new = capitalizar(new, config)
        try: auditor.cambiar(found, new, f'Título y numeración de {kind.lower()}')
        except ValueError as exc: issues.append(str(exc))
        if re.match(rf'^{kind}\s+{number}\b', old, re.I) is None:
            issues.append(f'{kind} {number}: numeración actualizada; revisar remisiones textuales.')
    else:
        context = 'Contenido Académico'
        before = element.getprevious()
        from .organizador_capitulos import clasificar
        while before is not None:
            if before.tag == qn('w:p'):
                h = clasificar(Paragraph(before, doc._body))
                if h and h.clase in ('capitulo','subcapitulo'):
                    context = re.sub(r'^(?:[IVXLCDM]+|\d+(?:\.\d+)+)[.\s:–-]+', '', texto(before))
                    break
            before = before.getprevious()
        found = doc.add_paragraph(capitalizar(f'{kind} {number}. {context}', config))
        element.addprevious(found._p)
        auditor.registrar_nuevo(found, f'Título contextual de {kind.lower()} propuesto')
        issues.append(f'{kind} {number} sin título: título contextual propuesto, confirmar correspondencia con contenido.')
    found.paragraph_format.keep_with_next = True
    found.paragraph_format.keep_together = True
    for run in found.runs: run.bold = True
    return texto(found._p)


def procesar(doc, auditor, config, issues):
    results = []
    # Las tablas anidadas conservan su contenido; se informa para revisión del ancho.
    tables = doc.tables
    width_twips = int((doc.sections[0].page_width-doc.sections[0].left_margin-doc.sections[0].right_margin)/635)
    for n, table in enumerate(tables, 1):
        title = caption(doc, table._tbl, 'Tabla', n, auditor, config, issues)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER; table.autofit = False
        grid = table._tbl.tblGrid.gridCol_lst
        widths = [int(col.w or width_twips / max(1,len(grid))) for col in grid]
        total = sum(widths) or 1
        new = [int(width_twips*w/total) for w in widths]
        for col, w in zip(grid,new): col.w = Twips(w)
        for row in table.rows:
            col = 0
            for tc in row._tr.tc_lst:
                span = int(tc.tcPr.gridSpan.val) if tc.tcPr is not None and tc.tcPr.gridSpan is not None else 1
                tc.get_or_add_tcPr().get_or_add_tcW().set(qn('w:w'),str(sum(new[col:col+span])))
                tc.tcPr.tcW.set(qn('w:type'),'dxa'); col += span
            pr = row._tr.get_or_add_trPr()
            if not pr.findall(qn('w:cantSplit')): pr.append(OxmlElement('w:cantSplit'))
        if table.rows:
            pr = table.rows[0]._tr.get_or_add_trPr()
            if not pr.findall(qn('w:tblHeader')): pr.append(OxmlElement('w:tblHeader'))
        if table._tbl.xpath('.//w:tbl'):
            issues.append(f'Tabla {n}: contiene tablas anidadas; revisar ajuste y paginación manualmente.')
        issues.append(f'Tabla {n}: verificar correspondencia semántica del título y nota de fuente.')
        results.append({'numero':n, 'titulo':title, 'filas':len(table.rows), 'columnas':len(grid),
                        'ancho_twips':sum(new), 'datos':'conservados', 'revision_semantica':'pendiente'})
    return results
