"""Encabezados explícitos y traslado seguro del índice, sin inventar capítulos."""
from dataclasses import dataclass
import re

from docx.text.paragraph import Paragraph
from .corrector_titulos import capitalizar, normalizar
from .lector_documentos import texto

ROMANOS = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
           'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX']
ROM = re.compile(r'^(?:Cap[ií]tulo\s+)?([IVXLCDM]+)[.\s:–-]+(.+)$', re.I)
DEC = re.compile(r'^(\d+(?:\.\d+)+)[.\s:–-]+(.+)$')
CAPTION = re.compile(r'^(Tabla|Figura|Gr[aá]fico)\s+\d+\s*[.:–-]', re.I)
SPECIAL = {'abstract': 'abstract', 'resumen ejecutivo': 'resumen', 'indice': 'indice',
           'indice de contenidos': 'indice', 'contenido': 'indice',
           'referencias': 'referencias', 'referencias bibliograficas': 'referencias',
           'bibliografia': 'referencias', 'keywords': 'keywords', 'palabras clave': 'keywords'}


@dataclass
class Encabezado:
    p: object
    clase: str
    numero: str = ''


def clasificar(p):
    s = texto(p._p).strip()
    if CAPTION.match(s) or not s or len(s) > 200:
        return None
    special = SPECIAL.get(normalizar(s.rstrip(':')))
    if special:
        return Encabezado(p, special)
    if p._p.xpath('.//w:drawing | .//w:fldChar | .//w:hyperlink'):
        return None
    # Una frase extensa/terminada en punto no se convierte en encabezado por su número.
    if len(s.split()) > 24 or (s.endswith('.') and not p.style.name.startswith('Heading')):
        return None
    m = ROM.match(s)
    if m and m[1].upper() in ROMANOS:
        return Encabezado(p, 'capitulo', m[1].upper())
    m = DEC.match(s)
    if m:
        return Encabezado(p, 'subcapitulo', m[1])
    if p.style.name.startswith('Heading'):
        return Encabezado(p, 'sin_numero')
    return None


def localizar_indice(doc):
    """Aceptar solo entradas numeradas contiguas tras un encabezado Índice explícito."""
    result = []
    ps = doc.paragraphs
    for i, p in enumerate(ps):
        h = clasificar(p)
        if h and h.clase == 'indice':
            block = [p]
            for q in ps[i+1:]:
                s = texto(q._p).strip()
                if not s:
                    if q._p.xpath('.//w:drawing | .//w:fldChar'):
                        break
                    continue
                if (ROM.match(s) or DEC.match(s)) and not q._p.xpath('.//w:hyperlink | .//w:fldChar'):
                    block.append(q)
                else:
                    break
            result.append(block)
    return result


def analizar(doc, excluir=()):
    excluded = {p._p for p in excluir}
    return [h for p in doc.paragraphs if p._p not in excluded and (h := clasificar(p))]


def verificar_estructura(headings):
    errors = []
    chapters = [h for h in headings if h.clase == 'capitulo']
    nums = [ROMANOS.index(h.numero) + 1 for h in chapters]
    if not chapters:
        errors.append('No se identificaron capítulos principales con numeración romana.')
    if nums and nums != list(range(1, max(nums) + 1)):
        errors.append(f'Capítulos duplicados, omitidos o fuera de orden: {nums}.')
    active, previous = 0, {}
    for h in headings:
        if h.clase == 'capitulo':
            active = ROMANOS.index(h.numero) + 1
        elif h.clase == 'subcapitulo':
            parts = tuple(map(int, h.numero.split('.')))
            parent = parts[:-1]
            if parts[0] != active or parts[-1] != previous.get(parent, 0) + 1:
                errors.append(f'Subcapítulo fuera de secuencia o de capítulo: {h.numero}.')
            if len(parts) > 2 and parent not in previous.get('vistos', set()):
                errors.append(f'Subcapítulo sin padre explícito: {h.numero}.')
            previous[parent] = parts[-1]
            previous.setdefault('vistos', set()).add(parts)
        elif h.clase == 'sin_numero':
            errors.append(f'Encabezado sin clasificación segura: {texto(h.p._p)}')
    for kind in ('abstract', 'resumen', 'referencias'):
        count = sum(h.clase == kind for h in headings)
        if count != 1:
            errors.append(f'Se esperaba una sección {kind}; encontradas: {count}. No se inventó contenido.')
    return errors


def organizar(doc, auditor, config, entry, issues):
    blocks = localizar_indice(doc)
    excluded = [p for b in blocks for p in b]
    headings = analizar(doc, excluded)
    issues.extend(verificar_estructura(headings))
    for h in headings:
        if h.clase in ('capitulo', 'subcapitulo', 'abstract', 'resumen', 'referencias', 'sin_numero'):
            try:
                if re.search(r'["“”«»]', texto(h.p._p)):
                    issues.append('Encabezado con cita textual conservado; capitalización pendiente.')
                    continue
                auditor.cambiar(h.p, capitalizar(texto(h.p._p), config), 'Capitalización de encabezado')
            except ValueError as exc:
                issues.append(str(exc))
    major = [h for h in headings if h.clase in ('abstract', 'resumen', 'capitulo', 'referencias')]
    order = [h.clase for h in major]
    if order and order != sorted(order, key=lambda k: {'abstract':0, 'resumen':1, 'capitulo':2, 'referencias':3}[k]):
        # Mover secciones completas solo cuando cada bloque tiene un límite inequívoco.
        nodes = list(doc.element.body)
        starts = [nodes.index(h.p._p) for h in major]
        chunks = [nodes[a:(starts[i+1] if i+1 < len(starts) else len(nodes))]
                  for i, a in enumerate(starts)]
        anchor = major[0].p._p
        ordered = sorted(zip(major, chunks), key=lambda v: {'abstract':0, 'resumen':1, 'capitulo':2, 'referencias':3}[v[0].clase])
        # Guardar posición de inserción antes de retirar nodos; sectPr siempre al final.
        at = starts[0]
        for chunk in chunks:
            for el in chunk:
                if el.tag.endswith('}sectPr'):
                    continue
                doc.element.body.remove(el)
        for _, chunk in ordered:
            for el in chunk:
                if not el.tag.endswith('}sectPr'):
                    doc.element.body.insert(at, el); at += 1
        auditor.cambios.append({'tipo':'Reordenación de secciones completas', 'antes':order,
                                'despues':[h.clase for h,_ in ordered]})
    title = next((p for p in doc.paragraphs if texto(p._p).strip()
                  and p._p not in {q._p for q in excluded}), None)
    if entry:
        if title and not clasificar(title):
            if config['titulos']['respetar_indice_oficial']:
                auditor.cambiar(title, entry.titulo, 'Título oficial verificado')
            elif texto(title._p) != entry.titulo:
                issues.append('Título conservado por configuración; revisar diferencia respecto al índice oficial.')
        else:
            title = doc.add_paragraph(entry.titulo, 'Title')
            doc.element.body.insert(0, title._p)
            auditor.registrar_nuevo(title, 'Portada: título oficial')
        title.style = doc.styles['Title']
    elif title and not clasificar(title):
        # No alterar título sin correspondencia oficial segura.
        title.style = doc.styles['Title']
    primary = next((h.p for h in analizar(doc, excluded) if h.clase in ('abstract','resumen','capitulo')), None)
    if title:
        if entry:
            label = doc.add_paragraph(f'Lectura {entry.numero:02d}')
            title._p.addprevious(label._p)
            auditor.registrar_nuevo(label, 'Número de lectura')
        if config['estructura']['programa_academico']:
            p = doc.add_paragraph(config['estructura']['programa_academico'])
            title._p.addnext(p._p); auditor.registrar_nuevo(p, 'Programa académico')
    if len(blocks) > 1:
        issues.append('Múltiples índices internos; se conservaron para revisión manual.')
    elif primary:
        if blocks:
            block = blocks[0]
            original = [texto(p._p) for p in block[1:]]
        else:
            p = doc.add_paragraph('Índice')
            auditor.registrar_nuevo(p, 'Índice generado a partir de encabezados reales')
            block, original = [p], []
        targets = [h for h in headings if h.clase == 'capitulo' or
                   (config['estructura']['incluir_subcapitulos_indice'] and h.clase == 'subcapitulo')]
        # Trasladar el bloque completo y armonizar por número, no por posición.
        for p in block:
            primary._p.addprevious(p._p)
            p.style = doc.styles['Normal']
            p.paragraph_format.page_break_before = False
        expected = [texto(h.p._p) for h in targets]
        if original != expected:
            auditor.cambios.append({'tipo':'Índice interno desactualizado', 'antes':original, 'despues':expected})
        for i, target in enumerate(expected):
            if i+1 < len(block):
                auditor.cambiar(block[i+1], target, 'Sincronización de entrada de índice')
            else:
                p = doc.add_paragraph(target)
                primary._p.addprevious(p._p)
                auditor.registrar_nuevo(p, 'Entrada de índice generada')
        if len(block) > len(expected)+1:
            issues.append('Entradas sobrantes del índice conservadas; revisar posible capítulo ausente.')
        if block[0].runs: block[0].runs[0].bold = True
    return analizar(doc, [p for b in localizar_indice(doc) for p in b])
