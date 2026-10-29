"""Lectura OOXML y ediciones segmentadas, sin reconstruir párrafos."""
from collections import Counter
from dataclasses import dataclass, field
import difflib
import hashlib
from pathlib import Path
import re
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from lxml import etree


def tokens(element):
    return element.xpath('.//w:t | .//w:tab | .//w:br[not(@w:type="page") and not(@w:type="column")] | .//w:cr')


def token_text(node):
    if node.tag == qn('w:t'): return node.text or ''
    if node.tag == qn('w:tab'): return '\t'
    return '\n'


def texto(element) -> str:
    return ''.join(token_text(node) for node in tokens(element))


def establecer_token(node, value):
    if node.tag == qn('w:t'):
        node.text = value
        node.set(qn('xml:space'),'preserve')
        return
    if value == token_text(node): return
    parent = node.getparent()
    if value:
        replacement = OxmlElement('w:t'); replacement.text = value
        replacement.set(qn('xml:space'),'preserve')
        parent.insert(parent.index(node),replacement)
    parent.remove(node)


def parrafos(doc):
    for element in doc.element.body.xpath('.//w:p'):
        yield Paragraph(element, doc._body)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def abrir(path: Path):
    # Limitar documentos ZIP anómalos antes de descomprimir OOXML.
    with ZipFile(path) as z:
        if sum(i.file_size for i in z.infolist()) > 250_000_000:
            raise ValueError('DOCX demasiado grande (límite descomprimido: 250 MB).')
        if 'word/document.xml' not in z.namelist():
            raise ValueError('No es un documento Word DOCX válido.')
        for name in z.namelist():
            if 'vbaProject' in name:
                raise ValueError('No se admiten documentos con macros.')
    return Document(path)


def xml_canon(element):
    return etree.tostring(element, method='c14n')


def protegidos(doc):
    # Estos elementos no se editan ni siquiera para corregir espacios.
    tags = ['w:hyperlink', 'w:fldSimple', 'w:instrText', 'm:oMath', 'm:oMathPara',
            'w:footnoteReference', 'w:endnoteReference', 'w:object']
    counter = Counter()
    ns = dict(doc.element.nsmap)
    ns['m'] = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    for tag in tags:
        for el in doc.element.xpath(f'.//{tag}'):
            counter[(tag, xml_canon(el))] += 1
    return counter


def tablas_firma(doc):
    return [(
        tuple(tuple(texto(tc) for tc in tr.xpath('./w:tc')) for tr in t.xpath('./w:tr')),
        tuple(t.xpath('.//w:gridSpan/@w:val')), tuple(t.xpath('.//w:vMerge/@w:val')),
    ) for t in doc.element.body.xpath('.//w:tbl')]


@dataclass
class Auditor:
    doc: object
    cambios: list = field(default_factory=list)

    def __post_init__(self):
        self.esperados = {p._p: texto(p._p) for p in parrafos(self.doc)}
        self.tablas = tablas_firma(self.doc)
        self.protegidos = protegidos(self.doc)
        self.medios = {p.partname: hashlib.sha256(p.blob).hexdigest()
                       for p in self.doc.part.package.parts if '/media/' in str(p.partname)}
        self.partes = {str(p.partname): hashlib.sha256(p.blob).hexdigest()
                       for p in self.doc.part.package.parts
                       if str(p.partname).startswith(('/word/header', '/word/footer', '/word/footnotes',
                                                    '/word/endnotes', '/word/comments'))}
        self.relaciones = {(r.reltype, str(r.target_ref)) for r in self.doc.part.rels.values()}
        self.original = [{'texto': v, 'xml': etree.tostring(k, encoding='unicode')}
                         for k, v in self.esperados.items()]

    def registrar_nuevo(self, p, motivo):
        self.esperados[p._p] = texto(p._p)
        self.cambios.append({'tipo': motivo, 'antes': '', 'despues': texto(p._p)})

    def eliminar_diagrama(self, p, archivo_auditoria):
        self.cambios.append({'tipo': 'Mermaid validado a figura', 'antes': texto(p._p),
                             'despues': '', 'auditoria': str(archivo_auditoria)})
        self.esperados.pop(p._p)
        p._p.getparent().remove(p._p)

    def cambiar(self, p, nuevo, motivo):
        antes = texto(p._p)
        if antes == nuevo:
            return
        if p._p.xpath('.//w:hyperlink | .//w:fldChar | .//w:fldSimple | .//m:oMath | .//w:drawing | .//w:object'):
            raise ValueError('Edición rechazada: párrafo con objetos/campos/hipervínculos.')
        nodes = tokens(p._p)
        # Operar de derecha a izquierda conserva offsets y el formato de los runs.
        edits = list(difflib.SequenceMatcher(None, antes, nuevo, autojunk=False).get_opcodes())
        for op, a, b, c, d in reversed(edits):
            if op == 'equal':
                continue
            nodes = tokens(p._p)
            starts, pos = [], 0
            for n in nodes:
                starts.append(pos)
                pos += len(token_text(n))
            touched = [i for i, n in enumerate(nodes)
                       if starts[i] < b and starts[i] + len(token_text(n)) > a]
            if not touched:
                # Inserción en un límite de run o en un párrafo vacío.
                idx = next((i for i, n in enumerate(nodes)
                            if starts[i] <= a <= starts[i] + len(token_text(n))), None)
                if idx is None:
                    p.add_run(nuevo[c:d])
                    nodes = p._p.xpath('.//w:t')
                else:
                    n = nodes[idx]; off = a - starts[idx]
                    establecer_token(n, token_text(n)[:off] + nuevo[c:d] + token_text(n)[off:])
            else:
                for i in touched:
                    n = nodes[i]; s = token_text(n)
                    left = max(0, a - starts[i]); right = min(len(s), b - starts[i])
                    establecer_token(n, s[:left] + (nuevo[c:d] if i == touched[0] else '') + s[right:])
        if texto(p._p) != nuevo:
            raise ValueError('La edición segmentada no produjo el texto esperado.')
        self.esperados[p._p] = nuevo
        self.cambios.append({'tipo': motivo, 'antes': antes, 'despues': nuevo})

    def validar(self, reabierto=None):
        actuales = {p._p: texto(p._p) for p in parrafos(self.doc)}
        errors = []
        if actuales != self.esperados:
            errors.append('Pérdida, inserción o modificación de párrafos no autorizada.')
        doc = reabierto or self.doc
        if reabierto and Counter(texto(p._p) for p in parrafos(doc)) != Counter(self.esperados.values()):
            errors.append('El texto cambió al guardar/reabrir el DOCX.')
        if tablas_firma(doc) != self.tablas:
            errors.append('Se alteraron filas, celdas, combinaciones o datos de tablas.')
        if protegidos(doc) != self.protegidos:
            errors.append('Se alteraron hipervínculos, campos, fórmulas o referencias protegidas.')
        parts = {str(p.partname): p for p in doc.part.package.parts}
        for name, digest in {**{str(k):v for k,v in self.medios.items()}, **self.partes}.items():
            if name not in parts or hashlib.sha256(parts[name].blob).hexdigest() != digest:
                errors.append(f'Se alteró o perdió una parte original: {name}')
        if not self.relaciones.issubset({(r.reltype, str(r.target_ref)) for r in doc.part.rels.values()}):
            errors.append('Se perdió una relación original del documento.')
        return errors


def es_protegido(p):
    t = texto(p._p)
    return bool(p._p.xpath('.//w:hyperlink | .//w:fldChar | .//w:fldSimple | .//m:oMath | .//w:object')
                or re.search(r'''["“”«»‘’]|'[^'\n]+'|\([A-ZÁÉÍÓÚÑ][^)]*\d{4}[^)]*\)|\b\d+\b|[=∑∫]''', t))
