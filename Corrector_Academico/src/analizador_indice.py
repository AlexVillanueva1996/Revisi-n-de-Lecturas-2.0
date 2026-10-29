"""No asignar un título solamente por el nombre del archivo."""
from dataclasses import dataclass
from difflib import SequenceMatcher
import re

from .corrector_titulos import CONECTORES, capitalizar, normalizar
from .lector_documentos import abrir, texto

TITULO = re.compile(r'^(?:Lecci[oó]n|Lectura|Tema|Unidad)\s*0*(\d+)\s*[:.\-–]\s*(.+)$', re.I)


@dataclass
class Entrada:
    numero: int
    titulo: str


def extraer_indice(path, config):
    doc = abrir(path)
    result = {}
    for p in doc.element.body.xpath('.//w:p'):
        s = texto(p).strip()
        m = TITULO.match(s) or re.match(r'^0*(\d+)[.)]\s+(.+)$', s)
        if m:
            n = int(m[1])
            title = s if TITULO.match(s) else f'Lección {n:02d}: {m[2]}'
            title = re.sub(r'\t+\d+\s*$', '', title)
            title = capitalizar(title, config)
            if n < 1 or n in result:
                raise ValueError(f'Índice ambiguo: número repetido o inválido {n}.')
            result[n] = Entrada(n, title)
    if not result:
        raise ValueError('El índice no contiene títulos reconocibles (Lección 01: Título).')
    return result


def resolver(doc, filename, entries):
    candidates = [texto(p).strip() for p in doc.element.body.xpath('./w:p') if texto(p).strip()][:6]
    first = candidates[0] if candidates else ''
    paragraphs = doc.element.body.xpath('./w:p')
    # El propio título no sirve de evidencia del contenido académico.
    text = normalizar(' '.join(texto(p) for p in paragraphs[1:]))
    tokens = set(text.split())
    scores = []
    for e in entries.values():
        subject = TITULO.sub(lambda m: m[2], e.titulo)
        keys = set(normalizar(subject).split()) - CONECTORES - {'leccion', 'lectura', 'tema', 'unidad'}
        cover = len(keys & tokens) / max(1, len(keys))
        sim = max((SequenceMatcher(None, normalizar(s), normalizar(e.titulo)).ratio()
                   for s in candidates[:2]), default=0)
        direct = any(normalizar(s) == normalizar(e.titulo) for s in candidates[:2])
        scores.append((1.0 if direct else .6 * sim + .4 * cover, cover, direct, e))
    scores.sort(key=lambda v: v[0], reverse=True)
    best = scores[0]
    m = re.search(r'(?:lecci[oó]n|lectura|tema|unidad)[ _-]*0*(\d+)', filename, re.I)
    file_num = int(m[1]) if m else None
    # Dos evidencias: título explícito + vocabulario, o nombre + cobertura única fuerte.
    margin = best[0] - scores[1][0] if len(scores) > 1 else 1
    evidence = (best[2] and best[1] >= .6) or (
        best[0] >= .7 and best[1] >= .75 and margin >= .12) or (
        file_num == best[3].numero and best[1] >= .9 and margin >= .2)
    if evidence and (file_num is None or file_num == best[3].numero):
        return best[3], []
    return None, [f'Correspondencia de título dudosa para {filename}; título conservado: {first!r}.']
