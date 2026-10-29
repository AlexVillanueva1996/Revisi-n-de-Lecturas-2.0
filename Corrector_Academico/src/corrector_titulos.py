import re
import unicodedata

CONECTORES = set('de del el la los las un una unos unas y e o u a al en con sin por para entre sobre desde hasta mediante hacia durante frente'.split())


def normalizar(s):
    return re.sub(r'\W+', ' ', ''.join(c for c in unicodedata.normalize('NFD', s.casefold())
                                      if unicodedata.category(c) != 'Mn')).strip()


def capitalizar(titulo, config):
    cfg = config['titulos']
    if not cfg['capitalizacion_personalizada']:
        return titulo
    especiales = {v.casefold(): v for v in cfg['siglas'] + cfg['nombres_propios']}
    first = True
    def word(m):
        nonlocal first
        w = m.group(); low = w.casefold()
        if low in especiales:
            result = especiales[low]
        elif not first and low in CONECTORES:
            result = low
        elif not w.isupper() and any(c.isupper() for c in w[1:]):
            result = w[0].upper() + w[1:]  # iPhone, McDonald, apellidos compuestos
        else:
            result = w[0].upper() + w[1:].lower()
        first = False
        return result
    # No contar numeración romana inicial como primera palabra del título.
    prefix = re.match(r'^((?:[IVXLCDM]+[.\s:–-]+|\d+(?:\.\d+)*[.\s:–-]+))', titulo)
    a = prefix.group() if prefix else ''
    return a + re.sub(r'[^\W\d_]+', word, titulo[len(a):], flags=re.UNICODE)
