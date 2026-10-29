"""Subconjunto estricto de flowchart Mermaid, con fallback sin pérdida.

Graphviz permite renderizado offline. No interpretar sintaxis desconocida ni
ejecutar click, scripts, includes o directivas de Mermaid.
"""
from dataclasses import dataclass, field, asdict
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import textwrap
from xml.etree import ElementTree as ET

from docx.shared import Inches
from PIL import Image
from .lector_documentos import texto


@dataclass
class Bloque:
    parrafos: list
    codigo: str


@dataclass
class Grafo:
    orientacion: str
    nodos: dict = field(default_factory=dict)
    conexiones: list = field(default_factory=list)
    grupos: dict = field(default_factory=dict)

    def firma(self):
        return {'nodos':self.nodos, 'conexiones':self.conexiones, 'grupos':self.grupos}


class MermaidNoSeguro(ValueError):
    pass


def detectar_bloques(doc):
    blocks, ps = [], doc.paragraphs
    i = 0
    while i < len(ps):
        s = texto(ps[i]._p).strip()
        if s.startswith('```mermaid'):
            start, lines = i, []
            if s != '```mermaid':
                lines.append(s[len('```mermaid'):].strip())
            i += 1
            while i < len(ps) and texto(ps[i]._p).strip() != '```':
                lines.append(texto(ps[i]._p)); i += 1
            if i < len(ps):
                blocks.append(Bloque(ps[start:i+1], '\n'.join(lines)))
            else:
                blocks.append(Bloque(ps[start:i], 'BLOQUE_SIN_CIERRE\n'+'\n'.join(lines)))
        elif re.match(r'^(flowchart|graph)\s+(LR|RL|TB|TD|BT)\b', s):
            start, lines = i, [texto(ps[i]._p)]
            # Un bloque sin fences requiere párrafos contiguos de sintaxis inequívoca.
            while i+1 < len(ps):
                following = texto(ps[i+1]._p).strip()
                if not following or not re.match(r'^(?:subgraph\b|end$|direction\b|[A-Za-z_]\w*\s*(?:\[|\{|\(|--|==|-.))', following):
                    break
                i += 1; lines.append(following)
            blocks.append(Bloque(ps[start:i+1], '\n'.join(lines)))
        i += 1
    return blocks


def statements(code):
    # Separar ; / saltos de línea solo fuera de etiquetas.
    out, buf, stack, quote = [], '', [], False
    for c in code:
        if c == '"': quote = not quote
        if not quote:
            if c in '[{(': stack.append(c)
            elif c in ']})':
                if not stack or {']':'[', '}':'{', ')':'('}[c] != stack[-1]:
                    raise MermaidNoSeguro('Delimitadores no balanceados.')
                stack.pop()
        if c in ';\n' and not stack and not quote:
            if buf.strip(): out.append(buf.strip())
            buf = ''
        else: buf += c
    if quote or stack: raise MermaidNoSeguro('Etiqueta sin cierre.')
    if buf.strip(): out.append(buf.strip())
    return out


def parsear(code):
    if len(code) > 100_000:
        raise MermaidNoSeguro('Diagrama excede el límite de tamaño.')
    lines = statements(code.strip())
    if not lines: raise MermaidNoSeguro('Diagrama vacío.')
    m = re.fullmatch(r'(?:flowchart|graph)\s+(LR|RL|TB|TD|BT)', lines.pop(0))
    if not m: raise MermaidNoSeguro('Solo flowchart/graph con orientación explícita.')
    graph = Grafo(m[1]); groups = []
    def node(s):
        m = re.match(r'([A-Za-z_][A-Za-z_0-9]*)', s)
        if not m: raise MermaidNoSeguro(f'Nodo no reconocido: {s[:60]}')
        ident, rest = m[1], s[m.end():].lstrip()
        label, shape, explicit = ident, 'box', False
        for a, b, sh in [('((', '))', 'ellipse'), ('[', ']', 'box'), ('{', '}', 'diamond'), ('(', ')', 'ellipse')]:
            if rest.startswith(a):
                quote, end = False, None
                for i in range(len(a), len(rest)):
                    if rest[i] == '"': quote = not quote
                    if not quote and rest[i:].startswith(b): end = i; break
                if end is None: raise MermaidNoSeguro('Etiqueta no balanceada.')
                label = rest[len(a):end].strip()
                if label.startswith('"') and label.endswith('"'): label = label[1:-1]
                label = re.sub(r'<br\s*/?>', '\n', label)
                if any(c in label for c in '<>{}[]'):
                    raise MermaidNoSeguro('Etiqueta con sintaxis HTML o anidada no compatible.')
                rest, shape, explicit = rest[end+len(b):].lstrip(), sh, True
                break
        old = graph.nodos.get(ident)
        if old and explicit and old['explicito'] and (old['etiqueta'], old['forma']) != (label, shape):
            raise MermaidNoSeguro('Nodo redefinido con etiqueta contradictoria.')
        if old and explicit and groups and old['grupos'] and old['grupos'] != groups:
            raise MermaidNoSeguro('Nodo declarado en subgrupos contradictorios.')
        membership = list(groups) if groups and (not old or not old['grupos']) else (old['grupos'] if old else [])
        if not old or explicit:
            graph.nodos[ident] = {'etiqueta':label, 'forma':shape, 'explicito':explicit,
                                     'grupos':membership}
        elif groups and not old['grupos']:
            old['grupos'] = membership
        return ident, rest
    for line in lines:
        if line.startswith('%%'):
            if line.startswith('%%{'): raise MermaidNoSeguro('Directivas no permitidas.')
            continue
        if line.startswith('subgraph '):
            label = line[9:].strip()
            m = re.fullmatch(r'([A-Za-z_]\w*)\s*\[(.+)\]', label)
            ident, label = (m[1], m[2].strip('"')) if m else (label, label)
            if not re.fullmatch(r'[A-Za-z_]\w*', ident) or ident in graph.grupos:
                raise MermaidNoSeguro('Subgrupo ambiguo o duplicado.')
            graph.grupos[ident] = {'etiqueta':label, 'padre':groups[-1] if groups else None}
            groups.append(ident); continue
        if line == 'end':
            if not groups: raise MermaidNoSeguro('end sin subgrupo.')
            groups.pop(); continue
        if re.fullmatch(r'direction (TB|TD|LR|RL|BT)', line):
            continue  # Restricción de diseño, no modifica relaciones.
        if re.match(r'^(?:click|class|classDef|style|linkStyle|flowchart|graph)\b', line):
            raise MermaidNoSeguro('Directiva fuera del subconjunto verificable.')
        source, rest = node(line)
        while rest:
            m = re.match(r'(-->|---|==>|-\.->)(?:\|([^|]+)\|)?\s*', rest)
            if m:
                arrow, label, rest = m[1], m[2] or '', rest[m.end():]
            else:
                m = re.match(r'--\s+([^|]+?)\s+-->\s*', rest)
                if not m: raise MermaidNoSeguro(f'Conexión no compatible: {rest[:60]}')
                arrow, label, rest = '-->', m[1], rest[m.end():]
            target, rest = node(rest)
            graph.conexiones.append({'origen':source, 'destino':target, 'etiqueta':label,
                                       'direccion':arrow != '---', 'estilo':arrow})
            source = target
    if groups: raise MermaidNoSeguro('Subgrupo sin cierre.')
    if not graph.nodos: raise MermaidNoSeguro('No se detectaron nodos.')
    return graph


def dot_source(graph, direction):
    q = lambda s: json.dumps(s, ensure_ascii=False)
    lines = ['digraph Diagrama {', f'graph [rankdir={direction}, bgcolor="white", pad=0.12, nodesep=0.25, ranksep=0.38];',
             'node [fontname="Arial", fontsize=11, style="rounded,filled", fillcolor="#eef3fa", margin="0.12,0.08"];',
             'edge [fontname="Arial", fontsize=11, arrowsize=0.7];']
    def emit(parent):
        for ident, group in graph.grupos.items():
            if group['padre'] == parent:
                lines.extend([f'subgraph {q("cluster_"+ident)} {{', f'label={q(group["etiqueta"])}; fontname="Arial"; fontsize=11;'])
                emit(ident); lines.append('}')
        for ident, node in graph.nodos.items():
            if (node['grupos'][-1] if node['grupos'] else None) == parent:
                label = '\n'.join(textwrap.fill(line, width=27, break_long_words=False, break_on_hyphens=False)
                                  for line in node['etiqueta'].split('\n'))
                lines.append(f'{q(ident)} [label={q(label)}, shape={node["forma"]}];')
    emit(None)
    for e in graph.conexiones:
        style = 'dashed' if e['estilo'] == '-.->' else 'bold' if e['estilo'] == '==>' else 'solid'
        lines.append(f'{q(e["origen"])} -> {q(e["destino"])} [label={q(e["etiqueta"])}, '
                     f'dir={"forward" if e["direccion"] else "none"}, style={style}];')
    return '\n'.join(lines + ['}'])


def renderizar(graph, folder, config, width_pt, height_pt):
    tool = config['herramientas']['dot'] or shutil.which('dot')
    if not tool: raise MermaidNoSeguro('Falta Graphviz (dot). Código original conservado.')
    fg = config['figuras']; original = 'TB' if graph.orientacion == 'TD' else graph.orientacion
    variants = [original]
    if fg['priorizar_orientacion_vertical'] and original in ('LR','RL'):
        variants.append('TB')
    results = []
    for orientation in variants:
        dot = dot_source(graph, orientation)
        stem = folder / orientation
        stem.with_suffix('.dot').write_text(dot, encoding='utf-8')
        svg = stem.with_suffix('.svg')
        subprocess.run([str(tool), '-Tsvg', str(stem.with_suffix('.dot')), '-o', str(svg)],
                       check=True, capture_output=True, timeout=config['herramientas']['timeout_seg'])
        root = ET.parse(svg).getroot()
        width = float(root.attrib['width'].rstrip('pt')); height = float(root.attrib['height'].rstrip('pt'))
        scale = min(1, width_pt / width, height_pt / height)
        sizes = [float(t.attrib.get('font-size', '11')) for t in root.iter() if t.tag.endswith('}text')]
        font = min(sizes, default=0)*scale
        # Verificar los IDs de los elementos realmente producidos por Graphviz.
        nodes, edges = [], []
        ns = {'s':'http://www.w3.org/2000/svg'}
        for g in root.findall('.//s:g', ns):
            title = g.find('s:title', ns)
            if title is not None:
                if g.attrib.get('class') == 'node': nodes.append(title.text)
                elif g.attrib.get('class') == 'edge': edges.append(title.text)
        from collections import Counter
        if set(nodes) != set(graph.nodos) or Counter(edges) != Counter(f'{e["origen"]}->{e["destino"]}' for e in graph.conexiones):
            raise MermaidNoSeguro('El SVG no conserva todos los nodos y conexiones.')
        results.append({'orientacion':orientation,'ancho_pt':width,'alto_pt':height,
                        'escala':scale,'fuente_final_pt':font,'svg':str(svg),'dot':str(stem.with_suffix('.dot'))})
    acceptable = [r for r in results if r['fuente_final_pt'] >= fg['tamano_texto_minimo_pt']]
    if not acceptable:
        raise MermaidNoSeguro('Diagrama ilegible al tamaño final; se conserva el código para revisión manual.')
    best = dict(max(acceptable, key=lambda r: (r['fuente_final_pt'],
              r['orientacion']=='TB' and r['ancho_pt']/r['alto_pt'] <= fg['relacion_horizontal_revisar'])))
    png = Path(best['dot']).with_suffix('.png')
    subprocess.run([str(tool), f'-Gdpi={fg["dpi_png"]}', '-Tpng', best['dot'], '-o', str(png)],
                   check=True, capture_output=True, timeout=config['herramientas']['timeout_seg'])
    with Image.open(png) as im: im.verify()
    best.update({'png':str(png), 'variantes':results, 'ancho_in':best['ancho_pt']*best['escala']/72,
                 'alto_in':best['alto_pt']*best['escala']/72, 'estructura_verificada':True})
    return best


def convertir(doc, auditor, config, temp, issues):
    results = []
    for i, block in enumerate(detectar_bloques(doc), 1):
        folder = temp / f'diagrama_{i:02d}'; folder.mkdir(parents=True, exist_ok=True)
        original = folder / 'original.mmd'; original.write_text(block.codigo, encoding='utf-8')
        item = {'codigo_original':str(original), 'convertido':False}
        try:
            if not config['figuras']['convertir_mermaid']:
                raise MermaidNoSeguro('Conversión Mermaid desactivada en configuración.')
            if any(p._p.xpath('.//w:hyperlink | .//w:fldChar | .//w:drawing | .//w:object') for p in block.parrafos):
                raise MermaidNoSeguro('El bloque contiene objetos protegidos.')
            graph = parsear(block.codigo)
            if len(graph.nodos) > config['figuras']['max_nodos']:
                raise MermaidNoSeguro('Demasiados nodos para conversión automática segura.')
            fingerprint = deepcopy(graph.firma())
            (folder/'estructura.json').write_text(json.dumps(fingerprint, ensure_ascii=False, indent=2), encoding='utf-8')
            # Reorganizar solo restricciones de layout; no editar etiquetas o aristas.
            sec = doc.sections[0]
            width = (sec.page_width-sec.left_margin-sec.right_margin)/12700
            height = (sec.page_height-sec.top_margin-sec.bottom_margin)/12700 - 90
            result = renderizar(graph, folder, config, width, height)
            if graph.firma() != fingerprint:
                raise MermaidNoSeguro('La reorganización cambió la estructura lógica.')
            p = doc.add_paragraph()
            block.parrafos[0]._p.addprevious(p._p)
            p.add_run().add_picture(result['png'], width=Inches(result['ancho_in']), height=Inches(result['alto_in']))
            auditor.registrar_nuevo(p, 'Figura Mermaid verificada')
            for old in block.parrafos:
                auditor.eliminar_diagrama(old, original)
            item.update(result); item['convertido'] = True
            item['vertical_reorganizado'] = graph.orientacion in ('LR','RL') and result['orientacion']=='TB'
            item['revision_visual_humana'] = 'pendiente: flechas, cruces y semántica'
            issues.append(f'Diagrama {i}: estructura y escala verificadas; revisar visualmente flechas y correspondencia académica.')
        except (ValueError, OSError, subprocess.SubprocessError, ET.ParseError) as exc:
            item['error'] = str(exc)
            issues.append(f'Diagrama {i} pendiente: {exc}')
        (folder/'auditoria.json').write_text(json.dumps(item, ensure_ascii=False, indent=2), encoding='utf-8')
        results.append(item)
    return results
