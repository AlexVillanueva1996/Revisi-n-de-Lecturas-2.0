from collections import Counter
import json
from pathlib import Path


def guardar(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.with_suffix('.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    # JSON legible también en Markdown sin omitir incidencias ni inventar métricas.
    lines = [f'# {data.get("original", "Resumen global")}', '',
             f'Estado: **{data.get("estado", "Ver resultados individuales")}**', '',
             '```json', json.dumps(data,ensure_ascii=False,indent=2), '```', '']
    path.with_suffix('.md').write_text('\n'.join(lines),encoding='utf-8')


def global_informe(reports, folder, extra=None):
    data = {'lecturas_procesadas':len(reports),
            'estados':dict(Counter(r['estado'] for r in reports)),
            'lecturas':[{'original':r['original'], 'corregido':r.get('corregido'),
                          'estado':r['estado'], 'incidencias':r.get('incidencias',[])} for r in reports],
            **(extra or {})}
    guardar(folder/'Informe_Global', data)
    return data
