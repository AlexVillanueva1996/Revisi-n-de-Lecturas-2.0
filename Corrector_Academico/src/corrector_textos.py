"""Solo cambios deterministas; LanguageTool local aporta sugerencias auditadas."""
import re
from .lector_documentos import es_protegido, texto
from .organizador_capitulos import clasificar


class Corrector:
    def __init__(self, config):
        self.config = config
        self.clients = {}
        self.error = None
        url = config['herramientas']['language_tool_url']
        if url:
            try:
                import language_tool_python
                # remote_server evita la descarga y ejecución automática de Java.
                for language in ('es', 'en-US'):
                    self.clients[language] = language_tool_python.LanguageTool(language, remote_server=url)
            except Exception as exc:
                self.error = f'LanguageTool local no disponible: {type(exc).__name__}: {exc}'

    def close(self):
        for client in self.clients.values():
            client.close()

    def revisar(self, doc, auditor, issues):
        from .procesador_mermaid import detectar_bloques
        excluded = {p._p for b in detectar_bloques(doc) for p in b.parrafos}
        language, references = 'es', False
        count, suggestions, checked = 0, [], {'es':0, 'en-US':0}
        if self.error:
            issues.append(self.error)
        if len(self.clients) != 2:
            issues.append('Revisión lingüística completa no ejecutada: configure LanguageTool local (es/en-US).')
        for p in doc.paragraphs:
            h = clasificar(p)
            if h:
                if h.clase == 'referencias': references = True
                elif h.clase == 'abstract': language = 'en-US'
                elif h.clase in ('resumen','capitulo'): language = 'es'
                continue
            if references or p._p in excluded or es_protegido(p):
                continue
            s = texto(p._p)
            if not s.strip(): continue
            new = s
            cfg = self.config['correccion']
            if cfg['ortografia_automatica_segura']:
                new = re.sub(r' {2,}', ' ', new)
                for source, target in cfg['reemplazos_seguros'].items():
                    new = re.sub(r'(?<!\w)' + re.escape(source) + r'(?!\w)', lambda _: target, new)
            if new != s:
                auditor.cambiar(p, new, 'Corrección segura de espacios/reemplazos autorizados')
                count += 1
            if cfg['corregir_redundancias_seguras']:
                for repeat in re.finditer(r'\b(\w+)\s+\1\b',new,re.I):
                    suggestions.append({'idioma':language,'texto':new,'regla':'REPETICION_ADYACENTE',
                                        'mensaje':'Posible repetición: confirmar intención antes de eliminar palabras.',
                                        'offset':repeat.start(),'longitud':len(repeat.group()),'sugerencias':[repeat[1]]})
            if language in self.clients and cfg['gramatica_conservadora']:
                try:
                    matches = self.clients[language].check(new)
                    checked[language] += 1
                    for m in matches:
                        suggestions.append({'idioma':language, 'texto':new, 'regla':m.rule_id,
                                            'offset':m.offset, 'longitud':m.error_length,
                                            'mensaje':m.message, 'sugerencias':m.replacements})
                except Exception as exc:
                    issues.append(f'Falló la revisión lingüística {language}: {exc}')
        if suggestions:
            issues.append(f'{len(suggestions)} sugerencias lingüísticas pendientes; no se reescribieron argumentos.')
        if not cfg['gramatica_conservadora']:
            issues.append('Revisión gramatical desactivada; validación lingüística pendiente.')
        return {'parrafos_corregidos':count, 'sugerencias':suggestions,
                'parrafos_revisados_linguisticamente':checked,
                'idiomas_revisados':[lang for lang,n in checked.items() if n],
                'citas_y_referencias':'preservadas sin edición automática'}
