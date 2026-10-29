"""CLI en español. python -m src.main --help"""
import argparse
from collections import Counter
from datetime import datetime, timezone
import json
import logging
import os
import re
from pathlib import Path
import shutil
import sys
import uuid

from .analizador_indice import extraer_indice, resolver
from .compresor import comprimir
from .configuracion import cargar
from .corrector_textos import Corrector
from .formateador_word import aplicar
from .generador_informes import guardar, global_informe
from .lector_documentos import Auditor, abrir, sha256
from .organizador_capitulos import organizar
from . import procesador_figuras, procesador_mermaid, procesador_tablas
from .validador_documentos import (herramienta, inspeccionar_pdf, renderizar,
                                   reparar_saltos, validar_estructura)

BASE = Path(__file__).resolve().parents[1]
CARPETAS = ['ENTRADAS/INDICE','ENTRADAS/LECTURAS','SALIDAS/WORD','SALIDAS/INFORMES',
            'SALIDAS/VISTAS_PREVIAS','SALIDAS/COMPRIMIDO','TEMPORALES']


def preparar(root):
    root.mkdir(parents=True, exist_ok=True)
    for path in CARPETAS:
        folder = root/path
        folder.mkdir(parents=True, exist_ok=True)
        if not folder.resolve().is_relative_to(root.resolve()):
            raise ValueError(f'Carpeta fuera de la raíz por enlace simbólico: {folder}')


def documentos(folder):
    def natural(path):
        return tuple((1,int(s)) if s.isdigit() else (0,s.casefold())
                     for s in re.split(r'(\d+)',path.relative_to(folder).as_posix()))
    return sorted((p for p in folder.rglob('*') if p.is_file() and p.suffix.lower()=='.docx'
                   and not p.name.startswith('~$')),key=natural)


def diagnostico(config):
    info = {'python':sys.version.split()[0], 'plataforma':sys.platform,
            'LibreOffice':herramienta(config,'soffice') or 'No disponible',
            'Graphviz':config['herramientas']['dot'] or shutil.which('dot') or 'No disponible',
            'Mermaid CLI':config['herramientas']['mmdc'] or shutil.which('mmdc') or 'Opcional, no disponible',
            'LanguageTool local':config['herramientas']['language_tool_url'] or 'No configurado; revisión lingüística pendiente',
            'Word COM':'Requiere Word y pywin32 en Windows' if os.name=='nt' else 'Solo Windows'}
    return info


def procesar_uno(path, number, entries, root, run, config, corrector):
    temp = root/'TEMPORALES'/run/f'lectura_{number:02d}'
    temp.mkdir(parents=True, exist_ok=True)
    report = {'original':str(path), 'corregido':None,'estado':'ERROR DE PROCESAMIENTO',
              'incidencias':[], 'paginas_originales':None,'paginas_finales':None}
    issues = report['incidencias']
    original_hash = sha256(path)
    report['sha256_original'] = original_hash
    try:
        doc = abrir(path)
        auditor = Auditor(doc)
        (temp/'contenido_original.json').write_text(json.dumps(auditor.original,ensure_ascii=False,indent=2),encoding='utf-8')
        entry, doubts = resolver(doc,path.stem,entries)
        issues.extend(doubts)
        report['titulo_oficial'] = entry.titulo if entry else None
        report['numero_oficial'] = entry.numero if entry else None
        if doc.element.body.xpath('.//w:ins | .//w:del | .//w:txbxContent | .//w:altChunk | .//w:sdt'):
            issues.append('Hay revisiones, cuadros, contenido incrustado o controles: revisión manual de estructura y orden de lectura.')
        for table in doc.tables:
            if any('flowchart ' in texto for texto in [p.text for row in table.rows for cell in row.cells for p in cell.paragraphs]):
                issues.append('Mermaid en tabla: conservado sin sustitución; requiere revisión manual.')
        # No permitir que un renderizador siga imágenes/plantillas externas del paquete.
        external = [r.reltype for part in doc.part.package.parts for r in part.rels.values()
                    if r.is_external and not r.reltype.endswith('/hyperlink')]
        field_instructions = ' '.join(doc.element.body.xpath('.//w:instrText/text()'))
        unsafe_fields = bool(re.search(r'\b(?:DDEAUTO|DDE|INCLUDETEXT|INCLUDEPICTURE|LINK)\b',field_instructions,re.I))
        render_allowed = not external and not unsafe_fields
        if external: issues.append('Renderizado omitido: relaciones externas de imágenes/plantillas; no se descargaron datos.')
        if unsafe_fields: issues.append('Renderizado omitido: campos con enlaces activos o automatización externa.')
        if render_allowed:
            try:
                pdf, engine = renderizar(path,temp/'original_pdf',config)
                import fitz
                with fitz.open(pdf) as original_pdf: report['paginas_originales'] = len(original_pdf)
            except Exception as exc:
                issues.append(f'Paginación original no verificada: {exc}')
        headings = organizar(doc,auditor,config,entry,issues)
        report['estructura'] = [{'clase':h.clase,'numero':h.numero,'titulo':h.p.text} for h in headings]
        report['correccion'] = corrector.revisar(doc,auditor,issues)
        report['formato'] = aplicar(doc,config)
        report['diagramas'] = procesador_mermaid.convertir(doc,auditor,config,temp,issues)
        # Formatear contenido nuevo antes de unir títulos con sus objetos.
        aplicar(doc,config)
        report['tablas'] = procesador_tablas.procesar(doc,auditor,config,issues)
        report['figuras'] = procesador_figuras.procesar(doc,auditor,config,issues)
        report['cambios'] = auditor.cambios
        candidate = temp/'candidato.docx'
        doc.save(candidate)
        integrity_errors = validar_estructura(candidate,auditor)
        report['integridad'] = {'verificada':not integrity_errors,'errores':integrity_errors,
                               'parrafos_originales':len(auditor.original),'parrafos_finales':len(auditor.esperados),
                               'partes_originales_verificadas':len(auditor.medios)+len(auditor.partes),
                               'tablas_datos_verificadas':len(auditor.tablas)}
        if sha256(path) != original_hash:
            integrity_errors.append('El archivo original cambió durante la ejecución.')
            report['integridad']['verificada'] = False
        if integrity_errors:
            # Reversión de toda transformación: no exportar candidato defectuoso.
            issues.extend(integrity_errors)
            restored = root/'SALIDAS/WORD'/run/'ERRORES'/f'Lectura {number:02d} - Original Conservado.docx'
            restored.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(path,restored)
            report['original_conservado'] = str(restored)
            raise ValueError('Integridad fallida. Candidato retenido solo en TEMPORALES; se conservó el original.')
        visual_ok = False
        if render_allowed and config['figuras']['validacion_visual']:
            try:
                pdf, engine = renderizar(candidate,temp/'final_pdf',config)
                preview_folder = root/'SALIDAS/VISTAS_PREVIAS'/run/f'Lectura {number:02d}'
                v = inspeccionar_pdf(pdf,doc,config,preview_folder if config['salida']['guardar_vistas_previas'] else None,
                                     report['diagramas'])
                if v['errores']:
                    repairs = reparar_saltos(doc)
                    if repairs:
                        report['reparaciones_formato'] = repairs
                        doc.save(candidate)
                        remaining = validar_estructura(candidate,auditor)
                        if remaining: raise ValueError('Integridad tras reparación: '+'; '.join(remaining))
                        pdf, engine = renderizar(candidate,temp/'final_pdf_reintento',config)
                        v = inspeccionar_pdf(pdf,doc,config,preview_folder if config['salida']['guardar_vistas_previas'] else None,
                                             report['diagramas'])
                report['validacion_visual'] = {**v,'motor':engine,'pdf':str(pdf)}
                report['paginas_finales'] = v['paginas']
                visual_ok = not v['errores']
                issues.extend(v['errores'])
                if config['salida']['guardar_vistas_previas']:
                    shutil.copy2(pdf,preview_folder/'Lectura_Validacion.pdf')
            except Exception as exc:
                issues.append(f'Validación visual no completada: {exc}')
        else:
            issues.append('Validación visual no ejecutada; no se puede aprobar el documento.')
        # Las incidencias de semántica/imagen/lingüística son requisitos pendientes,
        # no se transforman en una aprobación solo por abrir el archivo.
        report['estado'] = 'APROBADO' if visual_ok and not issues else 'REQUIERE REVISIÓN MANUAL'
        folder = root/'SALIDAS/WORD'/run
        if report['estado']=='REQUIERE REVISIÓN MANUAL': folder /= 'PENDIENTES'
        folder.mkdir(parents=True,exist_ok=True)
        destination = folder/f'Lectura {number:02d} - Revisada.docx'
        shutil.copy2(candidate,destination)
        report['corregido'] = str(destination)
        report['sha256_final'] = sha256(destination)
        report['original_sin_cambios'] = sha256(path)==original_hash
    except Exception as exc:
        logging.exception('Error procesando %s',path.name)
        report['estado'] = 'ERROR DE PROCESAMIENTO'
        issues.append(f'{type(exc).__name__}: {exc}')
    issues[:] = list(dict.fromkeys(issues))
    guardar(root/'SALIDAS/INFORMES'/run/f'Lectura_{number:02d}',report)
    return report


def ejecutar(root, config):
    preparar(root)
    indices = documentos(root/'ENTRADAS/INDICE')
    preferred = [p for p in indices if p.name.casefold()=='indice_curso.docx']
    if len(preferred)==1: indice = preferred[0]
    elif len(indices)==1: indice = indices[0]
    else: raise ValueError('Coloque un único índice en ENTRADAS/INDICE (preferentemente Indice_Curso.docx).')
    entries = extraer_indice(indice,config)
    inputs = documentos(root/'ENTRADAS/LECTURAS')
    if not inputs: raise ValueError('No hay lecturas .docx en ENTRADAS/LECTURAS.')
    if any(not p.resolve().is_relative_to(root.resolve()) for p in inputs+[indice]):
        raise ValueError('Las entradas deben permanecer dentro de la raíz del proyecto.')
    run = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')+'_'+uuid.uuid4().hex[:6]
    report_folder = root/'SALIDAS/INFORMES'/run
    report_folder.mkdir(parents=True,exist_ok=True)
    log_handler = logging.FileHandler(report_folder/'procesamiento.log',encoding='utf-8')
    logging.getLogger().addHandler(log_handler)
    print(f'Índice detectado: {indice.name}\nLecturas encontradas: {len(inputs)}\nEjecución: {run}',flush=True)
    corrector = Corrector(config)
    reports, seen = [], {}
    try:
        for n, path in enumerate(inputs,1):
            print(f'Procesando {n}/{len(inputs)}: {path.name}',flush=True)
            digest = sha256(path)
            if digest in seen:
                r = {'original':str(path),'estado':'REQUIERE REVISIÓN MANUAL','corregido':None,
                     'incidencias':[f'Duplicado exacto de {seen[digest]}; no se procesó dos veces.']}
                guardar(report_folder/f'Lectura_{n:02d}',r)
            else:
                seen[digest] = path.name
                r = procesar_uno(path,n,entries,root,run,config,corrector)
            reports.append(r)
            print(f'  {r["estado"]}',flush=True)
    finally:
        corrector.close()
        logging.getLogger().removeHandler(log_handler); log_handler.close()
    assignments = Counter(r.get('numero_oficial') for r in reports if r.get('numero_oficial'))
    global_issues = [f'Lectura oficial {n:02d} no encontrada.' for n in entries if not assignments[n]]
    for r in reports:
        n = r.get('numero_oficial')
        if n and assignments[n]>1:
            msg = f'Múltiples documentos corresponden a la lección {n}; revisar duplicados no idénticos.'
            r['incidencias'].append(msg); r['estado'] = 'REQUIERE REVISIÓN MANUAL'
            # Evitar una aparente aprobación si la ambigüedad se detecta al final.
            if r.get('corregido'):
                src = Path(r['corregido']); dst = root/'SALIDAS/WORD'/run/'PENDIENTES'/src.name
                dst.parent.mkdir(parents=True,exist_ok=True)
                if src!=dst: shutil.move(src,dst)
                r['corregido'] = str(dst)
            i = reports.index(r)+1; guardar(report_folder/f'Lectura_{i:02d}',r)
    zip_result = None
    if config['salida']['generar_zip']:
        folder = root/'SALIDAS/COMPRIMIDO'/run; folder.mkdir(parents=True,exist_ok=True)
        selected = [Path(r['corregido']) for r in reports
                    if r.get('corregido') and r['estado'] in ('APROBADO','APROBADO CON OBSERVACIONES')]
        zip_result = comprimir(selected,folder/'Lecturas_Corregidas.zip',config['salida']['generar_rar_si_disponible'])
        if config['salida']['incluir_pendientes_en_zip']:
            pending = [Path(r['corregido']) for r in reports
                       if r.get('corregido') and r['estado']=='REQUIERE REVISIÓN MANUAL']
            zip_result['pendientes'] = comprimir(pending,folder/'Lecturas_Pendientes_Revision.zip')
    result = global_informe(reports,report_folder,{'ejecucion':run,'indice':str(indice),
                            'incidencias_curso':global_issues,'comprimido':zip_result,
                            'herramientas':diagnostico(config)})
    print(json.dumps({'resumen':result['estados'],'comprimido':zip_result,'informes':str(report_folder)},ensure_ascii=False,indent=2))
    return result, reports


def main(argv=None):
    parser = argparse.ArgumentParser(description='Revisión académica conservadora y local de archivos Word.')
    parser.add_argument('--raiz',type=Path,default=BASE,help='Carpeta con ENTRADAS y SALIDAS.')
    parser.add_argument('--config',type=Path,default=BASE/'config.json')
    parser.add_argument('--diagnostico',action='store_true',help='Ver herramientas sin procesar documentos.')
    parser.add_argument('--crear-ejemplo',action='store_true',help='Crear ejemplos sintéticos en TEMPORALES/demo.')
    parser.add_argument('--abrir-salidas',action='store_true')
    args = parser.parse_args(argv)
    try:
        config = cargar(args.config)
        if args.diagnostico:
            print(json.dumps(diagnostico(config),ensure_ascii=False,indent=2)); return 0
        if args.crear_ejemplo:
            from .ejemplos import crear_demo
            target = args.raiz/'TEMPORALES/demo'
            crear_demo(target)
            print(f'Ejemplo sintético creado en {target}.\nEjecute: python -m src.main --raiz "{target}"')
            return 0
        result, reports = ejecutar(args.raiz.resolve(),config)
        if args.abrir_salidas and os.name=='nt': os.startfile(str((args.raiz/'SALIDAS').resolve()))
        if any(r['estado']=='ERROR DE PROCESAMIENTO' for r in reports): return 1
        if result['incidencias_curso'] or any(r['estado']=='REQUIERE REVISIÓN MANUAL' for r in reports): return 2
        return 0
    except Exception as exc:
        print(f'ERROR: {exc}',file=sys.stderr); return 1


if __name__=='__main__':
    raise SystemExit(main())
