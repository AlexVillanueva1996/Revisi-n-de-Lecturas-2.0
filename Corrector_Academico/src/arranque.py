"""Inicia y detiene únicamente el LanguageTool local de esta ejecución."""
import argparse
from contextlib import contextmanager
import json
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import time
from urllib.parse import urlencode
from urllib.request import build_opener, ProxyHandler, Request

from .configuracion import cargar
from .main import BASE, ejecutar


def leer_instalacion(path):
    info = json.loads(path.read_text(encoding='utf-8-sig'))
    if not isinstance(info, dict) or info.get('version') != 2:
        raise ValueError('Marcador de instalación inválido. Ejecute INSTALAR_TODO.bat.')
    for key in ('java','dot','soffice','python','languagetool'):
        if not isinstance(info.get(key), str) or not Path(info[key]).is_absolute() or not Path(info[key]).exists():
            raise ValueError(f'Falta una dependencia instalada: {key}. Ejecute INSTALAR_TODO.bat.')
    if not (Path(info['languagetool'])/'languagetool-server.jar').is_file():
        raise ValueError('Falta el servidor local LanguageTool.')
    return info


def solicitar(url, data=None):
    from urllib.parse import urlsplit
    parsed = urlsplit(url)
    if parsed.scheme != 'http' or parsed.hostname != '127.0.0.1' or parsed.username or parsed.password:
        raise ValueError('La verificación lingüística solo puede usar loopback local.')
    opener = build_opener(ProxyHandler({}))
    with opener.open(Request(url, data=data), timeout=60 if data is not None else 3) as response:
        return json.loads(response.read())


@contextmanager
def motor_local(info, folder, timeout=120):
    folder.mkdir(parents=True, exist_ok=True)
    with socket.socket() as sock:
        sock.bind(('127.0.0.1', 0)); port = sock.getsockname()[1]
    url = f'http://127.0.0.1:{port}'
    log = (folder/'languagetool.log').open('ab')
    process = None
    try:
        process = subprocess.Popen([info['java'], '-Xmx1024m', '-cp',
                                    str(Path(info['languagetool'])/'languagetool-server.jar'),
                                    'org.languagetool.server.HTTPServer', '--port', str(port)],
                                   cwd=info['languagetool'], stdout=log, stderr=subprocess.STDOUT,
                                   creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        start = time.monotonic(); last_notice = start
        while time.monotonic()-start < timeout:
            if process.poll() is not None:
                raise RuntimeError(f'LanguageTool se cerró ({process.returncode}). Consulte {folder}/languagetool.log')
            try:
                languages = solicitar(url+'/v2/languages')
                codes = {l.get('code') for l in languages}
                if {'es','en'}.issubset(codes): break
            except (OSError, ValueError): pass
            if time.monotonic()-last_notice > 15:
                print('Esperando el inicio de LanguageTool local...',flush=True); last_notice=time.monotonic()
            time.sleep(.3)
        else: raise RuntimeError('LanguageTool local no arrancó dentro del tiempo máximo.')
        yield url
    finally:
        if process is not None and process.poll() is None:
            process.terminate()
            try: process.wait(timeout=10)
            except subprocess.TimeoutExpired: process.kill(); process.wait(timeout=5)
        log.close()


def verificar_lenguas(url):
    result = {}
    for language, sample in [('en-US','This is an test.'),('es','Esto son una prueba.')]:
        response = solicitar(url+'/v2/check', urlencode({'language':language,'text':sample}).encode())
        matches = response.get('matches',[])
        if not matches: raise ValueError(f'El motor {language} no detectó el error de prueba.')
        result[language] = len(matches)
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description='Arranque con todas las dependencias locales.')
    parser.add_argument('--instalacion',type=Path,default=BASE/'instalacion.json')
    parser.add_argument('--config',type=Path,default=BASE/'config.json')
    parser.add_argument('--raiz',type=Path,default=BASE)
    parser.add_argument('--verificar-instalacion',action='store_true')
    args = parser.parse_args(argv)
    try:
        info = leer_instalacion(args.instalacion)
        config = cargar(args.config)
        config['herramientas'].update({'soffice':info['soffice'],'dot':info['dot']})
        # No sobrescribir config.json: solo añadir el servidor efímero en memoria.
        folder = args.raiz/'TEMPORALES/arranque'
        with motor_local(info,folder) as url:
            config['herramientas']['language_tool_url'] = url
            if args.verificar_instalacion:
                from .ejemplos import crear_demo
                languages = verificar_lenguas(url)
                with tempfile.TemporaryDirectory(prefix='verificacion_',dir=folder) as temp:
                    root = Path(temp); crear_demo(root)
                    result, reports = ejecutar(root,config)
                    (folder/'verificacion_documentos.json').write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')
                    if any(r['estado']=='ERROR DE PROCESAMIENTO' or not r.get('integridad',{}).get('verificada')
                           or r.get('validacion_visual',{}).get('errores') or not r.get('paginas_finales') for r in reports):
                        raise RuntimeError('Falló la prueba Word/PDF/integridad. Consulte instalacion.log.')
                    if not reports[0]['diagramas'][0].get('convertido'):
                        raise RuntimeError('No se pudo convertir el diagrama de prueba.')
                    if any(not r['correccion']['parrafos_revisados_linguisticamente']['es']
                           or not r['correccion']['parrafos_revisados_linguisticamente']['en-US'] for r in reports):
                        raise RuntimeError('El adaptador LanguageTool no revisó ambos idiomas.')
                (folder/'verificacion.json').write_text(json.dumps({'estado':'correcto', 'idiomas':languages,
                    'Word_PDF_Graphviz_integridad':True},ensure_ascii=False,indent=2),encoding='utf-8')
                print('Word, PDF, Graphviz y LanguageTool local (es/en-US): comprobados.',flush=True)
                return 0
            result, reports = ejecutar(args.raiz.resolve(),config)
            if any(r['estado']=='ERROR DE PROCESAMIENTO' for r in reports): return 1
            return 2 if result['incidencias_curso'] or any(r['estado']=='REQUIERE REVISIÓN MANUAL' for r in reports) else 0
    except Exception as exc:
        print(f'ERROR DE ARRANQUE: {exc}',file=sys.stderr)
        return 1


if __name__=='__main__': raise SystemExit(main())
