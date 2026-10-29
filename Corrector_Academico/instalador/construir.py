"""Compilar un instalador Windows con NSIS, también desde Linux."""
import argparse
from pathlib import Path
import shutil
import subprocess
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED

BASE = Path(__file__).resolve().parents[1]


def archivos():
    for folder in ('src','tests','scripts','instalador'):
        for p in (BASE/folder).rglob('*'):
            if p.is_file() and '__pycache__' not in p.parts and p.suffix not in ('.pyc','.exe'):
                yield p
    for name in ('config.json','requirements.txt','requirements-opcionales.txt',
                 'README.md','VALIDACION.md','AGENTS.md','INSTALAR_TODO.bat','INICIAR_CORRECCION.bat'):
        yield BASE/name
    for folder in ('ENTRADAS/INDICE','ENTRADAS/LECTURAS','SALIDAS/WORD','SALIDAS/INFORMES',
                   'SALIDAS/VISTAS_PREVIAS','SALIDAS/COMPRIMIDO','TEMPORALES'):
        yield BASE/folder/'.gitkeep'


def construir(makensis, output):
    output = output.resolve(); output.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='corrector_payload_') as temp:
        stage = Path(temp)
        for p in archivos():
            name = p.relative_to(BASE)
            if name.as_posix()=='config.json': name=Path('config.predeterminada.json')
            dst=stage/name; dst.parent.mkdir(parents=True,exist_ok=True)
            data=p.read_bytes()
            if p.suffix in ('.bat','.ps1'):
                data=data.replace(b'\r\n',b'\n').replace(b'\n',b'\r\n')
                # Windows PowerShell 5.1 necesita BOM para texto UTF-8 con acentos.
                if p.suffix=='.ps1' and not data.startswith(b'\xef\xbb\xbf'): data=b'\xef\xbb\xbf'+data
            dst.write_bytes(data)
        subprocess.run([str(makensis),f'-DPAYLOAD={stage}',f'-DOUTPUT={output}',
                        str(BASE/'instalador/Corrector.nsi')],check=True)
    if output.read_bytes()[:2]!=b'MZ': raise ValueError('La salida no es un ejecutable Windows.')
    with ZipFile(output.with_suffix('.zip'),'w',ZIP_DEFLATED) as z:
        z.write(output,output.name)
        z.writestr('LEEME.txt','Ejecute Corrector_Academico_Instalador.exe en Windows 10/11 x64.\r\n'
                   'Necesita Internet durante la instalacion de dependencias.\r\n'
                   'Despues coloque el indice y las lecturas en ENTRADAS y use el acceso directo.\r\n'
                   'El instalador no se ha ejecutado en un Windows real durante la compilacion en Linux.\r\n')
    return output


if __name__=='__main__':
    p=argparse.ArgumentParser(); p.add_argument('--makensis',default='makensis'); p.add_argument('--salida',type=Path,required=True)
    args=p.parse_args(); print(construir(args.makensis,args.salida))
