from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import shutil
import subprocess


def comprimir(paths, destination, rar=False):
    # Recibir la lista de ESTA ejecución; nunca arrastrar resultados antiguos.
    paths = list(paths)
    with ZipFile(destination,'w',compression=ZIP_DEFLATED) as z:
        for path in paths:
            if path.suffix.lower() != '.docx' or not path.is_file():
                raise ValueError('El comprimido solo admite los Word seleccionados.')
            z.write(path,path.name)
    with ZipFile(destination) as z:
        if z.testzip() is not None or sorted(z.namelist()) != sorted(p.name for p in paths):
            raise ValueError('Falló la comprobación del ZIP.')
    outcome = {'zip':str(destination), 'documentos':len(paths), 'rar':None}
    if rar and paths:
        tool = shutil.which('rar')
        if tool is None:
            for base in ('C:/Program Files/WinRAR','C:/Program Files (x86)/WinRAR'):
                p = Path(base)/'Rar.exe'
                if p.is_file(): tool = str(p); break
        if tool:
            target = destination.with_suffix('.rar')
            try:
                subprocess.run([tool,'a','-ep','-idq',str(target),*[str(p) for p in paths]],
                               check=True,capture_output=True,timeout=120)
                subprocess.run([tool,'t','-idq',str(target)],check=True,capture_output=True,timeout=120)
                outcome['rar'] = str(target)
            except (OSError,subprocess.SubprocessError) as exc:
                outcome['rar_error'] = str(exc)
        else: outcome['rar_estado'] = 'WinRAR/rar no disponible; ZIP generado.'
    return outcome
