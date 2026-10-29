"""Configuración estricta: no aceptar valores silenciosamente inesperados."""
from copy import deepcopy
import json
from pathlib import Path
from urllib.parse import urlsplit

DEFAULT = {
    'formato': {'papel': 'A4', 'fuente': 'Arial', 'tamano_cuerpo': 11,
                'interlineado': 1.5, 'alineacion': 'justificada', 'margenes_cm': 2.54},
    'estructura': {'capitulos_en_pagina_nueva': True, 'abstract_independiente': True,
                   'resumen_independiente': True, 'referencias_en_pagina_nueva': True,
                   'incluir_subcapitulos_indice': False, 'programa_academico': ''},
    'titulos': {'capitalizacion_personalizada': True, 'respetar_indice_oficial': True,
               'siglas': ['PDP', 'PMBOK', 'ISO', 'IA', 'EVM', 'RCPSP'],
               'nombres_propios': ['OpenAI', 'Microsoft', 'Word', 'Python', 'Mermaid',
                                   'Graphviz', 'McKinsey', 'PowerPoint']},
    'figuras': {'convertir_mermaid': True, 'priorizar_orientacion_vertical': True,
               'relacion_horizontal_revisar': 1.6, 'tamano_texto_minimo_pt': 10,
               'dpi_png': 300, 'validacion_visual': True, 'max_nodos': 80},
    'correccion': {'ortografia_automatica_segura': True, 'gramatica_conservadora': True,
                  'corregir_redundancias_seguras': True, 'modificar_citas_textuales': False,
                  'reemplazos_seguros': {}},
    'salida': {'generar_zip': True, 'generar_rar_si_disponible': True,
               'generar_informe': True, 'guardar_vistas_previas': True,
               'incluir_pendientes_en_zip': False},
    'herramientas': {'soffice': '', 'dot': '', 'mmdc': '', 'word_com': True,
                     'language_tool_url': '', 'timeout_seg': 120},
}


def cargar(path: Path | None = None) -> dict:
    config = deepcopy(DEFAULT)
    if path is not None:
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        if not isinstance(data, dict):
            raise ValueError('config.json debe contener un objeto.')
        for section, values in data.items():
            if section not in config or not isinstance(values, dict):
                raise ValueError(f'Sección desconocida o inválida: {section}')
            for key, value in values.items():
                if key not in config[section]:
                    raise ValueError(f'Opción desconocida: {section}.{key}')
                expected = DEFAULT[section][key]
                if type(expected) is bool:
                    valid = type(value) is bool
                elif type(expected) in (int, float):
                    valid = type(value) in (int, float) and value > 0
                else:
                    valid = isinstance(value, type(expected))
                if not valid:
                    raise ValueError(f'Valor inválido: {section}.{key}')
                config[section][key] = value
    f = config['formato']
    if f['papel'] != 'A4' or f['alineacion'] != 'justificada':
        raise ValueError('Solo se admite papel A4 y alineación justificada.')
    if not (8 <= f['tamano_cuerpo'] <= 18 and 1 <= f['interlineado'] <= 2.5
            and 1 <= f['margenes_cm'] <= 4 and f['fuente'].strip()):
        raise ValueError('Fuente, tamaño, márgenes o interlineado fuera de rango.')
    g = config['figuras']
    if not (150 <= g['dpi_png'] <= 600 and 8 <= g['tamano_texto_minimo_pt'] <= 18
            and 1 <= g['relacion_horizontal_revisar'] <= 5
            and type(g['max_nodos']) is int and 1 <= g['max_nodos'] <= 500):
        raise ValueError('Parámetros de figuras fuera de rango.')
    if config['correccion']['modificar_citas_textuales']:
        raise ValueError('No se permite modificar citas textuales automáticamente.')
    if not config['salida']['generar_informe']:
        raise ValueError('El informe de auditoría es obligatorio.')
    for key in ['siglas', 'nombres_propios']:
        if any(not isinstance(v, str) or not v.strip() for v in config['titulos'][key]):
            raise ValueError(f'Lista inválida: titulos.{key}')
    replacements = config['correccion']['reemplazos_seguros']
    if any(not isinstance(k, str) or not isinstance(v, str) or not k.strip() or not v.strip()
           or any(c.isdigit() for c in k + v) for k, v in replacements.items()):
        raise ValueError('Reemplazos seguros: usar cadenas no vacías sin cifras.')
    url = config['herramientas']['language_tool_url']
    if url:
        p = urlsplit(url)
        if (p.scheme != 'http' or p.hostname not in ('127.0.0.1', 'localhost', '::1')
                or p.username or p.password or p.query or p.fragment or p.path not in ('', '/')):
            raise ValueError('LanguageTool debe ser un servidor HTTP local sin credenciales.')
    if not 5 <= config['herramientas']['timeout_seg'] <= 900:
        raise ValueError('timeout_seg debe estar entre 5 y 900.')
    return config
