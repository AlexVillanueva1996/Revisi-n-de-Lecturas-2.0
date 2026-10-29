"""Documentos sintéticos; nunca se presentan como lecturas reales del usuario."""
from pathlib import Path
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from PIL import Image, ImageDraw

TITULOS = [
    'Lección 01: Planificación y Línea Base del Alcance',
    'Lección 02: Gestión de Riesgos e Interesados',
]
CAPITULOS = ['I. Fundamentos', 'II. Preparación', 'III. Desarrollo',
             'IV. Evaluación', 'V. Seguimiento', 'VI. Cierre']


def hipervinculo(p, label='Fuente académica', url='https://example.org/referencia'):
    link = OxmlElement('w:hyperlink')
    link.set(qn('r:id'),p.part.relate_to(url,RT.HYPERLINK,is_external=True))
    r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = label; r.append(t); link.append(r); p._p.append(link)


def lectura(path, numero=1, mermaid=True, image=True, table=True, indice_final=True):
    doc = Document()
    doc.add_paragraph(TITULOS[numero-1], 'Title')
    doc.sections[0].header.paragraphs[0].text = 'EJEMPLO SINTÉTICO — NO ES UNA LECTURA REAL'
    doc.add_paragraph('Abstract', 'Heading 1')
    doc.add_paragraph('This synthetic example describes scope planning, the baseline, risks and stakeholders. All academic content must be preserved.')
    doc.add_paragraph('Keywords: planning, scope, baseline, stakeholders.')
    doc.add_paragraph('Resumen Ejecutivo', 'Heading 1')
    doc.add_paragraph('La planificación y la línea base del alcance permiten organizar el proyecto. La gestión de riesgos e interesados requiere revisión académica.')
    for n, title in enumerate(CAPITULOS,1):
        doc.add_paragraph(title,'Heading 1')
        doc.add_paragraph(f'{n}.1 Conceptos de la Planificación','Heading 2')
        p = doc.add_paragraph('La planificación  del alcance ')
        p.add_run('conserva el significado').bold = True
        p.add_run(' y la línea base del proyecto.').italic = True
        doc.add_paragraph('Según Pérez (2020), “la planificación  conserva todas las evidencias” (p. 45).')
        doc.add_paragraph('El resultado es 12.5 y la fórmula es E = mc².')
        if n==1:
            p = doc.add_paragraph(); hipervinculo(p)
            if table:
                doc.add_paragraph('Tabla 4. datos DE la planificación')
                tb = doc.add_table(rows=3, cols=2)
                for row, values in zip(tb.rows,[('Concepto','Valor'),('Alcance','12.5'),('Riesgo','0.05')]):
                    for cell, value in zip(row.cells,values): cell.text = value
                doc.add_paragraph('Fuente: elaboración propia (2020).')
            if mermaid:
                doc.add_paragraph('```mermaid')
                for line in ['flowchart LR', 'A[Inicio] --> B[Planificación]',
                             'B --> C{¿Alcance aprobado?}', 'C -->|Sí| D[Resultados]',
                             'C -->|No| E[Revisión]']:
                    doc.add_paragraph(line)
                doc.add_paragraph('```')
                doc.add_paragraph('Fuente: elaboración propia; diagrama sintético.')
            if image:
                im = Image.new('RGB',(1000,400),'white'); draw = ImageDraw.Draw(im)
                draw.rectangle((40,40,960,360),outline='black',width=4)
                draw.text((100,150),'FIGURA SINTETICA: CONSERVAR IMAGEN ORIGINAL',fill='black')
                data = BytesIO(); im.save(data,format='PNG'); data.seek(0)
                doc.add_paragraph().add_run().add_picture(data)
                doc.add_paragraph('Fuente: imagen sintética original.')
    doc.add_paragraph('Referencias','Heading 1')
    doc.add_paragraph('Pérez, A. (2020). La planificación  académica. Editorial Ejemplo, pp. 45–49.')
    if indice_final:
        doc.add_paragraph('Índice')
        for title in CAPITULOS: doc.add_paragraph(title+'\t99')
    path.parent.mkdir(parents=True,exist_ok=True)
    doc.save(path)
    return path


def crear_demo(root):
    from .main import preparar
    preparar(root)
    index = root/'ENTRADAS/INDICE/Indice_Curso.docx'
    inputs = root/'ENTRADAS/LECTURAS'
    if index.exists() or list(inputs.glob('*.docx')):
        raise ValueError('La carpeta demo ya contiene documentos; no se sobrescriben.')
    doc = Document(); doc.add_paragraph('Índice del Curso — Ejemplo Sintético')
    for t in TITULOS: doc.add_paragraph(t)
    doc.save(index)
    lectura(inputs/'Lectura 01.docx')
    lectura(inputs/'Lectura 02.docx',numero=2,mermaid=False,image=False,table=False)
