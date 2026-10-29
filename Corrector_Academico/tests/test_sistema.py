import copy
from io import BytesIO
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

from src.analizador_indice import Entrada, extraer_indice, resolver
from src.compresor import comprimir
from src.configuracion import cargar, DEFAULT
from src.corrector_textos import Corrector
from src.corrector_titulos import capitalizar
from src.ejemplos import crear_demo, lectura, hipervinculo, TITULOS
from src.formateador_word import aplicar
from src.lector_documentos import Auditor, abrir, texto, sha256, tablas_firma, protegidos
from src.main import ejecutar
from src.organizador_capitulos import (analizar, localizar_indice, organizar,
                                       verificar_estructura)
from src.procesador_mermaid import (MermaidNoSeguro, parsear, dot_source, detectar_bloques,
                                    renderizar as render_diagrama, convertir)
from src.procesador_tablas import procesar as tablas
from src.validador_documentos import (renderizar, inspeccionar_pdf, herramienta,
                                      validar_estructura)
from unittest.mock import patch


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.config = cargar()


class Titulos(Base):
    def test_reglas(self):
        pairs = [
            ('fundamentos de la planificación de proyectos','Fundamentos de la Planificación de Proyectos'),
            ('EL DESARROLLO DEL PLAN PARA LA DIRECCIÓN DEL PROYECTO','El Desarrollo del Plan para la Dirección del Proyecto'),
            ('La Condena De Un Asesino Serial','La Condena de un Asesino Serial'),
            ('aplicación de PDP, PMBOK e ISO con IA y EVM','Aplicación de PDP, PMBOK e ISO con IA y EVM'),
            ('modelos de OpenAI y McKinsey','Modelos de OpenAI y McKinsey'),
            ('evaluación y análisis del cronograma','Evaluación y Análisis del Cronograma'),
            ('I. el marco conceptual','I. El Marco Conceptual'),
            ('1.1 la planificación','1.1 La Planificación'),
        ]
        for before, after in pairs:
            with self.subTest(before=before): self.assertEqual(capitalizar(before,self.config),after)

    def test_indice_otro_curso_en_tabla(self):
        doc = Document(); table=doc.add_table(rows=2,cols=1)
        table.cell(0,0).text = 'Unidad 01: Biología Molecular'
        table.cell(1,0).text = 'Unidad 02: Genética Aplicada'
        path=self.root/'indice.docx'; doc.save(path)
        entries=extraer_indice(path,self.config)
        self.assertEqual(entries[2].titulo,'Unidad 02: Genética Aplicada')

    def test_indice_duplicado(self):
        doc=Document()
        for _ in range(2): doc.add_paragraph('Lección 01: Título')
        path=self.root/'indice.docx'; doc.save(path)
        with self.assertRaises(ValueError): extraer_indice(path,self.config)

    def test_correspondencia_no_solo_archivo(self):
        doc=Document(); doc.add_paragraph('Astronomía'); doc.add_paragraph('Estrellas y galaxias lejanas.')
        e, issues=resolver(doc,'Lectura 01',{1:Entrada(1,TITULOS[0])})
        self.assertIsNone(e); self.assertTrue(issues)

    def test_titulo_correcto_pero_contenido_contradictorio(self):
        doc=Document(); doc.add_paragraph(TITULOS[0]); doc.add_paragraph('Estrellas y galaxias lejanas.')
        e,_=resolver(doc,'Lectura 01',{1:Entrada(1,TITULOS[0])})
        self.assertIsNone(e)

    def test_correspondencia_clara(self):
        doc=Document(); doc.add_paragraph(TITULOS[0]); doc.add_paragraph('Planificación y línea base del alcance.')
        e,issues=resolver(doc,'Lectura 01',{1:Entrada(1,TITULOS[0])})
        self.assertEqual(e.numero,1); self.assertFalse(issues)


class Configuracion(Base):
    def test_valores_invalidos(self):
        cases=[('formato','tamano_cuerpo',-1),('formato','interlineado',True),
               ('formato','papel','Carta'),('figuras','dpi_png',20),
               ('correccion','modificar_citas_textuales',True),
               ('herramientas','language_tool_url','https://externo.example'),
               ('herramientas','language_tool_url','http://localhost.evil'),
               ('formato','desconocido',2),('salida','generar_informe',False)]
        for section,key,value in cases:
            with self.subTest(key=key):
                p=self.root/'config.json'; p.write_text(json.dumps({section:{key:value}}))
                with self.assertRaises(ValueError): cargar(p)

    def test_servidor_local(self):
        p=self.root/'config.json'; p.write_text('{"herramientas":{"language_tool_url":"http://127.0.0.1:8081"}}')
        self.assertTrue(cargar(p)['herramientas']['language_tool_url'])


class Estructura(Base):
    def headings(self,names):
        doc=Document()
        for name in names: doc.add_paragraph(name,'Heading 1')
        return analizar(doc)

    def test_seis_capitulos(self):
        hs=self.headings(['Abstract','Resumen Ejecutivo','I. Uno','1.1 Tema','1.2 Otro',
                          'II. Dos','2.1 Tema','III. Tres','IV. Cuatro','V. Cinco','VI. Seis','Referencias'])
        self.assertEqual(verificar_estructura(hs),[])

    def test_orden_y_duplicados(self):
        for names in [['II. Dos','I. Uno'],['I. Uno','I. Uno']]:
            self.assertTrue(any('Capítulos' in s for s in verificar_estructura(self.headings(names))))

    def test_subcapitulos(self):
        for names in [['I. Uno','2.1 Error'],['I. Uno','1.2 Error'],['I. Uno','1.1.1 Sin Padre']]:
            self.assertTrue(any('Subcapítulo' in s for s in verificar_estructura(self.headings(names))))

    def test_no_confundir_tabla_parrafo(self):
        doc=Document(); doc.add_paragraph('Tabla 1. Datos'); doc.add_paragraph('1.1 Este es un párrafo normal terminado en punto.')
        self.assertEqual(analizar(doc),[])

    def test_indice_final_sincronizado_y_movido(self):
        path=lectura(self.root/'lectura.docx',mermaid=False,image=False,table=False)
        doc=abrir(path); auditor=Auditor(doc); issues=[]
        hs=organizar(doc,auditor,self.config,Entrada(1,TITULOS[0]),issues)
        blocks=localizar_indice(doc)
        self.assertEqual(len(blocks),1)
        self.assertEqual([texto(p._p) for p in blocks[0][1:]], [texto(h.p._p) for h in hs if h.clase=='capitulo'])
        self.assertLess(doc.element.body.index(blocks[0][0]._p),doc.element.body.index(next(h.p._p for h in hs if h.clase=='abstract')))
        self.assertFalse(auditor.validar())

    def test_reordenar_secciones_completas(self):
        doc=Document(); doc.add_paragraph(TITULOS[0]); doc.add_paragraph('Resumen Ejecutivo'); doc.add_paragraph('Resumen íntegro.')
        doc.add_paragraph('Abstract'); doc.add_paragraph('Complete abstract.'); doc.add_paragraph('I. Fundamentos'); doc.add_paragraph('Contenido íntegro.')
        doc.add_paragraph('Referencias'); doc.add_paragraph('Pérez (2020).')
        auditor=Auditor(doc); issues=[]
        organizar(doc,auditor,self.config,Entrada(1,TITULOS[0]),issues)
        hs=analizar(doc,[p for b in localizar_indice(doc) for p in b])
        self.assertEqual([h.clase for h in hs],['abstract','resumen','capitulo','referencias'])
        self.assertFalse(auditor.validar())


class IntegridadYWord(Base):
    def test_edicion_segmentada_conserva_runs(self):
        doc=Document(); p=doc.add_paragraph()
        p.add_run('Texto  ').bold=True; p.add_run('académico  ').italic=True; p.add_run('íntegro').font.superscript=True
        auditor=Auditor(doc); auditor.cambiar(p,'Texto académico íntegro','Prueba segura')
        self.assertEqual(p.text,'Texto académico íntegro')
        self.assertTrue(p.runs[0].bold); self.assertTrue(p.runs[1].italic); self.assertTrue(p.runs[2].font.superscript)
        self.assertFalse(auditor.validar())

    def test_deteccion_perdida_texto_y_tabla(self):
        doc=Document(); p=doc.add_paragraph('Contenido que no puede desaparecer.')
        table=doc.add_table(rows=1,cols=1); table.cell(0,0).text='12.50'
        auditor=Auditor(doc); p._p.getparent().remove(p._p); table.cell(0,0).text='0'
        errors=auditor.validar()
        self.assertTrue(any('párrafos' in s for s in errors)); self.assertTrue(any('tablas' in s for s in errors))

    def test_citas_cifras_y_referencias_no_modificadas(self):
        doc=Document(); doc.add_paragraph('Abstract'); doc.add_paragraph('Some  text.')
        doc.add_paragraph('Resumen Ejecutivo'); doc.add_paragraph('Texto  seguro.')
        protected=['“Una  cita textual”.',"'Otra  cita textual'.",'Según Pérez (2020), el  resultado es 12.5.','E = mc².']
        for s in protected: doc.add_paragraph(s)
        doc.add_paragraph('Referencias'); doc.add_paragraph('Pérez, A.  (2020).  Obra.')
        auditor=Auditor(doc); c=Corrector(self.config); result=c.revisar(doc,auditor,[]); c.close()
        texts=[p.text for p in doc.paragraphs]
        for s in protected+['Pérez, A.  (2020).  Obra.']: self.assertIn(s,texts)
        self.assertIn('Some text.',texts); self.assertIn('Texto seguro.',texts)
        self.assertEqual(result['parrafos_corregidos'],2)

    def test_hipervinculos_campos_formulas_guardar(self):
        doc=Document(); p=doc.add_paragraph(); hipervinculo(p)
        p=doc.add_paragraph(); fld=OxmlElement('w:fldSimple'); fld.set(qn('w:instr'),'PAGE'); p._p.append(fld)
        p=doc.add_paragraph(); eq=OxmlElement('m:oMath'); p._p.append(eq)
        auditor=Auditor(doc); aplicar(doc,self.config)
        path=self.root/'resultado.docx'; doc.save(path)
        self.assertFalse(validar_estructura(path,auditor))

    def test_no_perder_imagen(self):
        doc=Document(); data=BytesIO(); Image.new('RGB',(30,30),'blue').save(data,format='PNG'); data.seek(0)
        doc.add_paragraph().add_run().add_picture(data)
        auditor=Auditor(doc); aplicar(doc,self.config)
        path=self.root/'resultado.docx'; doc.save(path)
        self.assertFalse(validar_estructura(path,auditor))
        self.assertEqual(len(abrir(path).inline_shapes),1)

    def test_formato_y_saltos(self):
        doc=Document(); p=doc.add_paragraph('Cuerpo del documento.'); p.add_run(' Énfasis').bold=True
        for title in ['Abstract','Resumen Ejecutivo','I. Uno','II. Dos','Referencias']: doc.add_paragraph(title)
        aplicar(doc,self.config)
        self.assertEqual(p.runs[0].font.name,'Arial'); self.assertEqual(p.runs[0].font.size.pt,11)
        self.assertEqual(p.paragraph_format.line_spacing,1.5); self.assertEqual(int(p.alignment),3)
        self.assertTrue(all(h.p.paragraph_format.page_break_before for h in analizar(doc)))
        self.assertTrue(p.runs[1].bold)

    def test_tabla_editable_y_celdas_combinadas(self):
        doc=Document(); t=doc.add_table(rows=3,cols=3); t.cell(0,0).merge(t.cell(0,1)).text='Cabecera'
        t.cell(1,1).text='12.500'; auditor=Auditor(doc); before=tablas_firma(doc)
        aplicar(doc,self.config); tablas(doc,auditor,self.config,[])
        self.assertEqual(tablas_firma(doc),before); self.assertFalse(auditor.validar())
        widths=[int(g.w) for g in t._tbl.tblGrid.gridCol_lst]
        self.assertGreater(sum(widths),0)

    def test_docx_danado(self):
        path=self.root/'roto.docx'; path.write_bytes(b'No soy un archivo ZIP')
        with self.assertRaises(Exception): abrir(path)

    def test_saltos_y_tabs_en_edicion_segmentada(self):
        doc=Document(); p=doc.add_paragraph('I. fundamentos\t99')
        auditor=Auditor(doc); auditor.cambiar(p,'I. Fundamentos','Entrada de índice')
        self.assertEqual(texto(p._p),'I. Fundamentos'); self.assertFalse(auditor.validar())
        p=doc.add_paragraph('Texto  seguro.\nSegunda  línea.')
        auditor=Auditor(doc); auditor.cambiar(p,'Texto seguro.\nSegunda línea.','Espacios')
        self.assertEqual(p.text,'Texto seguro.\nSegunda línea.')

    def test_repeticiones_no_eliminan_palabras(self):
        doc=Document(); p=doc.add_paragraph('El alcance de de este proyecto necesita revisión.')
        auditor=Auditor(doc); corrector=Corrector(self.config)
        result=corrector.revisar(doc,auditor,[]); corrector.close()
        self.assertIn('de de',p.text); self.assertEqual(result['sugerencias'][0]['regla'],'REPETICION_ADYACENTE')


class Mermaid(Base):
    def test_horizontal_vertical_y_ramas(self):
        for orient in ('LR','TD','TB'):
            with self.subTest(orient=orient):
                g=parsear(f'flowchart {orient}\nA[Inicio] --> B{{¿Decisión?}}\nB -->|Sí| C[Fin]\nB -->|No| A')
                self.assertEqual(len(g.nodos),3); self.assertEqual(len(g.conexiones),3)
                self.assertEqual(g.nodos['B']['forma'],'diamond'); self.assertEqual(g.conexiones[1]['etiqueta'],'Sí')

    def test_subgrupos_etiquetas_largas_y_firma(self):
        code='flowchart LR\nsubgraph G[Proceso Académico]\nA[Texto muy largo con información académica que debe conservarse] --> B[Fin]\nend'
        graph=parsear(code); before=copy.deepcopy(graph.firma()); dot=dot_source(graph,'TB')
        self.assertEqual(graph.firma(),before); self.assertIn('cluster_G',dot)
        self.assertEqual(graph.nodos['A']['grupos'],['G'])
        graph=parsear('graph LR\nA --> B\nsubgraph G\nA[Inicio]\nend')
        self.assertEqual(graph.nodos['A']['grupos'],['G'])

    @unittest.skipUnless(shutil.which('dot'),'Graphviz no disponible')
    def test_subgrupos_decisiones_y_etiquetas_renderizadas(self):
        code='flowchart LR\nsubgraph G[Planificación Académica]\nA[Etiqueta académica larga que se distribuye en varias líneas sin eliminar palabras] --> B{¿Aprobado?}\nend\nB -->|Sí| C[Fin]\nB -->|No| A'
        g=parsear(code); result=render_diagrama(g,self.root,self.config,450,600)
        self.assertTrue(result['estructura_verificada']); self.assertGreaterEqual(result['fuente_final_pt'],10)
        svg=Path(result['svg']).read_text()
        self.assertIn('Planificación Académica',svg); self.assertIn('Sí',svg); self.assertIn('No',svg)

    def test_conexiones_encadenadas_y_etiquetas(self):
        g=parsear('graph LR; A[Uno] --> B[Dos] --> C[Tres]; C -- Regreso --> A')
        self.assertEqual(len(g.conexiones),3); self.assertEqual(g.conexiones[2]['etiqueta'],'Regreso')

    def test_rechazar_sintaxis_no_verificable(self):
        bad=['sequenceDiagram\nA->>B: hola','flowchart LR\nA --> B & C',
             'flowchart LR\nclick A "https://example.org"','flowchart LR\nA[Sin cierre',
             'flowchart LR\nsubgraph G\nA --> B', 'flowchart LR\nA[Uno]\nA[Dos]']
        for s in bad:
            with self.subTest(s=s), self.assertRaises(MermaidNoSeguro): parsear(s)

    def test_bloque_incompleto_se_conserva(self):
        doc=Document(); doc.add_paragraph('```mermaid'); doc.add_paragraph('flowchart LR'); doc.add_paragraph('A --> B')
        auditor=Auditor(doc); before=[p.text for p in doc.paragraphs]; issues=[]
        results=convertir(doc,auditor,self.config,self.root,issues)
        self.assertFalse(results[0]['convertido']); self.assertEqual(before,[p.text for p in doc.paragraphs])

    @unittest.skipUnless(shutil.which('dot'),'Graphviz no disponible')
    def test_render_horizontal_reorganizado_y_conexiones(self):
        graph=parsear('flowchart LR\nA[Inicio] --> B[Planificación] --> C[Ejecución] --> D[Evaluación] --> E[Resultados]')
        result=render_diagrama(graph,self.root,self.config,450,600)
        self.assertEqual(result['orientacion'],'TB')
        self.assertGreaterEqual(result['fuente_final_pt'],10)
        self.assertTrue(Path(result['png']).is_file()); self.assertTrue(result['estructura_verificada'])

    @unittest.skipUnless(shutil.which('dot'),'Graphviz no disponible')
    def test_complejo_ilegible_no_reemplazado(self):
        doc=Document(); doc.add_paragraph('```mermaid'); doc.add_paragraph('flowchart TB')
        for i in range(50): doc.add_paragraph(f'A{i}[Etiqueta muy larga {i}] --> A{i+1}[Etiqueta muy larga {i+1}]')
        doc.add_paragraph('```'); before=[p.text for p in doc.paragraphs]
        auditor=Auditor(doc); result=convertir(doc,auditor,self.config,self.root,[])
        self.assertFalse(result[0]['convertido']); self.assertEqual(before,[p.text for p in doc.paragraphs])

    @unittest.skipUnless(shutil.which('dot'),'Graphviz no disponible')
    def test_conversion_auditable_integridad(self):
        doc=Document(); doc.add_paragraph('```mermaid\nflowchart LR\nA[Inicio] --> B[Fin]\n')
        doc.add_paragraph('```'); auditor=Auditor(doc)
        result=convertir(doc,auditor,self.config,self.root,[])
        self.assertTrue(result[0]['convertido']); self.assertFalse(auditor.validar())
        self.assertEqual(len(doc.inline_shapes),1)
        self.assertIn('A[Inicio]',Path(result[0]['codigo_original']).read_text())


class Integracion(Base):
    def test_sin_entradas_error_claro_y_crea_carpetas(self):
        with self.assertRaisesRegex(ValueError,'índice'): ejecutar(self.root,self.config)
        self.assertTrue((self.root/'ENTRADAS/LECTURAS').is_dir())

    def test_campos_activos_no_renderizados(self):
        crear_demo(self.root)
        for f in (self.root/'ENTRADAS/LECTURAS').iterdir(): f.unlink()
        p=self.root/'ENTRADAS/LECTURAS/Lectura 01.docx'
        doc=Document(); doc.add_paragraph(TITULOS[0]); doc.add_paragraph('Planificación y línea base del alcance.')
        r=doc.add_paragraph().add_run(); instr=OxmlElement('w:instrText'); instr.text='DDEAUTO cmd'; r._r.append(instr)
        doc.save(p)
        with patch('src.main.renderizar',side_effect=AssertionError('Nunca debe invocarse')) as renderer:
            _,reports=ejecutar(self.root,self.config)
            self.assertFalse(renderer.called)
        self.assertEqual(reports[0]['estado'],'REQUIERE REVISIÓN MANUAL')
        self.assertTrue(any('campos' in s for s in reports[0]['incidencias']))

    @unittest.skipUnless(herramienta(cargar(),'soffice'),'LibreOffice no disponible')
    def test_desbordamiento_abstract_conserva_todo(self):
        p=lectura(self.root/'largo.docx',mermaid=False,image=False,table=False)
        doc=abrir(p)
        summary=next(q for q in doc.paragraphs if q.text=='Resumen Ejecutivo')
        for _ in range(40):
            q=doc.add_paragraph('Complete academic abstract content remains intact and legible. '*8)
            summary._p.addprevious(q._p)
        auditor=Auditor(doc); issues=[]
        organizar(doc,auditor,self.config,Entrada(1,TITULOS[0]),issues); aplicar(doc,self.config)
        final=self.root/'final.docx'; doc.save(final)
        self.assertFalse(validar_estructura(final,auditor))
        pdf,_=renderizar(final,self.root/'pdf',self.config)
        result=inspeccionar_pdf(pdf,doc,self.config)
        self.assertEqual(result['errores'],[])
        starts=result['inicios']
        self.assertEqual(starts[0]['pagina'],2)
        self.assertGreater(starts[1]['pagina'],3)
        self.assertGreater(starts[2]['pagina'],starts[1]['pagina'])

    def test_zip_solo_archivos_seleccionados(self):
        a=lectura(self.root/'Lectura 01 - Revisada.docx',image=False,table=False,mermaid=False)
        unused=lectura(self.root/'viejo.docx',image=False,table=False,mermaid=False)
        path=self.root/'Lecturas_Corregidas.zip'; outcome=comprimir([a],path)
        with ZipFile(path) as z:
            self.assertEqual(z.namelist(),[a.name]); self.assertIsNone(z.testzip())
        self.assertEqual(outcome['documentos'],1)

    def test_documentos_incompletos_duplicados_y_danados(self):
        crear_demo(self.root)
        files=self.root/'ENTRADAS/LECTURAS'
        for p in files.iterdir(): p.unlink()
        doc=Document(); doc.add_paragraph('Documento incompleto'); doc.add_paragraph('No inventar capítulos.')
        doc.save(files/'Lectura 01.docx'); shutil.copy2(files/'Lectura 01.docx',files/'Lectura 02.docx')
        (files/'Lectura 03.docx').write_bytes(b'corrupto')
        self.config['figuras']['validacion_visual']=False
        # Evitar renderizar el original en esta prueba de tratamiento de errores.
        self.config['herramientas']['soffice']='herramienta_inexistente'
        summary,reports=ejecutar(self.root,self.config)
        self.assertEqual(len(reports),3)
        self.assertEqual(reports[0]['estado'],'REQUIERE REVISIÓN MANUAL')
        self.assertIn('Duplicado',reports[1]['incidencias'][0])
        self.assertEqual(reports[2]['estado'],'ERROR DE PROCESAMIENTO')
        self.assertTrue(summary['incidencias_curso'])
        self.assertIsNotNone(reports[0]['corregido'])
        self.assertTrue(Path(reports[0]['corregido']).is_file())

    @unittest.skipUnless(herramienta(cargar(),'soffice'),'LibreOffice no disponible')
    def test_flujo_completo_paginacion_y_originales(self):
        crear_demo(self.root)
        original={p:sha256(p) for p in (self.root/'ENTRADAS').rglob('*.docx')}
        self.config['salida']['incluir_pendientes_en_zip']=True
        summary,reports=ejecutar(self.root,self.config)
        self.assertEqual(len(reports),2)
        for p,digest in original.items(): self.assertEqual(sha256(p),digest)
        for r in reports:
            self.assertNotEqual(r['estado'],'ERROR DE PROCESAMIENTO',r['incidencias'])
            self.assertTrue(r['integridad']['verificada'],r['integridad'])
            self.assertIsNotNone(r['paginas_originales'])
            self.assertEqual(r['validacion_visual']['errores'],[],r['validacion_visual'])
            starts=r['validacion_visual']['inicios']
            self.assertEqual([s['pagina'] for s in starts[:3]],[2,3,4])
            actual=[s['pagina'] for s in starts]
            self.assertTrue(all(b>a for a,b in zip(actual,actual[1:])))
            self.assertEqual(r['estado'],'REQUIERE REVISIÓN MANUAL')  # LanguageTool no configurado
            self.assertTrue(Path(r['corregido']).is_file())
        self.assertTrue(reports[0]['diagramas'][0]['convertido'])
        self.assertIn('pagina_pdf',reports[0]['diagramas'][0])
        zip_path=Path(summary['comprimido']['pendientes']['zip'])
        with ZipFile(zip_path) as z: self.assertEqual(len(z.namelist()),2)
        with ZipFile(summary['comprimido']['zip']) as z: self.assertEqual(len(z.namelist()),0)
        # Repetir no reutiliza ZIP, informes ni documentos de la ejecución anterior.
        previous=reports[0]['corregido']
        summary2,reports2=ejecutar(self.root,self.config)
        self.assertNotEqual(summary2['ejecucion'],summary['ejecucion'])
        self.assertNotEqual(reports2[0]['corregido'],previous)
        self.assertTrue(Path(previous).is_file())


if __name__=='__main__': unittest.main()
