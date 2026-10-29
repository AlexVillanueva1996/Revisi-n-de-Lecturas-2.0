from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from .lector_documentos import parrafos
from .organizador_capitulos import analizar, localizar_indice


def aplicar(doc, config):
    f = config['formato']; s = config['estructura']
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(f['margenes_cm'])
    for name, size in [('Normal',f['tamano_cuerpo']),('Title',18),('Heading 1',14),('Heading 2',12)]:
        st = doc.styles[name]; st.font.name = f['fuente']; st.font.size = Pt(size)
        st.font.color.rgb = RGBColor(0,0,0)
        st.paragraph_format.line_spacing = f['interlineado']
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.widow_control = True
        if name != 'Normal': st.font.bold = True
    excluded = [p for b in localizar_indice(doc) for p in b]
    headings = {h.p._p:h for h in analizar(doc, excluded)}
    for p in parrafos(doc):
        # No tocar el formato de ecuaciones/campos/hipervínculos protegidos.
        protected = bool(p._p.xpath('.//w:hyperlink | .//w:fldSimple | .//w:fldChar | .//m:oMath | .//w:object'))
        h = headings.get(p._p)
        if h and h.clase in ('capitulo','abstract','resumen','referencias'):
            p.style = doc.styles['Heading 1']
        elif h and h.clase == 'subcapitulo':
            p.style = doc.styles['Heading 2']
        title = p.style.name in ('Title','Heading 1','Heading 2')
        pf = p.paragraph_format
        pf.line_spacing = f['interlineado']; pf.space_before = Pt(0); pf.space_after = Pt(6)
        pf.widow_control = True
        pf.keep_with_next = title
        pf.keep_together = title
        pf.left_indent = pf.right_indent = Cm(0)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if title else WD_ALIGN_PARAGRAPH.JUSTIFY
        if h:
            enabled = ((h.clase == 'capitulo' and s['capitulos_en_pagina_nueva']) or
                       (h.clase == 'abstract' and s['abstract_independiente']) or
                       (h.clase == 'resumen' and s['resumen_independiente']) or
                       (h.clase == 'referencias' and s['referencias_en_pagina_nueva']))
            if enabled:
                pf.page_break_before = True
                # Usar una sola propiedad de salto. Retirar únicamente saltos explícitos adyacentes.
                prev = p._p.getprevious()
                while prev is not None and prev.tag == qn('w:p'):
                    if prev.xpath('.//w:t | .//w:drawing | .//w:fldChar | .//w:object'):
                        break
                    for br in prev.xpath('.//w:br[@w:type="page"]'):
                        br.getparent().remove(br)
                    prev = prev.getprevious()
                for br in p._p.xpath('.//w:br[@w:type="page"]'):
                    br.getparent().remove(br)
        if not protected:
            for r in p.runs:
                # Mantener negritas, cursivas, super/subíndices y énfasis internos.
                r.font.name = f['fuente']
                r.font.size = Pt(18 if p.style.name=='Title' else 14 if p.style.name=='Heading 1'
                                 else 12 if p.style.name=='Heading 2' else f['tamano_cuerpo'])
                r.font.color.rgb = RGBColor(0,0,0)
                fonts = r._r.get_or_add_rPr().rFonts
                if fonts is not None:
                    for a in ('ascii','hAnsi','eastAsia','cs'): fonts.set(qn('w:'+a), f['fuente'])
    return {'fuente':f['fuente'], 'cuerpo_pt':f['tamano_cuerpo'], 'interlineado':f['interlineado'],
            'papel':'A4', 'alineacion':'justificada', 'campos_y_objetos':'formato interno preservado'}
