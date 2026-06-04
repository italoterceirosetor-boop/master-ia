"""
Master IA — Servidor
Porta 5000. Uso: python master_server.py
"""

import os, io, json, re, hashlib, secrets
from datetime import datetime
from pathlib import Path
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

app = Flask(__name__)
CORS(app)

# Na nuvem usa /tmp, localmente usa pasta master_data
DATA_DIR = Path(os.environ.get("DATA_DIR", str(Path(__file__).parent / "master_data")))
DATA_DIR.mkdir(exist_ok=True, parents=True)
DOCS_DIR = DATA_DIR / "documentos"
DOCS_DIR.mkdir(exist_ok=True)

# ── Cores padrão ──────────────────────────────────────────────────────────────
C_AZUL      = colors.HexColor("#1a3a6e")
C_AZUL2     = colors.HexColor("#2d5fa3")
C_VERDE     = colors.HexColor("#00c2a8")
C_CINZA     = colors.HexColor("#8892a4")
C_BG_LIGHT  = colors.HexColor("#f5f7ff")
C_BG_ALT    = colors.HexColor("#eef2ff")
C_BORDER    = colors.HexColor("#d0d8f0")
C_WHITE     = colors.white

def now_str():
    return datetime.now().strftime("%d/%m/%Y às %H:%M")

def safe_name(s):
    s = re.sub(r'[^\w\s-]', '', s, flags=re.UNICODE)
    return re.sub(r'\s+', '_', s.strip())[:50]

def md_clean(line):
    """Converte markdown bold/italic para tags ReportLab"""
    line = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', line)
    line = re.sub(r'\*([^*]+)\*',     r'<i>\1</i>',  line)
    return line

def md_strip_links(line):
    """Remove markdown links [texto](url) deixando só o texto"""
    return re.sub(r'\[([^\]]+)\]\([^)]*\)', r'\1', line)

def md_strip(line):
    """Remove markdown bold/italic deixando texto puro"""
    line = re.sub(r'\*\*([^*]+)\*\*', r'\1', line)
    line = re.sub(r'\*([^*]+)\*',      r'\1', line)
    return line

def is_image_md(line):
    """Verifica se é uma linha de imagem markdown"""
    return bool(re.match(r'^!\[.*\]\(.*\)\s*$', line.strip()))


# ══════════════════════════════════════════════════════════════════════════════
#  GERADOR PDF — NÍVEL PROFISSIONAL
# ══════════════════════════════════════════════════════════════════════════════
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus.flowables import HRFlowable
import re as _re

def _pdf_parse_md(content):
    """Converte markdown em lista de blocos estruturados."""
    blocks = []
    lines = content.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Pula imagens markdown
        if _re.match(r'^!\[.*\]\(.*\)$', stripped):
            i += 1; continue

        # Tabela markdown
        if '|' in stripped and i+1 < len(lines) and _re.match(r'^\|[-| :]+\|', lines[i+1].strip()):
            headers = [c.strip() for c in stripped.split('|') if c.strip()]
            i += 2
            rows = []
            while i < len(lines) and '|' in lines[i]:
                r = [c.strip() for c in lines[i].split('|') if c.strip()]
                if r: rows.append(r)
                i += 1
            blocks.append({'type': 'table', 'headers': headers, 'rows': rows})
            continue

        # Heading levels
        m = _re.match(r'^(#{1,4})\s+(.+)', stripped)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            # Remove markdown inline
            text = _re.sub(r'\*\*(.+?)\*\*', r'\1', text)
            text = _re.sub(r'\*(.+?)\*', r'\1', text)
            blocks.append({'type': f'h{level}', 'text': text})
            i += 1; continue

        # Separador
        if _re.match(r'^---+$', stripped):
            blocks.append({'type': 'hr'})
            i += 1; continue

        # Lista com bullets
        if _re.match(r'^[-*•]\s+', stripped):
            items = []
            while i < len(lines) and _re.match(r'^[-*•]\s+', lines[i].strip()):
                items.append(lines[i].strip()[2:].strip())
                i += 1
            blocks.append({'type': 'bullets', 'items': items})
            continue

        # Lista numerada
        if _re.match(r'^\d+\.\s+', stripped):
            items = []
            while i < len(lines) and _re.match(r'^\d+\.\s+', lines[i].strip()):
                items.append(_re.sub(r'^\d+\.\s+', '', lines[i].strip()))
                i += 1
            blocks.append({'type': 'numbered', 'items': items})
            continue

        # Linha vazia
        if not stripped:
            blocks.append({'type': 'space'})
            i += 1; continue

        # Parágrafo normal
        # Acumula linhas consecutivas não especiais
        para_lines = []
        while i < len(lines):
            l = lines[i].strip()
            if (not l or l.startswith('#') or l.startswith('|') or
                _re.match(r'^[-*•]\s+', l) or _re.match(r'^\d+\.\s+', l) or
                _re.match(r'^---+$', l)):
                break
            para_lines.append(l)
            i += 1
        text = ' '.join(para_lines)
        if text.strip():
            blocks.append({'type': 'para', 'text': text})
        continue

    return blocks


def _md_to_rl(text):
    """Converte markdown inline para tags ReportLab."""
    text = _re.sub(r'&', '&amp;', text)
    text = _re.sub(r'<(?!/?[biu])', '&lt;', text)
    text = _re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    text = _re.sub(r'\*(.+?)\*',     r'<i>\1</i>', text)
    text = _re.sub(r'`(.+?)`',       r'<font name="Courier" size="9">\1</font>', text)
    text = _re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    return text


def gen_pdf(titulo, content, imagens=None):
    buf = io.BytesIO()

    # Paleta neutra profissional
    C_DARK   = colors.HexColor("#1a1f2e")
    C_MID    = colors.HexColor("#2d3748")
    C_ACCENT = colors.HexColor("#2b6cb0")
    C_LIGHT  = colors.HexColor("#ebf4ff")
    C_BG     = colors.HexColor("#f8fafc")
    C_BORDER = colors.HexColor("#cbd5e0")
    C_MUTED  = colors.HexColor("#718096")
    C_WHITE  = colors.white
    C_STRIPE = colors.HexColor("#f7fafc")

    doc = SimpleDocTemplate(buf, pagesize=A4,
        topMargin=2.5*cm, bottomMargin=2.5*cm,
        leftMargin=2.5*cm, rightMargin=2.5*cm,
        title=titulo)

    W = doc.width

    # ── Estilos ───────────────────────────────────────────────────────────────
    def sty(name, **kw):
        return ParagraphStyle(name, **kw)

    s_cover_title = sty('ct', fontSize=26, fontName='Helvetica-Bold',
        textColor=C_WHITE, alignment=TA_CENTER, leading=34, spaceAfter=8)
    s_cover_date  = sty('cd', fontSize=10, textColor=colors.HexColor("#a0aec0"),
        alignment=TA_CENTER)
    s_h1  = sty('h1', fontSize=14, fontName='Helvetica-Bold', textColor=C_DARK,
        spaceBefore=16, spaceAfter=6, leading=20,
        borderPad=(0,0,4,0))
    s_h2  = sty('h2', fontSize=12, fontName='Helvetica-Bold', textColor=C_ACCENT,
        spaceBefore=12, spaceAfter=4, leading=16)
    s_h3  = sty('h3', fontSize=10.5, fontName='Helvetica-Bold', textColor=C_MID,
        spaceBefore=8, spaceAfter=3, leading=14)
    s_h4  = sty('h4', fontSize=10, fontName='Helvetica-BoldOblique', textColor=C_MUTED,
        spaceBefore=6, spaceAfter=2, leading=13)
    s_body = sty('body', fontSize=10, leading=16, spaceAfter=6,
        alignment=TA_JUSTIFY, textColor=C_DARK,
        fontName='Helvetica')
    s_bul  = sty('bul', fontSize=10, leading=15, leftIndent=16,
        firstLineIndent=0, spaceAfter=3, textColor=C_DARK)
    s_num  = sty('num', fontSize=10, leading=15, leftIndent=20,
        firstLineIndent=-14, spaceAfter=3, textColor=C_DARK)
    s_foot = sty('foot', fontSize=8, textColor=C_MUTED, alignment=TA_CENTER)
    s_toc  = sty('toc', fontSize=10, textColor=C_ACCENT, leading=18,
        leftIndent=0, spaceAfter=2)
    s_toc2 = sty('toc2', fontSize=9.5, textColor=C_MID, leading=16,
        leftIndent=14, spaceAfter=1)

    story = []

    # ── CAPA ─────────────────────────────────────────────────────────────────
    if titulo:
        capa = Table([[Paragraph(titulo, s_cover_title)],
                      [Paragraph(now_str(), s_cover_date)]],
                     colWidths=[W])
        capa.setStyle(TableStyle([
            ('BACKGROUND',    (0,0), (-1,-1), C_DARK),
            ('TOPPADDING',    (0,0), (0,0),   28),
            ('BOTTOMPADDING', (0,0), (0,0),   6),
            ('TOPPADDING',    (0,1), (0,1),   4),
            ('BOTTOMPADDING', (0,1), (0,1),   24),
            ('LEFTPADDING',   (0,0), (-1,-1), 24),
            ('RIGHTPADDING',  (0,0), (-1,-1), 24),
            ('ROUNDEDCORNERS',(0,0), (-1,-1), [6,6,6,6]),
        ]))
        story.append(capa)
        story.append(Spacer(1, 0.4*cm))
        # Linha accent
        story.append(Table([['']], colWidths=[W], rowHeights=[4]))
        story[-1].setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),C_ACCENT),
                                        ('TOPPADDING',(0,0),(-1,-1),0),
                                        ('BOTTOMPADDING',(0,0),(-1,-1),0)]))
        story.append(Spacer(1, 0.6*cm))

    # ── SUMÁRIO AUTOMÁTICO ────────────────────────────────────────────────────
    blocks = _pdf_parse_md(content)
    headings = [(b['type'], b['text']) for b in blocks
                if b['type'] in ('h1','h2') and b.get('text','').strip()
                and b['text'].strip().lower() != (titulo or '').strip().lower()]

    if len(headings) >= 3:
        toc_title = Table([['ÍNDICE']], colWidths=[W])
        toc_title.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),C_BG),
            ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,-1),9),
            ('TEXTCOLOR',(0,0),(-1,-1),C_MUTED),
            ('TOPPADDING',(0,0),(-1,-1),6),
            ('BOTTOMPADDING',(0,0),(-1,-1),6),
            ('LEFTPADDING',(0,0),(-1,-1),10),
        ]))
        story.append(toc_title)
        for htype, htxt in headings:
            s = s_toc if htype == 'h1' else s_toc2
            prefix = '* ' if htype == 'h1' else '  - '
            story.append(Paragraph(prefix + htxt, s))
        story.append(Spacer(1, 0.5*cm))
        story.append(HRFlowable(width='100%', thickness=0.5,
                                 color=C_BORDER, spaceAfter=12))

    # ── CONTEÚDO ──────────────────────────────────────────────────────────────
    for block in blocks:
        btype = block.get('type')
        txt   = block.get('text','')

        # Pula título duplicado
        if btype in ('h1','h2','h3') and txt.strip().lower() == (titulo or '').strip().lower():
            continue

        if btype == 'h1':
            # Caixa de seção
            hbox = Table([[Paragraph(_md_to_rl(txt), s_h1)]], colWidths=[W])
            hbox.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,-1),C_LIGHT),
                ('LEFTPADDING',(0,0),(-1,-1),12),
                ('RIGHTPADDING',(0,0),(-1,-1),12),
                ('TOPPADDING',(0,0),(-1,-1),8),
                ('BOTTOMPADDING',(0,0),(-1,-1),8),
                ('LINEBEFORE',(0,0),(0,-1),4,C_ACCENT),
                ('LINEBELOW',(0,-1),(-1,-1),0.5,C_BORDER),
            ]))
            story.append(Spacer(1, 6))
            story.append(hbox)
            story.append(Spacer(1, 4))

        elif btype == 'h2':
            story.append(Spacer(1, 4))
            story.append(Paragraph(_md_to_rl(txt), s_h2))
            story.append(HRFlowable(width='40%', thickness=1.5,
                                     color=C_ACCENT, spaceAfter=4))

        elif btype == 'h3':
            story.append(Spacer(1, 2))
            story.append(Paragraph(_md_to_rl(txt), s_h3))

        elif btype == 'h4':
            story.append(Paragraph(_md_to_rl(txt), s_h4))

        elif btype == 'para':
            story.append(Paragraph(_md_to_rl(txt), s_body))

        elif btype == 'bullets':
            for item in block['items']:
                story.append(Paragraph(
                    f'<font color="#2b6cb0">-&gt;</font>  {_md_to_rl(item)}', s_bul))

        elif btype == 'numbered':
            for n, item in enumerate(block['items'], 1):
                story.append(Paragraph(
                    f'<b><font color="#2b6cb0">{n}.</font></b>  {_md_to_rl(item)}', s_num))

        elif btype == 'table':
            headers = block.get('headers', [])
            rows    = block.get('rows', [])
            if not headers: continue
            ncols = len(headers)
            all_rows = [headers] + rows
            # Calcula larguras proporcionais
            col_lens = [max((len(str(r[j])) if j < len(r) else 4)
                            for r in all_rows) for j in range(ncols)]
            total = sum(col_lens) or 1
            col_widths = [W * (cl/total) for cl in col_lens]

            is_wide = ncols > 5
            fs = 8 if is_wide else 9.5
            pad = 3 if is_wide else 6

            s_th = sty('th', fontSize=fs, fontName='Helvetica-Bold',
                       textColor=C_WHITE, alignment=TA_CENTER)
            s_td = sty('td', fontSize=fs, leading=fs+4,
                       textColor=C_DARK, alignment=TA_LEFT)

            tdata = [[Paragraph(str(h), s_th) for h in headers]]
            for ri, row in enumerate(rows):
                tdata.append([Paragraph(str(row[j]) if j < len(row) else '', s_td)
                               for j in range(ncols)])

            tbl = Table(tdata, colWidths=col_widths, repeatRows=1)
            ts = TableStyle([
                ('BACKGROUND',     (0,0), (-1,0), C_DARK),
                ('LINEBELOW',      (0,0), (-1,0), 2, C_ACCENT),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [C_WHITE, C_STRIPE]),
                ('GRID',           (0,0), (-1,-1), 0.3, C_BORDER),
                ('BOX',            (0,0), (-1,-1), 1,   C_BORDER),
                ('TOPPADDING',     (0,0), (-1,-1), pad),
                ('BOTTOMPADDING',  (0,0), (-1,-1), pad),
                ('LEFTPADDING',    (0,0), (-1,-1), pad+2),
                ('RIGHTPADDING',   (0,0), (-1,-1), pad+2),
                ('VALIGN',         (0,0), (-1,-1), 'MIDDLE'),
            ])
            tbl.setStyle(ts)
            story.append(Spacer(1, 6))
            story.append(tbl)
            story.append(Spacer(1, 8))

        elif btype == 'hr':
            story.append(HRFlowable(width='100%', thickness=0.5,
                                     color=C_BORDER, spaceBefore=6, spaceAfter=6))
        elif btype == 'space':
            story.append(Spacer(1, 4))

    # ── IMAGENS GERADAS ───────────────────────────────────────────────────────
    if imagens:
        from reportlab.platypus import Image as RLImage
        story.append(Spacer(1, 8))
        story.append(HRFlowable(width='100%', thickness=0.5, color=C_BORDER))
        story.append(Spacer(1, 6))
        for img_info in imagens:
            try:
                img_data = base64.b64decode(img_info['b64'])
                img_buf  = io.BytesIO(img_data)
                img_rl   = RLImage(img_buf, width=W, height=W*0.55)
                legenda  = img_info.get('legenda','')
                story.append(img_rl)
                if legenda:
                    story.append(Paragraph(legenda,
                        ParagraphStyle('leg', fontSize=8, textColor=C_MUTED,
                                       alignment=TA_CENTER, spaceAfter=6)))
                story.append(Spacer(1, 8))
            except Exception:
                pass

    # ── RODAPÉ ────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width='100%', thickness=0.5, color=C_BORDER))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"Gerado por Master IA · {now_str()}", s_foot))

    # Número de página via canvas
    def _add_page_num(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(C_MUTED)
        canvas.drawRightString(doc.pagesize[0] - 2*cm,
                               1.2*cm, f"Página {doc.page}")
        canvas.restoreState()

    doc.build(story, onLaterPages=_add_page_num, onFirstPage=_add_page_num)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
#  GERADOR WORD — NÍVEL PROFISSIONAL
# ══════════════════════════════════════════════════════════════════════════════
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.style import WD_STYLE_TYPE

def _word_set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)

def _word_add_border_bottom(paragraph, hex_color='2b6cb0', size=12):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:color'), hex_color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def gen_word(titulo, content):
    doc = Document()

    # Margens
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2.5)
        s.left_margin = s.right_margin = Cm(3)
        # Número de página no rodapé
        footer = s.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = fp.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    # ── Capa ────────────────────────────────────────────────────────────────
    if titulo:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(titulo)
        run.bold = True
        run.font.size = Pt(20)
        run.font.color.rgb = RGBColor(0x1a, 0x1f, 0x2e)

        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(now_str())
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

        _word_add_border_bottom(p2, '2b6cb0', 12)
        doc.add_paragraph()

    # ── Parser ──────────────────────────────────────────────────────────────
    blocks = _pdf_parse_md(content)

    for block in blocks:
        btype = block.get('type')
        txt   = _re.sub(r'\*\*(.+?)\*\*', r'\1', block.get('text',''))
        txt   = _re.sub(r'\*(.+?)\*', r'\1', txt)
        txt   = _re.sub(r'`(.+?)`', r'\1', txt)

        # Pula título duplicado
        if btype in ('h1','h2','h3') and txt.strip().lower() == (titulo or '').strip().lower():
            continue

        if btype == 'h1':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(txt.upper())
            run.bold = True
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0x1a, 0x1f, 0x2e)
            _word_add_border_bottom(p, '2b6cb0', 8)

        elif btype == 'h2':
            p = doc.add_paragraph()
            run = p.add_run(txt)
            run.bold = True
            run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(0x2b, 0x6c, 0xb0)

        elif btype == 'h3':
            p = doc.add_paragraph()
            run = p.add_run(txt)
            run.bold = True
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0x2d, 0x37, 0x48)

        elif btype == 'h4':
            p = doc.add_paragraph()
            run = p.add_run(txt)
            run.bold = True
            run.italic = True
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

        elif btype == 'para':
            raw = block.get('text','')
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            # Processa bold/italic inline
            parts = _re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)', raw)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    r = p.add_run(part[2:-2]); r.bold = True
                elif part.startswith('*') and part.endswith('*'):
                    r = p.add_run(part[1:-1]); r.italic = True
                elif part.startswith('`') and part.endswith('`'):
                    r = p.add_run(part[1:-1])
                    r.font.name = 'Courier New'; r.font.size = Pt(9)
                elif part:
                    p.add_run(part)
            for run in p.runs:
                run.font.size = Pt(10.5)

        elif btype == 'bullets':
            for item in block['items']:
                p = doc.add_paragraph(style='List Bullet')
                item_clean = _re.sub(r'\*\*(.+?)\*\*', r'\1', item)
                item_clean = _re.sub(r'\*(.+?)\*', r'\1', item_clean)
                run = p.add_run(item_clean)
                run.font.size = Pt(10.5)

        elif btype == 'numbered':
            for item in block['items']:
                p = doc.add_paragraph(style='List Number')
                item_clean = _re.sub(r'\*\*(.+?)\*\*', r'\1', item)
                run = p.add_run(item_clean)
                run.font.size = Pt(10.5)

        elif btype == 'table':
            headers = block.get('headers', [])
            rows    = block.get('rows', [])
            if not headers: continue
            ncols = len(headers)
            tbl = doc.add_table(rows=1+len(rows), cols=ncols)
            tbl.style = 'Table Grid'
            # Header
            for j, h in enumerate(headers):
                cell = tbl.rows[0].cells[j]
                cell.text = _re.sub(r'\*\*(.+?)\*\*', r'\1', h)
                run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(cell.text)
                run.bold = True
                run.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
                run.font.size = Pt(10)
                _word_set_cell_bg(cell, '1a1f2e')
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Rows
            for ri, row in enumerate(rows):
                bg = 'f7fafc' if ri % 2 == 0 else 'ffffff'
                for j in range(ncols):
                    cell = tbl.rows[ri+1].cells[j]
                    val = row[j] if j < len(row) else ''
                    cell.text = _re.sub(r'\*\*(.+?)\*\*', r'\1', str(val))
                    _word_set_cell_bg(cell, bg)
                    if cell.paragraphs[0].runs:
                        cell.paragraphs[0].runs[0].font.size = Pt(10)
            doc.add_paragraph()

        elif btype == 'hr':
            p = doc.add_paragraph()
            _word_add_border_bottom(p, 'cbd5e0', 4)

        elif btype == 'space':
            doc.add_paragraph()

    # ── Rodapé de conteúdo ──────────────────────────────────────────────────
    doc.add_paragraph()
    pf = doc.add_paragraph()
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _word_add_border_bottom(pf, 'cbd5e0', 4)
    rf = pf.add_run(f"Gerado por Master IA · {now_str()}")
    rf.font.size = Pt(8)
    rf.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
#  GERADOR EXCEL — NÍVEL PROFISSIONAL
# ══════════════════════════════════════════════════════════════════════════════
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, Rule
from openpyxl.utils import get_column_letter

def gen_excel(titulo, content):
    import re as _re2
    wb = openpyxl.Workbook()

    # ── Paleta ────────────────────────────────────────────────────────────────
    C_HEADER  = "1a1f2e"
    C_ACCENT  = "2b6cb0"
    C_ACCENT2 = "ebf4ff"
    C_STRIPE  = "f7fafc"
    C_WHITE   = "FFFFFF"
    C_MUTED   = "718096"
    C_GREEN   = "c6f6d5"
    C_RED     = "fed7d7"
    C_YELLOW  = "fefcbf"
    C_BORDER  = "cbd5e0"

    f_hdr   = Font(bold=True, color="FFFFFF", size=10, name='Calibri')
    f_title = Font(bold=True, color=C_HEADER, size=14, name='Calibri')
    f_sub   = Font(bold=True, color=C_ACCENT, size=11, name='Calibri')
    f_h3    = Font(bold=True, color="2d3748", size=10, name='Calibri')
    f_body  = Font(size=10, name='Calibri')
    f_num   = Font(size=10, name='Calibri')
    f_total = Font(bold=True, color="FFFFFF", size=10, name='Calibri')
    f_muted = Font(color=C_MUTED, size=8, italic=True, name='Calibri')

    fill_hdr    = PatternFill("solid", fgColor=C_HEADER)
    fill_accent = PatternFill("solid", fgColor=C_ACCENT)
    fill_stripe = PatternFill("solid", fgColor=C_STRIPE)
    fill_alt    = PatternFill("solid", fgColor=C_ACCENT2)
    fill_white  = PatternFill("solid", fgColor=C_WHITE)
    fill_total  = PatternFill("solid", fgColor=C_ACCENT)

    thin = Side(style="thin", color=C_BORDER)
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    brd_thick_bottom = Border(left=thin, right=thin, top=thin,
                               bottom=Side(style="medium", color=C_ACCENT))

    def safe_title(t):
        bad = r'[:\\/?*\[\]]'
        return _re2.sub(bad, '-', str(t or 'Dados'))[:28]

    def try_num(v):
        if isinstance(v, (int, float)): return v
        s = str(v).replace(' ', '').replace('.', '').replace(',', '.')
        # Remove R$ e %
        s = _re2.sub(r'[R$%]', '', s).strip()
        try: return float(s)
        except: return v

    def is_num_col(rows, col_idx):
        vals = [try_num(r[col_idx]) for r in rows if col_idx < len(r)]
        nums = [v for v in vals if isinstance(try_num(v), float)]
        return len(nums) >= len(vals) * 0.6 and len(nums) > 0

    # ── Parseia blocos de tabela e texto do markdown ───────────────────────
    blocks = _pdf_parse_md(content)
    tables = [(b['headers'], b['rows']) for b in blocks if b['type'] == 'table']
    texts  = [(b['type'], b.get('text','')) for b in blocks
              if b['type'] in ('h1','h2','h3','para','bullets','numbered')]

    # ── ABA PRINCIPAL (Dados) ──────────────────────────────────────────────
    ws = wb.active
    ws.title = safe_title(titulo)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 2  # margem

    row = 1

    # Título
    ws.merge_cells(f"B{row}:J{row}")
    c = ws[f"B{row}"]
    c.value   = titulo or "Relatório"
    c.font    = f_title
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 32
    row += 1

    ws.merge_cells(f"B{row}:J{row}")
    c2 = ws[f"B{row}"]
    c2.value = now_str()
    c2.font  = f_muted
    c2.alignment = Alignment(horizontal="left", indent=1)
    ws.row_dimensions[row].height = 14
    row += 2

    # Linha accent
    for col in range(2, 11):
        c = ws.cell(row=row, column=col)
        c.fill = fill_accent
    ws.row_dimensions[row].height = 3
    row += 2

    # Textos/headings antes das tabelas
    for btype, txt in texts:
        clean = _re2.sub(r'\*\*(.+?)\*\*', r'\1', txt)
        clean = _re2.sub(r'\*(.+?)\*',     r'\1', clean)
        if btype == 'h1':
            ws.merge_cells(f"B{row}:J{row}")
            c = ws[f"B{row}"]
            c.value = clean.upper()
            c.font  = f_sub
            c.fill  = fill_alt
            c.alignment = Alignment(horizontal="left", indent=2, vertical="center")
            c.border = Border(left=Side(style="medium", color=C_ACCENT))
            ws.row_dimensions[row].height = 22
            row += 1
        elif btype == 'h2':
            ws.merge_cells(f"B{row}:J{row}")
            c = ws[f"B{row}"]
            c.value = clean
            c.font  = f_h3
            c.alignment = Alignment(horizontal="left", indent=1)
            ws.row_dimensions[row].height = 18
            row += 1
        elif btype in ('para',):
            ws.merge_cells(f"B{row}:J{row}")
            c = ws[f"B{row}"]
            c.value = clean[:500]
            c.font  = f_body
            c.alignment = Alignment(wrap_text=True, horizontal="left", indent=1)
            ws.row_dimensions[row].height = 15
            row += 1

    if texts: row += 1

    # Tabelas
    for tidx, (headers, rows_data) in enumerate(tables):
        if not headers: continue
        ncols = len(headers)
        start_col = 2  # começa na col B

        # Detecta colunas numéricas
        num_cols = {j for j in range(ncols) if is_num_col(rows_data, j)}

        # Header
        for j, h in enumerate(headers):
            c = ws.cell(row=row, column=start_col+j)
            clean_h = _re2.sub(r'\*\*(.+?)\*\*', r'\1', str(h))
            c.value = clean_h
            c.font  = f_hdr
            c.fill  = fill_hdr
            c.border = brd_thick_bottom
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
        ws.row_dimensions[row].height = 20
        data_start_row = row + 1
        row += 1

        # Dados
        for ri, data_row in enumerate(rows_data):
            fill = fill_stripe if ri % 2 == 0 else fill_white
            for j in range(ncols):
                c = ws.cell(row=row, column=start_col+j)
                raw_val = data_row[j] if j < len(data_row) else ''
                num_val = try_num(raw_val)
                c.value = num_val
                c.font  = f_num if j in num_cols else f_body
                c.fill  = fill
                c.border = brd
                c.alignment = Alignment(
                    horizontal="right" if j in num_cols else "left",
                    vertical="center", wrap_text=True, indent=1)
                if j in num_cols and isinstance(num_val, float):
                    # Detecta % ou R$
                    raw_s = str(raw_val)
                    if '%' in raw_s:
                        c.number_format = '0.00%'
                    elif any(x in raw_s for x in ['R$','r$']):
                        c.number_format = 'R$ #,##0.00'
                    else:
                        c.number_format = '#,##0.00'
            ws.row_dimensions[row].height = 18
            row += 1

        # Linha de totais para colunas numéricas
        if rows_data and num_cols:
            for j in range(ncols):
                c = ws.cell(row=row, column=start_col+j)
                if j == 0:
                    c.value = "TOTAL"
                    c.font  = f_total
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif j in num_cols:
                    col_letter = get_column_letter(start_col+j)
                    c.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{row-1})"
                    c.font  = f_total
                    c.number_format = '#,##0.00'
                    c.alignment = Alignment(horizontal="right", vertical="center")
                c.fill   = fill_accent
                c.border = brd
            ws.row_dimensions[row].height = 20

            # Formatação condicional nas colunas numéricas
            for j in num_cols:
                col_letter = get_column_letter(start_col+j)
                cell_range = f"{col_letter}{data_start_row}:{col_letter}{row-1}"
                ws.conditional_formatting.add(cell_range,
                    ColorScaleRule(
                        start_type='min', start_color='FED7D7',
                        mid_type='percentile', mid_value=50, mid_color='FEFCBF',
                        end_type='max', end_color='C6F6D5'
                    )
                )
            row += 1

        row += 2  # espaço entre tabelas

    # ── ABA RESUMO (se tiver múltiplas tabelas) ────────────────────────────
    if len(tables) > 1:
        ws_res = wb.create_sheet("Resumo", 0)
        ws_res.sheet_view.showGridLines = False
        ws_res.merge_cells("A1:E1")
        c = ws_res["A1"]
        c.value = f"Resumo — {titulo or 'Relatório'}"
        c.font  = f_title
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_res.row_dimensions[1].height = 30
        ws_res.merge_cells("A2:E2")
        c2 = ws_res["A2"]
        c2.value = now_str()
        c2.font  = f_muted
        c2.alignment = Alignment(horizontal="center")
        r = 4
        for i, (headers, rows_data) in enumerate(tables):
            ws_res.cell(row=r, column=1, value=f"Tabela {i+1}").font = f_sub
            ws_res.cell(row=r, column=2, value=f"{len(rows_data)} linhas × {len(headers)} colunas").font = f_body
            r += 1

    # ── Largura automática das colunas ────────────────────────────────────
    from openpyxl.cell.cell import MergedCell
    for ws_item in wb.worksheets:
        for col in ws_item.columns:
            max_len = 10
            col_letter = None
            for c in col:
                if isinstance(c, MergedCell): continue
                if col_letter is None:
                    try: col_letter = c.column_letter
                    except: pass
                try:
                    if c.value: max_len = max(max_len, len(str(c.value)))
                except: pass
            if col_letter:
                ws_item.column_dimensions[col_letter].width = min(max(max_len+2, 12), 45)

    # Rodapé
    row += 1
    ws.merge_cells(f"B{row}:J{row}")
    c = ws[f"B{row}"]
    c.value = f"Gerado por Master IA · {now_str()}"
    c.font  = f_muted
    c.alignment = Alignment(horizontal="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
#  ROTAS
# ══════════════════════════════════════════════════════════════════════════════
def auth(req):
    token = req.headers.get("X-Token") or ""
    if not token: return None
    users = load_users()
    for u, d in users.items():
        if d.get("token") == token:
            return u
    return None

def hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()

# Usuários padrão — base que nunca some
_USERS_DEFAULT = {
    "italo":   {"senha": hash_pw("master123"), "nome": "Ítalo"},
    "willian": {"senha": hash_pw("master123"), "nome": "Willian"},
    "augusto": {"senha": hash_pw("master123"), "nome": "Augusto"},
    "trinid":  {"senha": hash_pw("master123"), "nome": "Trinid"},
    "rafael":  {"senha": hash_pw("master123"), "nome": "Rafael"},
    "diogo":   {"senha": hash_pw("master123"), "nome": "Diogo"},
    "bia":     {"senha": hash_pw("master123"), "nome": "Bia"},
}

USERS_FILE = DATA_DIR / "users.json"

def load_users():
    """Carrega usuários do arquivo, mesclando com defaults."""
    try:
        if USERS_FILE.exists():
            saved = json.loads(USERS_FILE.read_text(encoding="utf-8"))
            # Mescla: defaults + saved (saved tem prioridade — preserva senhas e tokens)
            merged = {**{k: {**v, "token": ""} for k, v in _USERS_DEFAULT.items()}}
            for k, v in saved.items():
                merged[k] = v
            return merged
    except Exception:
        pass
    return {k: {**v, "token": ""} for k, v in _USERS_DEFAULT.items()}

def save_users(users):
    """Salva usuários em arquivo persistente."""
    try:
        USERS_FILE.write_text(json.dumps(users, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

def seed_users():
    """Garante que todos os usuários default existem no arquivo."""
    users = load_users()
    changed = False
    for k, v in _USERS_DEFAULT.items():
        if k not in users:
            users[k] = {**v, "token": ""}
            changed = True
    if changed:
        save_users(users)

@app.route("/", methods=["GET"])
def index():
    html = Path(__file__).parent / "master_chat.html"
    return html.read_text(encoding="utf-8") if html.exists() else "master_chat.html não encontrado", 404

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "time": now_str()})

@app.route("/login", methods=["POST"])
def login():
    d = request.get_json()
    user = (d.get("username") or "").strip().lower()
    pw   = (d.get("senha") or "").strip()
    users = load_users()
    u = users.get(user)
    if not u or u["senha"] != hash_pw(pw):
        return jsonify({"erro": "Usuário ou senha inválidos"}), 401
    token = secrets.token_hex(16)
    users[user]["token"] = token
    save_users(users)
    registrar_evento("login", user)
    return jsonify({"token": token, "nome": u["nome"], "username": user})

@app.route("/gerar", methods=["POST"])
def gerar():
    d        = request.get_json()
    tipo     = (d.get("tipo") or "pdf").lower()
    titulo   = d.get("titulo") or "Documento"
    conteudo = d.get("conteudo") or ""
    imagens  = d.get("imagens") or []  # lista de {b64, legenda}
    if not conteudo:
        return jsonify({"erro": "Conteúdo vazio"}), 400
    try:
        if tipo in ("word", "docx"):
            data = gen_word(titulo, conteudo)
            ext  = "docx"
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif tipo in ("excel", "xlsx"):
            data = gen_excel(titulo, conteudo)
            ext  = "xlsx"
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            data = gen_pdf(titulo, conteudo, imagens=imagens)
            ext  = "pdf"
            mime = "application/pdf"

        fname = f"{safe_name(titulo)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        (DOCS_DIR / fname).write_bytes(data)

        # Registra evento se autenticado
        u = auth(request) or "anonimo"
        registrar_evento("documento", u, tipo)

        return send_file(io.BytesIO(data), as_attachment=True,
                         download_name=f"{safe_name(titulo)}.{ext}", mimetype=mime)
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500

# ── CHATS ─────────────────────────────────────────────────────────────────────
CHATS_DIR = DATA_DIR / "chats"
CHATS_DIR.mkdir(exist_ok=True)

def chats_file(username):
    return CHATS_DIR / f"{username}.json"

def load_chats(username):
    f = chats_file(username)
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else []
    except Exception:
        return []

def save_chats(username, chats):
    chats_file(username).write_text(json.dumps(chats, ensure_ascii=False), encoding="utf-8")

@app.route("/chats", methods=["GET"])
def get_chats():
    user = auth(request)
    if not user: return jsonify({"erro": "Não autenticado"}), 401
    return jsonify(load_chats(user))

@app.route("/chats", methods=["POST"])
def post_chat():
    user = auth(request)
    if not user: return jsonify({"erro": "Não autenticado"}), 401
    d = request.get_json()
    chats = load_chats(user)
    idx = next((i for i, c in enumerate(chats) if c.get("id") == d.get("id")), -1)
    if idx >= 0:
        chats[idx] = d
    else:
        chats.append(d)
        # Conta nova mensagem do usuário
        msgs = d.get("msgs", [])
        n_user = sum(1 for m in msgs if m.get("role") == "user")
        if n_user > 0:
            registrar_evento("mensagem", user, d.get("id",""))
    save_chats(user, chats)
    return jsonify({"ok": True})

@app.route("/chats/<cid>", methods=["GET"])
def get_chat(cid):
    user = auth(request)
    if not user: return jsonify({"erro": "Não autenticado"}), 401
    chats = load_chats(user)
    chat = next((c for c in chats if c.get("id") == cid), None)
    if not chat: return jsonify({"erro": "Chat não encontrado"}), 404
    return jsonify(chat)

@app.route("/chats/<cid>", methods=["DELETE"])
def del_chat(cid):
    user = auth(request)
    if not user: return jsonify({"erro": "Não autenticado"}), 401
    chats = [c for c in load_chats(user) if c.get("id") != cid]
    save_chats(user, chats)
    return jsonify({"ok": True})


@app.route("/admin/usuarios", methods=["POST"])
def criar_usuario():
    d = request.get_json()
    username = (d.get("username") or "").strip().lower()
    nome     = (d.get("nome") or "").strip()
    senha    = (d.get("senha") or "master123").strip()
    if not username or not nome:
        return jsonify({"erro": "username e nome obrigatórios"}), 400
    users = load_users()
    if username in users:
        return jsonify({"erro": "Usuário já existe"}), 409
    users[username] = {"senha": hash_pw(senha), "nome": nome}
    save_users(users)
    return jsonify({"ok": True})

if __name__ == "__main__":
    import socket
    seed_users()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80)); ip = s.getsockname()[0]; s.close()
    except: ip = "127.0.0.1"
    print("="*55)
    print("  Master IA — Servidor")
    print("="*55)
    print(f"  ✅ Rodando!")
    print(f"  Local:  http://localhost:5000")
    print(f"  Rede:   http://{ip}:5000")
    print("="*55)
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

# ══════════════════════════════════════════════════════════════════════════════
#  MÓDULO DE LOGS / DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
import threading
from collections import defaultdict

_lock = threading.Lock()
LOGS_FILE = Path(os.environ.get("DATA_DIR", "/tmp")) / "master_logs.json"

def _load_logs():
    try:
        if LOGS_FILE.exists():
            return json.loads(LOGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {"events": []}

def _save_logs(data):
    try:
        LOGS_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

def registrar_evento(tipo: str, usuario: str, detalhe: str = ""):
    """Registra um evento (mensagem, documento, login) no log."""
    with _lock:
        data = _load_logs()
        data["events"].append({
            "tipo": tipo,          # "mensagem" | "documento" | "login"
            "usuario": usuario,
            "detalhe": detalhe,
            "ts": datetime.now().isoformat()
        })
        # Mantém só últimos 5000 eventos
        if len(data["events"]) > 5000:
            data["events"] = data["events"][-5000:]
        _save_logs(data)

@app.route("/dashboard/stats", methods=["GET"])
def dashboard_stats():
    user = auth(request)
    if not user:
        return jsonify({"erro": "Não autenticado"}), 401

    with _lock:
        data = _load_logs()

    events = data.get("events", [])

    # ── Totais gerais ──
    total_msgs   = sum(1 for e in events if e["tipo"] == "mensagem")
    total_docs   = sum(1 for e in events if e["tipo"] == "documento")
    total_logins = sum(1 for e in events if e["tipo"] == "login")
    usuarios_ativos = len({e["usuario"] for e in events})

    # ── Uso por usuário (mensagens) ──
    uso_usuario = defaultdict(int)
    for e in events:
        if e["tipo"] == "mensagem":
            uso_usuario[e["usuario"]] += 1
    uso_usuario = [{"nome": k, "total": v}
                   for k, v in sorted(uso_usuario.items(), key=lambda x: -x[1])]

    # ── Documentos por tipo ──
    docs_tipo = defaultdict(int)
    for e in events:
        if e["tipo"] == "documento":
            docs_tipo[e.get("detalhe", "pdf")] += 1
    docs_tipo = [{"tipo": k.upper(), "total": v}
                 for k, v in docs_tipo.items()]

    # ── Atividade últimos 7 dias ──
    from datetime import timedelta
    hoje = datetime.now().date()
    dias = {}
    for i in range(6, -1, -1):
        d = (hoje - timedelta(days=i)).isoformat()
        dias[d] = {"mensagens": 0, "documentos": 0}
    for e in events:
        try:
            d = e["ts"][:10]
            if d in dias:
                if e["tipo"] == "mensagem":
                    dias[d]["mensagens"] += 1
                elif e["tipo"] == "documento":
                    dias[d]["documentos"] += 1
        except Exception:
            pass
    atividade = [{"data": k[-5:], "mensagens": v["mensagens"], "documentos": v["documentos"]}
                 for k, v in dias.items()]

    # ── Top 5 chats mais longos ──
    chats_count = defaultdict(int)
    for e in events:
        if e["tipo"] == "mensagem":
            chats_count[e.get("detalhe", "")] += 1
    # Não expõe conteúdo — só contagem anônima

    # ── Últimos eventos (para o admin) ──
    ultimos = []
    for e in reversed(events[-30:]):
        ultimos.append({
            "tipo": e["tipo"],
            "usuario": e["usuario"],
            "detalhe": e.get("detalhe", ""),
            "ts": e["ts"][11:16]  # HH:MM
        })

    return jsonify({
        "totais": {
            "mensagens": total_msgs,
            "documentos": total_docs,
            "logins": total_logins,
            "usuarios_ativos": usuarios_ativos,
        },
        "uso_usuario": uso_usuario,
        "docs_tipo": docs_tipo,
        "atividade": atividade,
        "ultimos": ultimos,
        "usuario_logado": user,
        "is_admin": user == "italo",
    })


# ══════════════════════════════════════════════════════════════════════════════
#  MÓDULO ADMIN
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/admin/listar", methods=["GET"])
def admin_listar():
    user = auth(request)
    if user != "italo":
        return jsonify({"erro": "Acesso negado"}), 403
    users = load_users()
    lista = []
    for u, d in users.items():
        lista.append({
            "username": u,
            "nome": d.get("nome", u),
            "tem_token": bool(d.get("token")),
        })
    return jsonify(lista)

@app.route("/admin/resetar_senha", methods=["POST"])
def admin_resetar():
    user = auth(request)
    if user != "italo":
        return jsonify({"erro": "Acesso negado"}), 403
    d = request.get_json()
    alvo = (d.get("username") or "").strip().lower()
    nova  = (d.get("nova_senha") or "master123").strip()
    users = load_users()
    if alvo not in users:
        return jsonify({"erro": "Usuário não encontrado"}), 404
    users[alvo]["senha"] = hash_pw(nova)
    save_users(users)
    return jsonify({"ok": True})

@app.route("/admin/remover", methods=["POST"])
def admin_remover():
    user = auth(request)
    if user != "italo":
        return jsonify({"erro": "Acesso negado"}), 403
    d = request.get_json()
    alvo = (d.get("username") or "").strip().lower()
    if alvo == "italo":
        return jsonify({"erro": "Não pode remover o admin"}), 400
    users = load_users()
    if alvo not in users:
        return jsonify({"erro": "Usuário não encontrado"}), 404
    del users[alvo]
    save_users(users)
    return jsonify({"ok": True})

@app.route("/admin/adicionar", methods=["POST"])
def admin_adicionar():
    user = auth(request)
    if user != "italo":
        return jsonify({"erro": "Acesso negado"}), 403
    d = request.get_json()
    username = (d.get("username") or "").strip().lower()
    nome     = (d.get("nome") or "").strip()
    senha    = (d.get("senha") or "master123").strip()
    if not username or not nome:
        return jsonify({"erro": "username e nome obrigatórios"}), 400
    users = load_users()
    if username in users:
        return jsonify({"erro": "Usuário já existe"}), 409
    users[username] = {"senha": hash_pw(senha), "nome": nome, "token": ""}
    save_users(users)
    return jsonify({"ok": True})


# ══════════════════════════════════════════════════════════════════════════════
#  MÓDULO COMPARATIVO FISCONTECH x DOMÍNIO
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/comparativo", methods=["POST"])
def comparativo():
    user = auth(request)
    if not user:
        return jsonify({"erro": "Não autenticado"}), 401

    d       = request.get_json()
    dados_a = d.get("dados_a") or ""  # CSV/texto do Fiscontech
    dados_b = d.get("dados_b") or ""  # CSV/texto do Domínio
    label_a = d.get("label_a") or "Fiscontech"
    label_b = d.get("label_b") or "Domínio"

    if not dados_a or not dados_b:
        return jsonify({"erro": "Envie dados_a e dados_b"}), 400

    # Gera Excel comparativo
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparativo"

    az_hdr  = PatternFill("solid", fgColor="1a3a6e")
    verde   = PatternFill("solid", fgColor="e6faf7")
    amarelo = PatternFill("solid", fgColor="fff8e1")
    vermelho= PatternFill("solid", fgColor="ffe5e5")
    alt     = PatternFill("solid", fgColor="eef2ff")
    white   = PatternFill("solid", fgColor="FFFFFF")

    f_hdr   = Font(bold=True, color="FFFFFF", size=11)
    f_tit   = Font(bold=True, color="1a3a6e", size=14)
    f_sub   = Font(bold=True, color="2d5fa3", size=11)
    f_body  = Font(size=10)
    f_diff  = Font(bold=True, color="c0392b", size=10)
    f_ok    = Font(color="1a7a5c", size=10)
    brd = Border(
        left=Side(style="thin", color="d0d8f0"),
        right=Side(style="thin", color="d0d8f0"),
        top=Side(style="thin", color="d0d8f0"),
        bottom=Side(style="thin", color="d0d8f0"),
    )

    def parsecsv(txt):
        """Parseia CSV simples ou texto tabulado."""
        rows = []
        for line in txt.strip().split("\n"):
            if not line.strip():
                continue
            # tenta ; depois , depois \t
            for sep in (";", ",", "\t"):
                parts = line.split(sep)
                if len(parts) >= 2:
                    rows.append([p.strip() for p in parts])
                    break
            else:
                rows.append([line.strip()])
        return rows

    rows_a = parsecsv(dados_a)
    rows_b = parsecsv(dados_b)

    # Título
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = f"Comparativo {label_a} × {label_b}"
    c.font  = f_tit
    c.fill  = PatternFill("solid", fgColor="f5f7ff")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:H2")
    c2 = ws["A2"]
    c2.value = now_str()
    c2.font  = Font(color="AAAAAA", size=8, italic=True)
    c2.alignment = Alignment(horizontal="center")

    row = 4

    # ── Seção A ──────────────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:H{row}")
    ca = ws[f"A{row}"]
    ca.value = f"▌ {label_a}"
    ca.font  = f_sub
    ca.fill  = PatternFill("solid", fgColor="dbeafe")
    ca.alignment = Alignment(indent=1)
    row += 1

    for ri, r in enumerate(rows_a):
        for ci, val in enumerate(r[:8], 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.font  = Font(bold=(ri==0), color="FFFFFF" if ri==0 else "2c3347", size=10)
            c.fill  = az_hdr if ri == 0 else (alt if ri%2==0 else white)
            c.border = brd
            c.alignment = Alignment(wrap_text=True)
        row += 1
    row += 1

    # ── Seção B ──────────────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:H{row}")
    cb = ws[f"A{row}"]
    cb.value = f"▌ {label_b}"
    cb.font  = f_sub
    cb.fill  = PatternFill("solid", fgColor="dcfce7")
    cb.alignment = Alignment(indent=1)
    row += 1

    for ri, r in enumerate(rows_b):
        for ci, val in enumerate(r[:8], 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.font  = Font(bold=(ri==0), color="FFFFFF" if ri==0 else "2c3347", size=10)
            c.fill  = PatternFill("solid", fgColor="166534") if ri==0 else (verde if ri%2==0 else white)
            c.border = brd
            c.alignment = Alignment(wrap_text=True)
        row += 1
    row += 1

    # ── Diferenças numéricas automáticas ─────────────────────────────────────
    def extrair_valores(rows):
        """Extrai pares (chave, valor_float) da última coluna numérica."""
        vals = {}
        for r in rows[1:]:  # pula header
            if len(r) >= 2:
                chave = r[0]
                for cell_val in reversed(r):
                    v = cell_val.replace(".", "").replace(",", ".").replace("R$","").strip()
                    try:
                        vals[chave] = float(v)
                        break
                    except ValueError:
                        pass
        return vals

    va = extrair_valores(rows_a)
    vb = extrair_valores(rows_b)
    todas_chaves = sorted(set(va) | set(vb))

    if todas_chaves:
        ws.merge_cells(f"A{row}:H{row}")
        cd = ws[f"A{row}"]
        cd.value = "▌ Diferenças"
        cd.font  = Font(bold=True, color="c0392b", size=11)
        cd.fill  = PatternFill("solid", fgColor="fee2e2")
        cd.alignment = Alignment(indent=1)
        row += 1

        # Header diferenças
        hdrs = ["Chave/Empresa", label_a, label_b, "Diferença", "Status"]
        for ci, h in enumerate(hdrs, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.font = f_hdr; c.fill = az_hdr; c.border = brd
            c.alignment = Alignment(horizontal="center")
        row += 1

        for ri, chave in enumerate(todas_chaves):
            val_a = va.get(chave)
            val_b = vb.get(chave)
            diff  = None
            status = ""
            if val_a is not None and val_b is not None:
                diff = val_b - val_a
                status = "✅ OK" if abs(diff) < 0.02 else f"⚠ Dif"
            elif val_a is None:
                status = "❌ Só no " + label_b
            else:
                status = "❌ Só no " + label_a

            bg = verde if status.startswith("✅") else (amarelo if "Dif" in status else vermelho)
            row_data = [
                chave,
                f"R$ {val_a:,.2f}".replace(",","X").replace(".",",").replace("X",".") if val_a is not None else "—",
                f"R$ {val_b:,.2f}".replace(",","X").replace(".",",").replace("X",".") if val_b is not None else "—",
                f"R$ {diff:,.2f}".replace(",","X").replace(".",",").replace("X",".") if diff is not None else "—",
                status,
            ]
            for ci, val in enumerate(row_data, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.font  = f_diff if "⚠" in str(status) or "❌" in str(status) else f_ok
                c.fill  = bg if ri%2==0 else bg
                c.border = brd
                c.alignment = Alignment(horizontal="center" if ci > 1 else "left", wrap_text=True)
            row += 1

    # Largura das colunas
    from openpyxl.cell.cell import MergedCell
    for col in ws.columns:
        ml = 10
        col_letter = None
        for c in col:
            if isinstance(c, MergedCell): continue
            if col_letter is None:
                try: col_letter = c.column_letter
                except: pass
            try:
                if c.value: ml = max(ml, len(str(c.value)))
            except: pass
        if col_letter:
            ws.column_dimensions[col_letter].width = min(max(ml + 2, 12), 50)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    data_bytes = buf.read()

    registrar_evento("documento", user, "comparativo")

    fname = f"Comparativo_{safe_name(label_a)}_{safe_name(label_b)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    (DOCS_DIR / fname).write_bytes(data_bytes)

    return send_file(
        io.BytesIO(data_bytes),
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



# ══════════════════════════════════════════════════════════════════════════════
#  FERRAMENTAS / TOOL USE
# ══════════════════════════════════════════════════════════════════════════════
import base64, subprocess, sys, textwrap, traceback as tb
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# Paleta Master
MASTER_COLORS = ["#3d7eff","#00c2a8","#ff6b6b","#ffd93d","#a78bfa","#fb923c","#34d399","#60a5fa"]

def _apply_master_style(fig, ax_list=None):
    fig.patch.set_facecolor("#181c26")
    for ax in (ax_list or fig.get_axes()):
        ax.set_facecolor("#1e2330")
        ax.tick_params(colors="#8892a4", labelsize=9)
        ax.xaxis.label.set_color("#8892a4")
        ax.yaxis.label.set_color("#8892a4")
        ax.title.set_color("#e8eaf0")
        for spine in ax.spines.values():
            spine.set_edgecolor("#2a3045")
        ax.grid(color="#2a3045", linewidth=0.5, alpha=0.7)


def _fig_to_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()


# ── Gráfico de barras ──────────────────────────────────────────────────────────
def _safe_floats(valores):
    """Converte lista para floats de forma segura."""
    result = []
    for v in valores:
        try:
            result.append(float(v))
        except (TypeError, ValueError):
            result.append(0.0)
    return result

def _safe_labels(labels):
    """Garante que labels são strings."""
    return [str(l) for l in labels]

def tool_grafico_barras(labels, valores, titulo="", xlabel="", ylabel="", horizontal=False):
    labels  = _safe_labels(labels)
    valores = _safe_floats(valores)
    # Garante mesmo tamanho
    n = min(len(labels), len(valores))
    labels, valores = labels[:n], valores[:n]
    if n == 0: raise ValueError("Dados vazios para gráfico de barras")

    fig, ax = plt.subplots(figsize=(max(7, n*0.8), 4))
    _apply_master_style(fig, [ax])
    cores = [MASTER_COLORS[i % len(MASTER_COLORS)] for i in range(n)]
    y = np.arange(n)
    max_v = max(valores) if valores else 1

    if horizontal:
        bars = ax.barh(y, valores, color=cores, edgecolor="#2a3045", linewidth=0.5)
        ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=9)
        for bar, val in zip(bars, valores):
            ax.text(bar.get_width() + max_v*0.01, bar.get_y()+bar.get_height()/2,
                    f"{val:,.0f}".replace(",","."),
                    va="center", ha="left", color="#e8eaf0", fontsize=8)
    else:
        bars = ax.bar(y, valores, color=cores, edgecolor="#2a3045", linewidth=0.5, width=0.65)
        ax.set_xticks(y)
        ax.set_xticklabels(labels, rotation=25, ha="right", fontsize=8)
        for bar, val in zip(bars, valores):
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+max_v*0.01,
                    f"{val:,.0f}".replace(",","."),
                    ha="center", va="bottom", color="#e8eaf0", fontsize=8)

    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10, color="#e8eaf0")
    if xlabel: ax.set_xlabel(xlabel, color="#8892a4")
    if ylabel: ax.set_ylabel(ylabel, color="#8892a4")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f"{x:,.0f}".replace(",",".")))
    fig.tight_layout()
    return _fig_to_b64(fig)


# ── Gráfico de linhas ──────────────────────────────────────────────────────────
def tool_grafico_linhas(series, titulo="", xlabel="", ylabel=""):
    """series: [{"nome": str, "dados": [float,...], "labels": [str,...]}]"""
    if not series: raise ValueError("Series vazio")
    # Garante que cada série é um dict com campos corretos
    series_clean = []
    for s in series:
        if not isinstance(s, dict): continue
        nome   = str(s.get("nome","Série"))
        dados  = _safe_floats(s.get("dados") or s.get("values") or s.get("data") or [])
        labels = _safe_labels(s.get("labels") or s.get("x") or [])
        if dados:
            series_clean.append({"nome": nome, "dados": dados, "labels": labels})

    if not series_clean: raise ValueError("Nenhuma série válida")

    fig, ax = plt.subplots(figsize=(8, 4))
    _apply_master_style(fig, [ax])

    for i, s in enumerate(series_clean):
        cor = MASTER_COLORS[i % len(MASTER_COLORS)]
        xs  = range(len(s["dados"]))
        ax.plot(xs, s["dados"], color=cor, linewidth=2.5, marker="o",
                markersize=5, label=s["nome"])
        ax.fill_between(xs, s["dados"], alpha=0.08, color=cor)

    # Labels do eixo X
    all_labels = series_clean[0].get("labels",[])
    if all_labels:
        ax.set_xticks(range(len(all_labels)))
        ax.set_xticklabels(all_labels, rotation=20, ha="right", fontsize=8)

    if len(series_clean) > 1 or series_clean[0]["nome"] != "Série":
        ax.legend(fontsize=8, facecolor="#1e2330", edgecolor="#2a3045", labelcolor="#e8eaf0")
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10, color="#e8eaf0")
    if xlabel: ax.set_xlabel(xlabel, color="#8892a4")
    if ylabel: ax.set_ylabel(ylabel, color="#8892a4")
    fig.tight_layout()
    return _fig_to_b64(fig)


# ── Gráfico de pizza ───────────────────────────────────────────────────────────
def tool_grafico_pizza(labels, valores, titulo=""):
    labels  = _safe_labels(labels)
    valores = _safe_floats(valores)
    n = min(len(labels), len(valores))
    labels, valores = labels[:n], valores[:n]
    if n == 0: raise ValueError("Dados vazios para pizza")

    # Limita labels longos
    labels_short = [l[:18]+'…' if len(l) > 18 else l for l in labels]
    cores = [MASTER_COLORS[i % len(MASTER_COLORS)] for i in range(n)]

    fig, ax = plt.subplots(figsize=(7, 5.5))
    _apply_master_style(fig, [ax])

    wedges, texts, autotexts = ax.pie(
        valores,
        labels=None,
        colors=cores,
        autopct="%1.1f%%",
        startangle=90,
        wedgeprops=dict(edgecolor="#181c26", linewidth=1.5),
        pctdistance=0.75
    )
    for at in autotexts:
        at.set_color("#e8eaf0"); at.set_fontsize(9); at.set_fontweight("bold")

    # Legenda fora do gráfico para não sobrepor
    ax.legend(
        wedges, labels_short,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=min(3, n),
        fontsize=8,
        facecolor="#1e2330",
        edgecolor="#2a3045",
        labelcolor="#e8eaf0",
        framealpha=0.9,
        handlelength=1.2,
        handleheight=0.8
    )
    if titulo:
        ax.set_title(titulo, fontsize=12, fontweight="bold", color="#e8eaf0", pad=12)

    fig.subplots_adjust(bottom=0.25)
    return _fig_to_b64(fig)


# ── Gráfico de dispersão ───────────────────────────────────────────────────────
def tool_grafico_dispersao(series, titulo="", xlabel="", ylabel=""):
    """series: [{"nome": str, "x": [float], "y": [float]}]"""
    if not series: raise ValueError("Series vazio")
    series_clean = []
    for s in series:
        if not isinstance(s, dict): continue
        nome = str(s.get("nome",""))
        x    = _safe_floats(s.get("x",[]))
        y    = _safe_floats(s.get("y",[]))
        n    = min(len(x), len(y))
        if n > 0:
            series_clean.append({"nome": nome, "x": x[:n], "y": y[:n]})
    if not series_clean: raise ValueError("Nenhuma série válida")

    fig, ax = plt.subplots(figsize=(7, 4))
    _apply_master_style(fig, [ax])
    for i, s in enumerate(series_clean):
        cor = MASTER_COLORS[i % len(MASTER_COLORS)]
        ax.scatter(s["x"], s["y"], color=cor, label=s["nome"],
                   alpha=0.85, s=50, edgecolors="#181c26", linewidth=0.5)
    if len(series_clean) > 1:
        ax.legend(fontsize=8, facecolor="#1e2330", edgecolor="#2a3045", labelcolor="#e8eaf0")
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10, color="#e8eaf0")
    if xlabel: ax.set_xlabel(xlabel, color="#8892a4")
    if ylabel: ax.set_ylabel(ylabel, color="#8892a4")
    fig.tight_layout()
    return _fig_to_b64(fig)


# ── Execução de Python segura ──────────────────────────────────────────────────
EXEC_TIMEOUT = 10  # segundos

def tool_executar_python(codigo: str):
    """
    Executa código Python em subprocesso isolado.
    Retorna {"stdout": str, "stderr": str, "erro": bool, "imagem_b64": str|None}
    """
    # Injeta captura de plt.show() → base64
    wrapper = textwrap.dedent(f"""
import sys, io, base64, json
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as _plt_orig
_fig_b64 = None

def _capture_show():
    global _fig_b64
    buf = io.BytesIO()
    _plt_orig.savefig(buf, format="png", dpi=130, bbox_inches="tight",
                      facecolor=_plt_orig.gcf().get_facecolor())
    _plt_orig.close()
    buf.seek(0)
    _fig_b64 = base64.b64encode(buf.read()).decode()

import matplotlib.pyplot as plt
plt.show = _capture_show

# ─── CÓDIGO DO USUÁRIO ───
{codigo}
# ─────────────────────────

# Se gerou figura mas não chamou show, captura
if _plt_orig.get_fignums() and _fig_b64 is None:
    _capture_show()

print("__FIG__:" + (_fig_b64 or ""), file=sys.stderr)
""")
    try:
        proc = subprocess.run(
            [sys.executable, "-c", wrapper],
            capture_output=True, text=True, timeout=EXEC_TIMEOUT
        )
        # Extrai imagem do stderr
        imagem_b64 = None
        stderr_lines = []
        for line in proc.stderr.splitlines():
            if line.startswith("__FIG__:"):
                val = line[8:].strip()
                if val:
                    imagem_b64 = val
            else:
                stderr_lines.append(line)

        return {
            "stdout": proc.stdout.strip(),
            "stderr": "\n".join(stderr_lines).strip(),
            "erro": proc.returncode != 0,
            "imagem_b64": imagem_b64
        }
    except subprocess.TimeoutExpired:
        return {"stdout": "", "stderr": f"Timeout: execução excedeu {EXEC_TIMEOUT}s", "erro": True, "imagem_b64": None}
    except Exception as e:
        return {"stdout": "", "stderr": str(e), "erro": True, "imagem_b64": None}


# ── Excel avançado com gráfico ─────────────────────────────────────────────────
def tool_excel_avancado(titulo, colunas, linhas, grafico_tipo=None, grafico_series=None):
    """
    colunas: [str]
    linhas:  [[valor, ...]]
    grafico_tipo: "barras" | "linhas" | "pizza" | None
    grafico_series: índices das colunas numéricas para o gráfico
    """
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference, Series as XLSeries
    from openpyxl.chart.series import DataPoint

    wb = openpyxl.Workbook()
    ws = wb.active
    # Caracteres inválidos em nomes de aba Excel: : \ / ? * [ ]
    import re as _re
    titulo_aba = _re.sub(r'[:\\/?*\[\]]', '-', titulo or 'Dados')[:28]
    ws.title = titulo_aba

    az    = PatternFill("solid", fgColor="1a3a6e")
    az2   = PatternFill("solid", fgColor="2d5fa3")
    verde = PatternFill("solid", fgColor="e6faf7")
    alt   = PatternFill("solid", fgColor="f5f7ff")
    white = PatternFill("solid", fgColor="FFFFFF")
    brd   = Border(
        left=Side(style="thin",   color="d0d8f0"),
        right=Side(style="thin",  color="d0d8f0"),
        top=Side(style="thin",    color="d0d8f0"),
        bottom=Side(style="thin", color="d0d8f0"),
    )

    ncols = len(colunas)

    # ── Título ──
    ws.merge_cells(f"A1:{chr(64+ncols)}1")
    c = ws["A1"]
    c.value = titulo
    c.font  = Font(bold=True, color="FFFFFF", size=13)
    c.fill  = az
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(f"A2:{chr(64+ncols)}2")
    c2 = ws["A2"]
    c2.value = now_str()
    c2.font  = Font(color="AAAAAA", size=8, italic=True)
    c2.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 14

    # ── Header ──
    for j, col in enumerate(colunas, 1):
        c = ws.cell(row=3, column=j, value=col)
        c.font  = Font(bold=True, color="FFFFFF", size=10)
        c.fill  = az2
        c.border = brd
        c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 20

    # ── Dados ──
    for ri, row in enumerate(linhas):
        bg = alt if ri % 2 == 0 else white
        for j, val in enumerate(row[:ncols], 1):
            # Tenta converter número
            cell_val = val
            try:
                if isinstance(val, str):
                    v = val.replace(".", "").replace(",", ".")
                    cell_val = float(v) if "." in val or val.lstrip("-").isdigit() else val
            except Exception:
                pass
            c = ws.cell(row=4+ri, column=j, value=cell_val)
            c.font   = Font(size=10)
            c.fill   = bg
            c.border = brd
            is_num   = isinstance(cell_val, (int, float))
            c.alignment = Alignment(
                horizontal="right" if is_num else "left",
                wrap_text=True, vertical="center"
            )
            if is_num:
                c.number_format = '#,##0.00'
        ws.row_dimensions[4+ri].height = 18

    # ── Totais automáticos para colunas numéricas ──
    tot_row = 4 + len(linhas)
    ws.merge_cells(f"A{tot_row}:{chr(64+min(2,ncols))}{tot_row}")
    c = ws[f"A{tot_row}"]
    c.value = "TOTAL"
    c.font  = Font(bold=True, color="FFFFFF", size=10)
    c.fill  = az
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[tot_row].height = 18

    for j in range(3, ncols+1):
        # Verifica se coluna é numérica
        try:
            vals = [ws.cell(row=4+ri, column=j).value for ri in range(len(linhas))]
            if all(isinstance(v, (int, float)) for v in vals if v is not None):
                tc = ws.cell(row=tot_row, column=j)
                tc.value = f"=SUM({chr(64+j)}4:{chr(64+j)}{tot_row-1})"
                tc.font  = Font(bold=True, color="FFFFFF", size=10)
                tc.fill  = az
                tc.border = brd
                tc.number_format = '#,##0.00'
                tc.alignment = Alignment(horizontal="right", vertical="center")
        except Exception:
            pass

    # ── Largura das colunas ──
    from openpyxl.cell.cell import MergedCell
    for col in ws.columns:
        ml = 10
        col_letter = None
        for c in col:
            if isinstance(c, MergedCell): continue
            if col_letter is None:
                try: col_letter = c.column_letter
                except: pass
            try:
                if c.value: ml = max(ml, len(str(c.value)))
            except: pass
        if col_letter:
            ws.column_dimensions[col_letter].width = min(max(ml + 2, 12), 40)

    # ── Gráfico embutido ──
    if grafico_tipo and len(linhas) > 0:
        try:
            data_row_start = 3  # header na linha 3
            data_row_end   = 3 + len(linhas)
            series_cols    = grafico_series or [j for j in range(2, ncols+1)]

            if grafico_tipo == "pizza":
                chart = PieChart()
                chart.title = titulo
                chart.style = 10
                data   = Reference(ws, min_col=series_cols[0], min_row=data_row_start, max_row=data_row_end)
                labels = Reference(ws, min_col=1, min_row=4, max_row=data_row_end)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(labels)
            elif grafico_tipo == "linhas":
                chart = LineChart()
                chart.title = titulo
                chart.style = 10
                chart.grouping = "standard"
                for sc in series_cols:
                    data = Reference(ws, min_col=sc, min_row=data_row_start, max_row=data_row_end)
                    chart.add_data(data, titles_from_data=True)
                cats = Reference(ws, min_col=1, min_row=4, max_row=data_row_end)
                chart.set_categories(cats)
            else:  # barras (padrão)
                chart = BarChart()
                chart.type = "col"
                chart.title = titulo
                chart.style = 10
                chart.grouping = "clustered"
                for sc in series_cols:
                    data = Reference(ws, min_col=sc, min_row=data_row_start, max_row=data_row_end)
                    chart.add_data(data, titles_from_data=True)
                cats = Reference(ws, min_col=1, min_row=4, max_row=data_row_end)
                chart.set_categories(cats)

            chart.width  = 16
            chart.height = 10
            ws.add_chart(chart, f"A{tot_row + 2}")
        except Exception:
            pass  # gráfico opcional — não quebra o arquivo

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Definição das ferramentas para a API Anthropic ────────────────────────────
TOOLS = [
    {
        "name": "gerar_grafico_barras",
        "description": "Gera um gráfico de barras bonito. Use quando o usuário pedir gráfico de barras, comparativo de valores, ranking.",
        "input_schema": {
            "type": "object",
            "properties": {
                "labels":     {"type":"array","items":{"type":"string"},"description":"Nomes das barras"},
                "valores":    {"type":"array","items":{"type":"number"},"description":"Valores numéricos"},
                "titulo":     {"type":"string","description":"Título do gráfico"},
                "xlabel":     {"type":"string","description":"Rótulo eixo X"},
                "ylabel":     {"type":"string","description":"Rótulo eixo Y"},
                "horizontal": {"type":"boolean","description":"True para barras horizontais"}
            },
            "required": ["labels","valores"]
        }
    },
    {
        "name": "gerar_grafico_linhas",
        "description": "Gera gráfico de linhas. Use para séries temporais, evolução ao longo do tempo, tendências.",
        "input_schema": {
            "type": "object",
            "properties": {
                "series": {
                    "type":"array",
                    "description":"Lista de séries",
                    "items": {
                        "type":"object",
                        "properties": {
                            "nome":   {"type":"string"},
                            "dados":  {"type":"array","items":{"type":"number"}},
                            "labels": {"type":"array","items":{"type":"string"}}
                        },
                        "required":["nome","dados"]
                    }
                },
                "titulo":  {"type":"string"},
                "xlabel":  {"type":"string"},
                "ylabel":  {"type":"string"}
            },
            "required": ["series"]
        }
    },
    {
        "name": "gerar_grafico_pizza",
        "description": "Gera gráfico de pizza/rosca. Use para mostrar proporções, participação percentual, distribuição.",
        "input_schema": {
            "type": "object",
            "properties": {
                "labels":  {"type":"array","items":{"type":"string"},"description":"Categorias"},
                "valores": {"type":"array","items":{"type":"number"},"description":"Valores"},
                "titulo":  {"type":"string"}
            },
            "required": ["labels","valores"]
        }
    },
    {
        "name": "gerar_grafico_dispersao",
        "description": "Gera gráfico de dispersão (scatter plot). Use para correlação entre dois valores numéricos.",
        "input_schema": {
            "type": "object",
            "properties": {
                "series": {
                    "type":"array",
                    "items": {
                        "type":"object",
                        "properties": {
                            "nome": {"type":"string"},
                            "x": {"type":"array","items":{"type":"number"}},
                            "y": {"type":"array","items":{"type":"number"}}
                        },
                        "required":["nome","x","y"]
                    }
                },
                "titulo": {"type":"string"},
                "xlabel": {"type":"string"},
                "ylabel": {"type":"string"}
            },
            "required": ["series"]
        }
    },
    {
        "name": "executar_python",
        "description": "Executa código Python e retorna o resultado. Use quando o usuário pedir para rodar, calcular, processar dados com Python. Pode usar pandas, numpy, matplotlib.",
        "input_schema": {
            "type": "object",
            "properties": {
                "codigo": {"type":"string","description":"Código Python completo para executar"}
            },
            "required": ["codigo"]
        }
    },
    {
        "name": "gerar_excel_avancado",
        "description": "Gera planilha Excel profissional com formatação Master e gráfico embutido opcional. Use para tabelas de dados, relatórios, comparativos.",
        "input_schema": {
            "type": "object",
            "properties": {
                "titulo":         {"type":"string","description":"Título da planilha"},
                "colunas":        {"type":"array","items":{"type":"string"},"description":"Nomes das colunas"},
                "linhas":         {"type":"array","items":{"type":"array"},"description":"Linhas de dados"},
                "grafico_tipo":   {"type":"string","enum":["barras","linhas","pizza"],"description":"Tipo de gráfico embutido (opcional)"},
                "grafico_series": {"type":"array","items":{"type":"integer"},"description":"Índices (1-based) das colunas numéricas para o gráfico"}
            },
            "required": ["titulo","colunas","linhas"]
        }
    }
]


# ── Rota principal: chat com ferramentas ──────────────────────────────────────
import requests as _requests

@app.route("/chat_tools", methods=["POST"])
def chat_tools():
    user = auth(request)
    if not user:
        return jsonify({"erro": "Não autenticado"}), 401

    d        = request.get_json()
    messages = d.get("messages") or []
    api_key  = d.get("api_key") or ""
    model    = d.get("model") or "claude-opus-4-7"

    if not api_key:
        return jsonify({"erro": "api_key obrigatória"}), 400
    if not messages:
        return jsonify({"erro": "messages vazio"}), 400

    system = """Você é o Master IA — uma inteligência artificial de altíssimo nível, completa e abrangente.

QUEM VOCÊ É:
Você é um assistente extraordinariamente capaz, com conhecimento profundo em praticamente qualquer área: ciências, tecnologia, direito, medicina, filosofia, criatividade, matemática, finanças, programação, história, artes e muito mais. Você pensa de forma estruturada, usa raciocínio rigoroso e comunica com clareza e elegância.

Você tem especialidade máxima em contabilidade, fiscal e tributário brasileiro — mas não se limita a isso. Você é igualmente brilhante em qualquer outra área.

PERSONALIDADE:
- Confiante e direto — vai ao ponto sem rodeios desnecessários
- Intelectualmente curioso — demonstra genuíno interesse pelo problema
- Honesto — admite incerteza quando existe, nunca inventa
- Adaptável — tom mais formal para documentos técnicos, mais conversacional para perguntas simples
- Proativo — antecipa o que o usuário vai precisar a seguir
- NUNCA começa respostas com "Claro!", "Ótima pergunta!", "Com certeza!" ou frases vazias

COMO RESPONDE:
- Para perguntas simples: resposta direta e concisa, sem exagerar no tamanho
- Para análises complexas: estruturado com markdown, headers, tabelas, listas quando útil
- Para código: sempre com syntax highlight, comentários relevantes
- Para documentos: completo, bem estruturado, profissional
- NUNCA faz perguntas desnecessárias — responde sempre com o melhor possível

USO INTELIGENTE DE FERRAMENTAS:
- Python para qualquer cálculo, mesmo simples — prefere confirmar do que estimar
- Gráficos quando tem dados que vale visualizar — não precisa ser pedido explicitamente
- Excel quando o usuário precisa de algo para trabalhar, não só ver
- Usa múltiplas ferramentas em sequência naturalmente (calcula → grafíca → explica)

CONHECIMENTO FISCAL BRASILEIRO (especialidade máxima):
SPED, EFD-Reinf, eSocial, DCTFWeb, DARF, DIFAL, CPRB, Simples Nacional, Lucro Presumido, Lucro Real, terceiro setor (MROSC, Lei 13.019/2014), ICMS, PIS, COFINS, ISS, IRPJ, CSLL, INSS, FGTS, NF-e, NFS-e, CT-e, XMLs, Rondônia (SEFIN, SEFISC, alíquotas).

Responda sempre em português brasileiro. Seja extraordinário."""

    # ── Limpa histórico ───────────────────────────────────────────────────────
    msgs_loop = []
    for m in messages:
        role    = m.get("role","")
        content = m.get("content","")
        if role not in ("user","assistant"):
            continue
        if isinstance(content, str) and content.strip():
            content = content.split("<tool_call>")[0].strip()
            if content:
                msgs_loop.append({"role": role, "content": content})
        elif isinstance(content, list):
            # Mantém blocos de imagem (vision) intactos
            clean = []
            for b in content:
                if b.get("type") == "text" and b.get("text","").strip():
                    clean.append({"type":"text","text":b["text"].split("<tool_call>")[0].strip()})
                elif b.get("type") == "image":
                    clean.append(b)
            if clean:
                msgs_loop.append({"role": role, "content": clean})

    if not msgs_loop:
        return jsonify({"erro": "messages vazio"}), 400

    # ── Executa ferramenta e retorna resultado ───────────────────────────────
    def executar_ferramenta(tool_name, tool_input):
        """Executa a ferramenta e retorna (result_block, tool_result_content)."""
        try:
            if tool_name == "gerar_grafico_barras":
                img = tool_grafico_barras(
                    tool_input["labels"], tool_input["valores"],
                    tool_input.get("titulo",""), tool_input.get("xlabel",""),
                    tool_input.get("ylabel",""), tool_input.get("horizontal", False))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Gráfico")}
                result_text = f"Gráfico '{tool_input.get('titulo','')}' gerado com sucesso."

            elif tool_name == "gerar_grafico_linhas":
                img = tool_grafico_linhas(
                    tool_input["series"], tool_input.get("titulo",""),
                    tool_input.get("xlabel",""), tool_input.get("ylabel",""))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Gráfico")}
                result_text = f"Gráfico de linhas '{tool_input.get('titulo','')}' gerado com sucesso."

            elif tool_name == "gerar_grafico_pizza":
                img = tool_grafico_pizza(
                    tool_input["labels"], tool_input["valores"],
                    tool_input.get("titulo",""))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Pizza")}
                result_text = f"Gráfico de pizza '{tool_input.get('titulo','')}' gerado com sucesso."

            elif tool_name == "gerar_grafico_dispersao":
                img = tool_grafico_dispersao(
                    tool_input["series"], tool_input.get("titulo",""),
                    tool_input.get("xlabel",""), tool_input.get("ylabel",""))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Dispersão")}
                result_text = f"Gráfico gerado com sucesso."

            elif tool_name == "executar_python":
                res = tool_executar_python(tool_input["codigo"])
                block = {
                    "tipo": "codigo_resultado",
                    "codigo": tool_input["codigo"],
                    "stdout": res["stdout"],
                    "stderr": res["stderr"],
                    "erro":   res["erro"],
                    "imagem_b64": res.get("imagem_b64")
                }
                # Resultado textual para a IA processar no próximo turno
                out = res["stdout"] or ""
                err = res["stderr"] or ""
                img_info = " [imagem gerada]" if res.get("imagem_b64") else ""
                result_text = f"Código executado.{img_info}\nSaída: {out[:2000]}" + (f"\nErro: {err[:500]}" if err and res["erro"] else "")

            elif tool_name == "gerar_excel_avancado":
                xls_bytes = tool_excel_avancado(
                    tool_input["titulo"], tool_input["colunas"],
                    tool_input["linhas"], tool_input.get("grafico_tipo"),
                    tool_input.get("grafico_series"))
                xls_b64 = base64.b64encode(xls_bytes).decode()
                fname   = safe_name(tool_input["titulo"]) + ".xlsx"
                block   = {"tipo":"excel","b64":xls_b64,"nome":fname,
                           "legenda":tool_input["titulo"]}
                registrar_evento("documento", user, "excel_avancado")
                result_text = f"Planilha Excel '{tool_input['titulo']}' gerada com {len(tool_input['linhas'])} linhas."

            else:
                block = {"tipo":"erro_ferramenta","ferramenta":tool_name,
                         "mensagem":"Ferramenta desconhecida"}
                result_text = f"Ferramenta '{tool_name}' não reconhecida."

            return block, result_text

        except Exception as e:
            block = {"tipo":"erro_ferramenta","ferramenta":tool_name,
                     "mensagem":str(e),"trace":tb.format_exc()}
            return block, f"Erro ao executar {tool_name}: {str(e)}"

    # ── Loop multi-step (até 6 iterações) ────────────────────────────────────
    result_blocks = []
    MAX_ITER = 6

    for iteration in range(MAX_ITER):
        try:
            resp = _requests.post(
                "https://api.iacontaai.com.br/v1/messages",
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {api_key}",
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                },
                json={
                    "model": model,
                    "max_tokens": 8192,
                    "system": system,
                    "tools": TOOLS,
                    "messages": msgs_loop,
                },
                timeout=180
            )
            if not resp.ok:
                return jsonify({"erro": resp.text}), resp.status_code
            data = resp.json()
        except Exception as e:
            import traceback as _tb
            return jsonify({"erro": str(e), "detalhe": _tb.format_exc()}), 500

        stop_reason = data.get("stop_reason","")
        content     = data.get("content", [])

        # Coleta texto e tool_use desta iteração
        tool_uses   = []
        texto_iter  = []

        for block in content:
            btype = block.get("type")
            if btype == "text":
                txt = block.get("text","").split("<tool_call>")[0].strip()
                if txt:
                    texto_iter.append(txt)
                    result_blocks.append({"tipo":"texto","conteudo":txt})
            elif btype == "tool_use":
                tool_uses.append(block)

        # Terminou — sem mais ferramentas
        if stop_reason != "tool_use" or not tool_uses:
            break

        # Executa todas as ferramentas desta iteração
        tool_results = []
        for tu in tool_uses:
            tool_name  = tu.get("name","")
            tool_input = tu.get("input",{})
            tool_id    = tu.get("id", f"tool_{iteration}_{tool_name}")

            rb, result_text = executar_ferramenta(tool_name, tool_input)
            result_blocks.append(rb)

            tool_results.append({
                "type":       "tool_result",
                "tool_use_id": tool_id,
                "content":    result_text
            })

        # Adiciona a resposta da IA + resultados ao histórico para próxima iteração
        msgs_loop.append({"role": "assistant", "content": content})
        msgs_loop.append({"role": "user",      "content": tool_results})

    registrar_evento("mensagem", user)
    return jsonify({"blocos": result_blocks, "stop_reason": stop_reason})


