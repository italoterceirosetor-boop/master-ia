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
#  GERADOR PDF
# ══════════════════════════════════════════════════════════════════════════════
def gen_pdf(titulo, content):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        topMargin=1.2*cm, bottomMargin=1.8*cm,
        leftMargin=2*cm, rightMargin=2*cm)

    # Cores
    AZUL      = colors.HexColor("#1a3a6e")
    AZUL2     = colors.HexColor("#2d5fa3")
    VERDE     = colors.HexColor("#00c2a8")
    VERDE_BG  = colors.HexColor("#e6faf7")
    AZUL_BG   = colors.HexColor("#eef2ff")
    CINZA     = colors.HexColor("#8892a4")
    CINZA_BG  = colors.HexColor("#f5f7fb")
    BORDA     = colors.HexColor("#d0d8f0")
    AMARELO   = colors.HexColor("#fff8e1")
    AMARELO_B = colors.HexColor("#f59e0b")

    W = doc.width

    # Estilos
    s_tit  = ParagraphStyle("tit", fontSize=21, alignment=TA_CENTER,
                             fontName="Helvetica-Bold", textColor=colors.white,
                             spaceAfter=0, spaceBefore=0, leading=26)
    s_dat  = ParagraphStyle("dat", fontSize=8, alignment=TA_CENTER,
                             textColor=colors.HexColor("#90aad4"), spaceAfter=0)
    s_h1   = ParagraphStyle("h1", fontSize=12, fontName="Helvetica-Bold",
                             textColor=colors.white, spaceAfter=0, spaceBefore=0, leading=16)
    s_h2   = ParagraphStyle("h2", fontSize=11, fontName="Helvetica-Bold",
                             textColor=AZUL2, spaceBefore=0, spaceAfter=2, leading=15)
    s_h3   = ParagraphStyle("h3", fontSize=10, fontName="Helvetica-Bold",
                             textColor=CINZA, spaceBefore=0, spaceAfter=2, leading=14)
    s_bod  = ParagraphStyle("bod", fontSize=9.5, leading=15, spaceAfter=4,
                             alignment=TA_JUSTIFY, textColor=colors.HexColor("#2c3347"))
    s_bul  = ParagraphStyle("bul", fontSize=9.5, leading=14,
                             leftIndent=14, firstLineIndent=-10, spaceAfter=2,
                             textColor=colors.HexColor("#2c3347"))
    s_num  = ParagraphStyle("num", fontSize=9.5, leading=14,
                             leftIndent=18, firstLineIndent=-14, spaceAfter=2,
                             textColor=colors.HexColor("#2c3347"))
    s_rod  = ParagraphStyle("rod", fontSize=7.5, textColor=CINZA,
                             alignment=TA_CENTER, spaceBefore=0)

    story = []

    # ── HEADER ────────────────────────────────────────────────────────────────
    if titulo:
        hdata = [[Paragraph(titulo, s_tit)],
                 [Paragraph(now_str(), s_dat)]]
        htbl = Table(hdata, colWidths=[W])
        htbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), AZUL),
            ("TOPPADDING",    (0,0), (0,0),  18),
            ("BOTTOMPADDING", (0,0), (0,0),  4),
            ("TOPPADDING",    (0,1), (0,1),  2),
            ("BOTTOMPADDING", (0,1), (0,1),  14),
            ("LEFTPADDING",   (0,0), (-1,-1), 20),
            ("RIGHTPADDING",  (0,0), (-1,-1), 20),
        ]))
        # Barra verde embaixo do header
        bar = Table([[""]], colWidths=[W])
        bar.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), VERDE),
            ("TOPPADDING",    (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ]))
        story.append(htbl)
        story.append(bar)
        story.append(Spacer(1, 10))

    # ── PARSER ────────────────────────────────────────────────────────────────
    lines = content.split("\n")
    i = 0
    # Remove título duplicado (primeira linha que igual ao título)
    titulo_clean = md_strip(titulo).strip().lower() if titulo else ""

    while i < len(lines):
        line = lines[i]
        line_clean = md_strip(re.sub(r'^#+\s*', '', line)).strip().lower()

        # Ignora imagens markdown
        if is_image_md(line):
            i += 1; continue

        # Ignora linha que é igual ao título (evita duplicata)
        if titulo_clean and line_clean == titulo_clean and line.startswith('#'):
            i += 1; continue

        # Ignora links de sumário (bullet com link markdown)
        if re.match(r'^[-*●]\s*\[', line.strip()):
            i += 1; continue

        # Converte links inline em texto simples
        line = md_strip_links(line)

        # Tabela markdown
        if "|" in line and i+1 < len(lines) and re.match(r'^\|[-| :]+\|', lines[i+1]):
            headers = [md_strip(c.strip()) for c in line.split("|") if c.strip()]
            i += 2
            rows = []
            while i < len(lines) and "|" in lines[i]:
                r = [md_strip(c.strip()) for c in lines[i].split("|") if c.strip()]
                if r: rows.append(r)
                i += 1
            if headers:
                ncols = len(headers)
                # Calcula largura proporcional ao conteúdo
                all_rows = [headers] + rows
                col_lens = []
                for j in range(ncols):
                    max_len = max((len(str(r[j])) if j < len(r) else 0 for r in all_rows), default=10)
                    col_lens.append(max(max_len, 6))
                total_len = sum(col_lens)
                col_widths = [W * (cl / total_len) for cl in col_lens]

                # Tabelas largas: fonte menor e padding reduzido
                is_wide = ncols > 6
                fsize_hdr  = 7 if is_wide else 9
                fsize_body = 7 if is_wide else 9
                pad        = 3 if is_wide else 5

                tdata = [headers] + rows
                tbl = Table(tdata, colWidths=col_widths, repeatRows=1)
                tbl.setStyle(TableStyle([
                    # Header
                    ("BACKGROUND",     (0,0), (-1,0), AZUL2),
                    ("TEXTCOLOR",      (0,0), (-1,0), colors.white),
                    ("FONTNAME",       (0,0), (-1,0), "Helvetica-Bold"),
                    ("FONTSIZE",       (0,0), (-1,0), fsize_hdr),
                    ("ALIGN",          (0,0), (-1,0), "CENTER"),
                    ("LINEBELOW",      (0,0), (-1,0), 2, VERDE),
                    # Body
                    ("FONTSIZE",       (0,1), (-1,-1), fsize_body),
                    ("TEXTCOLOR",      (0,1), (-1,-1), colors.HexColor("#2c3347")),
                    ("ROWBACKGROUNDS", (0,1), (-1,-1), [CINZA_BG, colors.white]),
                    ("ALIGN",          (0,1), (-1,-1), "CENTER"),
                    ("WORDWRAP",       (0,0), (-1,-1), True),
                    # Bordas
                    ("GRID",           (0,0), (-1,-1), 0.4, BORDA),
                    ("BOX",            (0,0), (-1,-1), 1, AZUL2),
                    # Padding
                    ("TOPPADDING",     (0,0), (-1,-1), pad),
                    ("BOTTOMPADDING",  (0,0), (-1,-1), pad),
                    ("LEFTPADDING",    (0,0), (-1,-1), pad+2),
                    ("RIGHTPADDING",   (0,0), (-1,-1), pad+2),
                ]))
                story.append(Spacer(1, 6))
                story.append(tbl)
                story.append(Spacer(1, 8))
            continue

        # H1 — box colorido
        if line.startswith("# "):
            txt = md_strip(line[2:])
            hbox = Table([[Paragraph(txt, s_h1)]], colWidths=[W])
            hbox.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), AZUL),
                ("LEFTPADDING",   (0,0), (-1,-1), 12),
                ("RIGHTPADDING",  (0,0), (-1,-1), 12),
                ("TOPPADDING",    (0,0), (-1,-1), 7),
                ("BOTTOMPADDING", (0,0), (-1,-1), 7),
                ("LINEAFTER",     (0,0), (0,-1),  4, VERDE),
            ]))
            story.append(Spacer(1, 8))
            story.append(hbox)
            story.append(Spacer(1, 6))

        # H2 — linha esquerda colorida
        elif line.startswith("## "):
            txt = md_clean(line[3:])
            hbox = Table([[Paragraph(txt, s_h2)]], colWidths=[W])
            hbox.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), AZUL_BG),
                ("LEFTPADDING",   (0,0), (-1,-1), 10),
                ("TOPPADDING",    (0,0), (-1,-1), 5),
                ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                ("LINEBEFORE",    (0,0), (0,-1),  3, VERDE),
            ]))
            story.append(Spacer(1, 6))
            story.append(hbox)
            story.append(Spacer(1, 4))

        # H3 — sutil
        elif line.startswith("### "):
            txt = md_clean(line[4:])
            story.append(Spacer(1, 4))
            story.append(Paragraph("› " + txt, s_h3))
            story.append(HRFlowable(width="30%", thickness=1,
                                     color=BORDA, spaceAfter=3))

        # Bullets
        elif re.match(r'^[-*•] ', line):
            txt = md_clean(line[2:].strip())
            story.append(Paragraph(
                '<font color="#00c2a8">●</font> ' + txt, s_bul))

        # Numerados
        elif re.match(r'^\d+\. ', line):
            m = re.match(r'^(\d+)\. (.*)', line)
            if m:
                num_box = Table([[
                    Paragraph(f'<b>{m.group(1)}</b>',
                               ParagraphStyle("nb", fontSize=9, fontName="Helvetica-Bold",
                                               textColor=colors.white, alignment=TA_CENTER)),
                    Paragraph(md_clean(m.group(2)), s_bod)
                ]], colWidths=[16, W-16])
                num_box.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (0,0), AZUL2),
                    ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                    ("TOPPADDING",    (0,0), (-1,-1), 3),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 3),
                    ("LEFTPADDING",   (0,0), (0,0), 2),
                    ("RIGHTPADDING",  (0,0), (0,0), 2),
                    ("LEFTPADDING",   (1,0), (1,0), 6),
                    ("RIGHTPADDING",  (1,0), (1,0), 0),
                ]))
                story.append(num_box)
                story.append(Spacer(1, 2))

        # Separador
        elif re.match(r'^---+$', line.strip()):
            story.append(HRFlowable(width="100%", thickness=0.5,
                                     color=BORDA, spaceBefore=4, spaceAfter=4))

        # Linha vazia
        elif line.strip() == "":
            story.append(Spacer(1, 4))

        # Texto normal
        else:
            story.append(Paragraph(md_clean(line), s_bod))

        i += 1

    # ── RODAPÉ ────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 8))
    rod_data = [[Paragraph(f"Gerado por Master IA · {now_str()}", s_rod)]]
    rod_tbl = Table(rod_data, colWidths=[W])
    rod_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), CINZA_BG),
        ("TOPPADDING",    (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LINEBEFORE",    (0,0), (0,0),   3, VERDE),
    ]))
    story.append(rod_tbl)

    doc.build(story)
    buf.seek(0)
    return buf.read()


def gen_word(titulo, content):
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2.5)
        s.left_margin = s.right_margin = Cm(3)

    # Título
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(titulo)
    run.bold = True; run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(26, 58, 110)

    # Data
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(now_str())
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(136, 146, 164)

    # Linha separadora
    p3 = doc.add_paragraph()
    pPr = p3._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), '1a3a6e')
    pBdr.append(bottom)
    pPr.append(pBdr)
    doc.add_paragraph()

    # Conteúdo
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]

        # Tabela markdown
        if "|" in line and i+1 < len(lines) and re.match(r'^\|[-| :]+\|', lines[i+1]):
            headers = [c.strip() for c in line.split("|") if c.strip()]
            i += 2
            rows = []
            while i < len(lines) and "|" in lines[i]:
                r = [c.strip() for c in lines[i].split("|") if c.strip()]
                if r: rows.append(r)
                i += 1
            if headers:
                headers = [md_strip(h) for h in headers]
                rows = [[md_strip(c) for c in r] for r in rows]
                tbl = doc.add_table(rows=1+len(rows), cols=len(headers))
                tbl.style = "Table Grid"
                for j, h in enumerate(headers):
                    cell = tbl.rows[0].cells[j]
                    cell.text = h
                    run = cell.paragraphs[0].runs[0]
                    run.bold = True
                    run.font.color.rgb = RGBColor(255,255,255)
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), '1a3a6e')
                    shd.set(qn('w:val'), 'clear')
                    tcPr.append(shd)
                for ri, row in enumerate(rows):
                    for j, val in enumerate(row):
                        if j < len(tbl.rows[ri+1].cells):
                            tbl.rows[ri+1].cells[j].text = val
                doc.add_paragraph()
            continue

        if line.startswith("#### "):
            doc.add_paragraph(line[5:], style="Heading 3")
        elif line.startswith("# "):
            p = doc.add_paragraph(line[2:], style="Heading 1")
        elif line.startswith("## "):
            p = doc.add_paragraph(line[3:], style="Heading 2")
        elif line.startswith("### "):
            p = doc.add_paragraph(line[4:], style="Heading 3")
        elif re.match(r'^[-*•] ', line):
            doc.add_paragraph(line[2:].strip(), style="List Bullet")
        elif re.match(r'^\d+\. ', line):
            doc.add_paragraph(re.sub(r'^\d+\. ', '', line), style="List Number")
        elif line.strip() == "":
            doc.add_paragraph()
        else:
            p = doc.add_paragraph()
            for part in re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', line):
                if part.startswith("**") and part.endswith("**"):
                    p.add_run(part[2:-2]).bold = True
                elif part.startswith("*") and part.endswith("*"):
                    p.add_run(part[1:-1]).italic = True
                else:
                    p.add_run(part)
        i += 1

    # Rodapé
    doc.add_paragraph()
    pf = doc.add_paragraph()
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rf = pf.add_run(f"Gerado por Master IA · {now_str()}")
    rf.font.size = Pt(8)
    rf.font.color.rgb = RGBColor(160,160,160)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
#  GERADOR EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def gen_excel(titulo, content):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = (titulo[:30] if titulo else "Dados")

    # Estilos
    f_hdr  = Font(bold=True, color="FFFFFF", size=11)
    f_tit  = Font(bold=True, color="1a3a6e", size=13)
    f_h2   = Font(bold=True, color="2d5fa3", size=11)
    f_h3   = Font(bold=True, color="8892a4", size=10)
    f_body = Font(size=10)
    f_rod  = Font(color="AAAAAA", size=8, italic=True)

    az     = PatternFill("solid", fgColor="1a3a6e")
    az2    = PatternFill("solid", fgColor="2d5fa3")
    verde  = PatternFill("solid", fgColor="e6faf7")
    light  = PatternFill("solid", fgColor="f5f7ff")
    alt    = PatternFill("solid", fgColor="eef2ff")
    white  = PatternFill("solid", fgColor="FFFFFF")

    brd = Border(
        left=Side(style="thin",   color="d0d8f0"),
        right=Side(style="thin",  color="d0d8f0"),
        top=Side(style="thin",    color="d0d8f0"),
        bottom=Side(style="thin", color="d0d8f0")
    )

    MAX_COL = "H"  # até coluna H para merge

    def mset(r, val, font=None, fill=None, align=None):
        """Merge A:H e define célula com segurança"""
        try:
            ws.merge_cells(f"A{r}:{MAX_COL}{r}")
        except Exception:
            pass
        c = ws[f"A{r}"]
        c.value = val
        if font:  c.font = font
        if fill:  c.fill = fill
        if align: c.alignment = align
        else:     c.alignment = Alignment(wrap_text=True)

    row = 1

    # Cabeçalho
    mset(row, titulo.upper() if titulo else "", f_tit, light,
         Alignment(horizontal="center"))
    row += 1
    mset(row, now_str(), f_rod, align=Alignment(horizontal="center"))
    row += 2

    # Parser
    lines = content.split("\n")
    titulo_clean = md_strip(titulo).strip().lower() if titulo else ""
    i = 0

    while i < len(lines):
        line = lines[i]

        # Filtros
        if is_image_md(line): i += 1; continue
        if re.match(r'^[-*●]\s*\[', line.strip()): i += 1; continue
        line = md_strip_links(line)
        line_clean = md_strip(re.sub(r'^#+\s*', '', line)).strip().lower()
        if titulo_clean and line_clean == titulo_clean and line.startswith('#'):
            i += 1; continue

        # Tabela markdown
        if "|" in line and i+1 < len(lines) and re.match(r'^\|[-| :]+\|', lines[i+1]):
            headers = [md_strip(c.strip()) for c in line.split("|") if c.strip()]
            i += 2
            trows = []
            while i < len(lines) and "|" in lines[i]:
                r = [md_strip(c.strip()) for c in lines[i].split("|") if c.strip()]
                if r: trows.append(r)
                i += 1
            if headers:
                ncols = len(headers)
                for j, h in enumerate(headers):
                    cell = ws.cell(row=row, column=j+1, value=h)
                    cell.font = f_hdr; cell.fill = az
                    cell.alignment = Alignment(horizontal="center", wrap_text=True)
                    cell.border = brd
                row += 1
                for ri, dr in enumerate(trows):
                    bg = light if ri % 2 == 0 else white
                    for j, val in enumerate(dr[:ncols]):
                        cell = ws.cell(row=row, column=j+1, value=val)
                        cell.font = f_body; cell.fill = bg
                        cell.alignment = Alignment(wrap_text=True); cell.border = brd
                    row += 1
                row += 1
            continue

        # Headings e texto
        try:
            if line.startswith("# "):
                mset(row, md_strip(line[2:]).upper(), f_tit, light,
                     Alignment(horizontal="left"))
            elif line.startswith("## "):
                mset(row, "◆ " + md_strip(line[3:]), f_h2, alt)
            elif line.startswith("### "):
                mset(row, "› " + md_strip(line[4:]), f_h3, verde)
            elif re.match(r'^[-*•] ', line):
                mset(row, "• " + md_strip(line[2:].strip()), f_body)
            elif re.match(r'^\d+\. ', line):
                m = re.match(r'^(\d+)\. (.*)', line)
                mset(row, f"{m.group(1)}. {md_strip(m.group(2))}" if m else md_strip(line), f_body)
            elif line.strip() == "":
                row += 1; i += 1; continue
            else:
                mset(row, md_strip(line), f_body)
            row += 1
        except Exception:
            row += 1
        i += 1

    # Rodapé
    row += 1
    mset(row, f"Gerado por Master IA · {now_str()}", f_rod,
         align=Alignment(horizontal="center"))

    # Largura automática — ignora células mescladas
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
            ws.column_dimensions[col_letter].width = min(max(ml + 2, 12), 55)

    # Altura das linhas
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14

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

# Usuários fixos no código — sem arquivo JSON
USERS = {
    "italo":   {"senha": hash_pw("master123"), "nome": "Ítalo",   "token": ""},
    "willian": {"senha": hash_pw("master123"), "nome": "Willian", "token": ""},
    "augusto": {"senha": hash_pw("master123"), "nome": "Augusto", "token": ""},
    "trinid":  {"senha": hash_pw("master123"), "nome": "Trinid",  "token": ""},
    "rafael":  {"senha": hash_pw("master123"), "nome": "Rafael",  "token": ""},
    "diogo":   {"senha": hash_pw("master123"), "nome": "Diogo",   "token": ""},
    "bia":     {"senha": hash_pw("master123"), "nome": "Bia",     "token": ""},
}

def load_users(): return USERS
def save_users(u): USERS.update(u)
def seed_users(): pass

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
    return jsonify({"token": token, "nome": u["nome"], "username": user})

@app.route("/gerar", methods=["POST"])
def gerar():
    d        = request.get_json()
    tipo     = (d.get("tipo") or "pdf").lower()
    titulo   = d.get("titulo") or "Documento"
    conteudo = d.get("conteudo") or ""
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
            data = gen_pdf(titulo, conteudo)
            ext  = "pdf"
            mime = "application/pdf"

        fname = f"{safe_name(titulo)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        (DOCS_DIR / fname).write_bytes(data)

        return send_file(io.BytesIO(data), as_attachment=True,
                         download_name=f"{safe_name(titulo)}.{ext}", mimetype=mime)
    except Exception as e:
        import traceback
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500

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
