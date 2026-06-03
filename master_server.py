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
    registrar_evento("login", user)
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
def tool_grafico_barras(labels, valores, titulo="", xlabel="", ylabel="", horizontal=False):
    fig, ax = plt.subplots(figsize=(7, 4))
    _apply_master_style(fig, [ax])
    cores = MASTER_COLORS[:len(labels)]
    y = np.arange(len(labels))
    if horizontal:
        bars = ax.barh(y, valores, color=cores, edgecolor="#2a3045", linewidth=0.5)
        ax.set_yticks(y); ax.set_yticklabels(labels)
        for bar, val in zip(bars, valores):
            ax.text(bar.get_width() + max(valores)*0.01, bar.get_y()+bar.get_height()/2,
                    f"{val:,.2f}".replace(",","X").replace(".",",").replace("X","."),
                    va="center", ha="left", color="#e8eaf0", fontsize=8)
    else:
        bars = ax.bar(y, valores, color=cores, edgecolor="#2a3045", linewidth=0.5, width=0.65)
        ax.set_xticks(y); ax.set_xticklabels(labels, rotation=20, ha="right")
        for bar, val in zip(bars, valores):
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+max(valores)*0.01,
                    f"{val:,.2f}".replace(",","X").replace(".",",").replace("X","."),
                    ha="center", va="bottom", color="#e8eaf0", fontsize=8)
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10)
    if xlabel: ax.set_xlabel(xlabel)
    if ylabel: ax.set_ylabel(ylabel)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f"{x:,.0f}".replace(",",".")))
    fig.tight_layout()
    return _fig_to_b64(fig)


# ── Gráfico de linhas ──────────────────────────────────────────────────────────
def tool_grafico_linhas(series, titulo="", xlabel="", ylabel=""):
    """series: [{"nome": str, "dados": [float,...], "labels": [str,...]}]"""
    fig, ax = plt.subplots(figsize=(7, 4))
    _apply_master_style(fig, [ax])
    for i, s in enumerate(series):
        cor = MASTER_COLORS[i % len(MASTER_COLORS)]
        xs = range(len(s["dados"]))
        ax.plot(xs, s["dados"], color=cor, linewidth=2, marker="o", markersize=5, label=s["nome"])
        ax.fill_between(xs, s["dados"], alpha=0.07, color=cor)
    if series and series[0].get("labels"):
        ax.set_xticks(range(len(series[0]["labels"])))
        ax.set_xticklabels(series[0]["labels"], rotation=20, ha="right")
    if len(series) > 1:
        leg = ax.legend(fontsize=8, facecolor="#1e2330", edgecolor="#2a3045", labelcolor="#e8eaf0")
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10)
    if xlabel: ax.set_xlabel(xlabel)
    if ylabel: ax.set_ylabel(ylabel)
    fig.tight_layout()
    return _fig_to_b64(fig)


# ── Gráfico de pizza ───────────────────────────────────────────────────────────
def tool_grafico_pizza(labels, valores, titulo=""):
    fig, ax = plt.subplots(figsize=(6, 5))
    _apply_master_style(fig, [ax])
    wedges, texts, autotexts = ax.pie(
        valores, labels=None,
        colors=MASTER_COLORS[:len(labels)],
        autopct="%1.1f%%", startangle=90,
        wedgeprops=dict(edgecolor="#181c26", linewidth=1.5),
        pctdistance=0.82
    )
    for at in autotexts:
        at.set_color("#e8eaf0"); at.set_fontsize(8)
    patches = [mpatches.Patch(color=MASTER_COLORS[i % len(MASTER_COLORS)], label=l)
               for i, l in enumerate(labels)]
    ax.legend(handles=patches, loc="lower center", bbox_to_anchor=(0.5, -0.12),
              ncol=3, fontsize=8, facecolor="#1e2330",
              edgecolor="#2a3045", labelcolor="#e8eaf0", framealpha=0.9)
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", color="#e8eaf0", pad=10)
    fig.tight_layout()
    return _fig_to_b64(fig)


# ── Gráfico de dispersão ───────────────────────────────────────────────────────
def tool_grafico_dispersao(series, titulo="", xlabel="", ylabel=""):
    """series: [{"nome": str, "x": [float], "y": [float]}]"""
    fig, ax = plt.subplots(figsize=(7, 4))
    _apply_master_style(fig, [ax])
    for i, s in enumerate(series):
        cor = MASTER_COLORS[i % len(MASTER_COLORS)]
        ax.scatter(s["x"], s["y"], color=cor, label=s["nome"], alpha=0.85, s=50, edgecolors="#181c26", linewidth=0.5)
    if len(series) > 1:
        ax.legend(fontsize=8, facecolor="#1e2330", edgecolor="#2a3045", labelcolor="#e8eaf0")
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10)
    if xlabel: ax.set_xlabel(xlabel)
    if ylabel: ax.set_ylabel(ylabel)
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
    ws.title = titulo[:28] if titulo else "Dados"

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

    system = (
        "Você é o Master IA, assistente especializado em contabilidade, fiscal e tributário brasileiro. "
        "Responda sempre em português. "
        "Quando o usuário pedir gráficos, tabelas, planilhas ou código Python, USE as ferramentas disponíveis — "
        "não apenas descreva, EXECUTE. "
        "Para dados fiscais/contábeis, prefira gráficos de barras ou linhas. "
        "Ao gerar Excel, inclua gráfico quando fizer sentido. "
        "Para código Python solicitado pelo usuário, execute e mostre o resultado. "
        "Após usar uma ferramenta, comente o resultado brevemente em português."
    )

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
                "max_tokens": 4096,
                "system": system,
                "tools": TOOLS,
                "messages": messages,
            },
            timeout=120
        )
        if not resp.ok:
            return jsonify({"erro": resp.text}), resp.status_code
        data = resp.json()
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

    # Processa blocos de resposta
    result_blocks = []
    tool_calls_made = []

    for block in data.get("content", []):
        btype = block.get("type")
        if btype == "text":
            result_blocks.append({"tipo": "texto", "conteudo": block.get("text","")})

        elif btype == "tool_use":
            tool_name  = block.get("name","")
            tool_input = block.get("input", {})
            tool_calls_made.append({"nome": tool_name, "input": tool_input})

            try:
                if tool_name == "gerar_grafico_barras":
                    img = tool_grafico_barras(
                        tool_input["labels"], tool_input["valores"],
                        tool_input.get("titulo",""), tool_input.get("xlabel",""),
                        tool_input.get("ylabel",""), tool_input.get("horizontal", False)
                    )
                    result_blocks.append({"tipo": "imagem", "b64": img,
                                          "legenda": tool_input.get("titulo","Gráfico")})

                elif tool_name == "gerar_grafico_linhas":
                    img = tool_grafico_linhas(
                        tool_input["series"],
                        tool_input.get("titulo",""), tool_input.get("xlabel",""),
                        tool_input.get("ylabel","")
                    )
                    result_blocks.append({"tipo": "imagem", "b64": img,
                                          "legenda": tool_input.get("titulo","Gráfico de Linhas")})

                elif tool_name == "gerar_grafico_pizza":
                    img = tool_grafico_pizza(
                        tool_input["labels"], tool_input["valores"],
                        tool_input.get("titulo","")
                    )
                    result_blocks.append({"tipo": "imagem", "b64": img,
                                          "legenda": tool_input.get("titulo","Gráfico de Pizza")})

                elif tool_name == "gerar_grafico_dispersao":
                    img = tool_grafico_dispersao(
                        tool_input["series"],
                        tool_input.get("titulo",""), tool_input.get("xlabel",""),
                        tool_input.get("ylabel","")
                    )
                    result_blocks.append({"tipo": "imagem", "b64": img,
                                          "legenda": tool_input.get("titulo","Dispersão")})

                elif tool_name == "executar_python":
                    res = tool_executar_python(tool_input["codigo"])
                    result_blocks.append({
                        "tipo": "codigo_resultado",
                        "codigo": tool_input["codigo"],
                        "stdout": res["stdout"],
                        "stderr": res["stderr"],
                        "erro":   res["erro"],
                        "imagem_b64": res.get("imagem_b64")
                    })

                elif tool_name == "gerar_excel_avancado":
                    xls_bytes = tool_excel_avancado(
                        tool_input["titulo"],
                        tool_input["colunas"],
                        tool_input["linhas"],
                        tool_input.get("grafico_tipo"),
                        tool_input.get("grafico_series")
                    )
                    xls_b64 = base64.b64encode(xls_bytes).decode()
                    fname   = safe_name(tool_input["titulo"]) + ".xlsx"
                    result_blocks.append({"tipo": "excel", "b64": xls_b64,
                                          "nome": fname,
                                          "legenda": tool_input["titulo"]})
                    registrar_evento("documento", user, "excel_avancado")

            except Exception as e:
                result_blocks.append({"tipo": "erro_ferramenta",
                                       "ferramenta": tool_name,
                                       "mensagem": str(e),
                                       "trace": tb.format_exc()})

    registrar_evento("mensagem", user)
    return jsonify({"blocos": result_blocks, "stop_reason": data.get("stop_reason","")})

