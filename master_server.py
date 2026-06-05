"""
Master IA — Servidor Completo v3.0
Capacidades: OCR, PDF, Excel, Word, Python, Gráficos, Dashboards, Documentos, Fiscal
"""

import os, io, json, re, hashlib, secrets, base64, subprocess, sys
import textwrap, traceback as tb, threading, contextlib
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# ── Docs
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.platypus.flowables import HRFlowable

# ── Viz
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ── Data
import pandas as pd

# ── HTTP
import requests as _requests

app = Flask(__name__)
CORS(app)

DATA_DIR = Path(os.environ.get("DATA_DIR", str(Path(__file__).parent / "master_data")))
DATA_DIR.mkdir(exist_ok=True, parents=True)
DOCS_DIR = DATA_DIR / "documentos"
DOCS_DIR.mkdir(exist_ok=True)

# ══════════════════════════════════════════════════════════════════
#  HELPERS GERAIS
# ══════════════════════════════════════════════════════════════════
def now_str():
    return datetime.now().strftime("%d/%m/%Y às %H:%M")

def safe_name(s):
    s = re.sub(r'[^\w\s-]', '', str(s), flags=re.UNICODE)
    return re.sub(r'\s+', '_', s.strip())[:50]

def hash_pw(p): return hashlib.sha256(p.encode()).hexdigest()

# ── Usuários ──────────────────────────────────────────────────────
USERS_FILE = DATA_DIR / "users.json"

def load_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text())
        except: pass
    default = {"italo": {"senha": hash_pw("master9713"), "nome": "Ítalo", "token": ""}}
    save_users(default); return default

def save_users(u): USERS_FILE.write_text(json.dumps(u, ensure_ascii=False, indent=2))

def auth(req):
    tok = req.headers.get("X-Token","")
    if not tok: return None
    for u, d in load_users().items():
        if d.get("token") == tok: return u
    return None

def registrar_evento(tipo, usuario, detalhe=""):
    log = DATA_DIR / "eventos.jsonl"
    with open(log, "a") as f:
        f.write(json.dumps({"tipo":tipo,"usuario":usuario,"detalhe":detalhe,
                             "ts":datetime.now().isoformat()}) + "\n")

# ══════════════════════════════════════════════════════════════════
#  ROTAS BÁSICAS
# ══════════════════════════════════════════════════════════════════
@app.route("/", methods=["GET"])
def index():
    html = Path(__file__).parent / "master_chat.html"
    if html.exists():
        return html.read_text(encoding="utf-8")
    return jsonify({"status":"Master IA v3.0 online","ts":now_str()})

@app.route("/status", methods=["GET"])
def status(): return jsonify({"status":"Master IA v3.0 online","ts":now_str()})

@app.route("/health", methods=["GET"])
def health(): return jsonify({"ok":True})

@app.route("/login", methods=["POST"])
def login():
    d = request.get_json() or {}
    username = (d.get("username") or "").strip().lower()
    senha    = (d.get("senha") or "").strip()
    users = load_users()
    u = users.get(username)
    if not u or u.get("senha") != hash_pw(senha):
        return jsonify({"erro":"Usuário ou senha inválidos"}), 401
    token = secrets.token_hex(32)
    users[username]["token"] = token
    save_users(users)
    registrar_evento("login", username)
    return jsonify({"token": token, "nome": u.get("nome", username), "username": username})

@app.route("/chats", methods=["GET"])
def list_chats():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    d = DATA_DIR / "chats" / user
    d.mkdir(parents=True, exist_ok=True)
    chats = []
    for f in sorted(d.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
        try:
            data = json.loads(f.read_text())
            chats.append({"id": f.stem, "titulo": data.get("titulo","Conversa"),
                          "ts": data.get("ts", "")})
        except: pass
    return jsonify(chats)

@app.route("/chats", methods=["POST"])
def save_chat():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    d_req = request.get_json() or {}
    cid   = d_req.get("id") or secrets.token_hex(8)
    d = DATA_DIR / "chats" / user
    d.mkdir(parents=True, exist_ok=True)
    (d / f"{cid}.json").write_text(json.dumps({
        "id": cid, "titulo": d_req.get("titulo","Conversa"),
        "messages": d_req.get("messages",[]),
        "ts": now_str()
    }, ensure_ascii=False))
    return jsonify({"ok":True, "id": cid})

@app.route("/chats/<cid>", methods=["GET"])
def get_chat(cid):
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    f = DATA_DIR / "chats" / user / f"{cid}.json"
    if not f.exists(): return jsonify({"erro":"Não encontrado"}), 404
    return jsonify(json.loads(f.read_text()))

@app.route("/chats/<cid>", methods=["DELETE"])
def del_chat(cid):
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    f = DATA_DIR / "chats" / user / f"{cid}.json"
    if f.exists(): f.unlink()
    return jsonify({"ok":True})

# ══════════════════════════════════════════════════════════════════
#  ADMIN
# ══════════════════════════════════════════════════════════════════
@app.route("/admin/listar", methods=["GET"])
def admin_listar():
    user = auth(request)
    if user != "italo": return jsonify({"erro":"Acesso negado"}), 403
    users = load_users()
    return jsonify([{"username":u,"nome":d.get("nome",u),"tem_token":bool(d.get("token"))}
                    for u,d in users.items()])

@app.route("/admin/adicionar", methods=["POST"])
def admin_adicionar():
    user = auth(request)
    if user != "italo": return jsonify({"erro":"Acesso negado"}), 403
    d = request.get_json()
    username = (d.get("username") or "").strip().lower()
    nome     = (d.get("nome") or "").strip()
    senha    = (d.get("senha") or "master123").strip()
    if not username or not nome: return jsonify({"erro":"username e nome obrigatórios"}), 400
    users = load_users()
    if username in users: return jsonify({"erro":"Usuário já existe"}), 409
    users[username] = {"senha": hash_pw(senha), "nome": nome, "token": ""}
    save_users(users)
    return jsonify({"ok":True})

@app.route("/admin/resetar_senha", methods=["POST"])
def admin_resetar():
    user = auth(request)
    if user != "italo": return jsonify({"erro":"Acesso negado"}), 403
    d = request.get_json()
    alvo = (d.get("username") or "").strip().lower()
    nova = (d.get("nova_senha") or "master123").strip()
    users = load_users()
    if alvo not in users: return jsonify({"erro":"Usuário não encontrado"}), 404
    users[alvo]["senha"] = hash_pw(nova)
    save_users(users)
    return jsonify({"ok":True})

@app.route("/admin/remover", methods=["POST"])
def admin_remover():
    user = auth(request)
    if user != "italo": return jsonify({"erro":"Acesso negado"}), 403
    d = request.get_json()
    alvo = (d.get("username") or "").strip().lower()
    if alvo == "italo": return jsonify({"erro":"Não pode remover o admin"}), 400
    users = load_users()
    if alvo not in users: return jsonify({"erro":"Usuário não encontrado"}), 404
    del users[alvo]
    save_users(users)
    return jsonify({"ok":True})

@app.route("/dashboard/stats", methods=["GET"])
def dashboard_stats():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    log = DATA_DIR / "eventos.jsonl"
    events = []
    if log.exists():
        for line in log.read_text().strip().split("\n"):
            try: events.append(json.loads(line))
            except: pass
    total_msgs = sum(1 for e in events if e["tipo"]=="mensagem")
    total_docs = sum(1 for e in events if e["tipo"]=="documento")
    total_logins = sum(1 for e in events if e["tipo"]=="login")
    usuarios = set(e["usuario"] for e in events if e.get("usuario"))
    ultimos = [{"tipo":e["tipo"],"usuario":e["usuario"],"detalhe":e.get("detalhe",""),"ts":e["ts"][11:16]}
               for e in reversed(events[-30:])]
    return jsonify({"totais":{"mensagens":total_msgs,"documentos":total_docs,
                              "logins":total_logins,"usuarios_ativos":len(usuarios)},
                    "ultimos":ultimos,"usuario_logado":user,"is_admin":user=="italo"})


# ══════════════════════════════════════════════════════════════════
#  OCR — LER IMAGENS E PDFs
# ══════════════════════════════════════════════════════════════════
def ocr_imagem(img_b64: str, lang: str = "por") -> str:
    """OCR em imagem base64. Retorna texto extraído."""
    try:
        import pytesseract
        from PIL import Image
        img_data = base64.b64decode(img_b64)
        img = Image.open(io.BytesIO(img_data))
        # Pré-processamento: converte para RGB se necessário
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        texto = pytesseract.image_to_string(img, lang=lang,
                    config="--oem 3 --psm 6")
        return texto.strip()
    except ImportError:
        return "[OCR indisponível: pytesseract não instalado]"
    except Exception as e:
        return f"[Erro OCR: {e}]"

def ocr_pdf(pdf_b64: str, lang: str = "por", max_pages: int = 10) -> str:
    """OCR em PDF base64. Converte páginas em imagens e extrai texto."""
    try:
        import pytesseract
        from PIL import Image
        try:
            from pdf2image import convert_from_bytes
        except ImportError:
            return "[pdf2image não instalado — não é possível fazer OCR em PDF]"

        pdf_data = base64.b64decode(pdf_b64)
        pages = convert_from_bytes(pdf_data, dpi=200, first_page=1, last_page=max_pages)
        textos = []
        for i, page in enumerate(pages, 1):
            txt = pytesseract.image_to_string(page, lang=lang, config="--oem 3 --psm 6")
            textos.append(f"--- Página {i} ---\n{txt.strip()}")
        return "\n\n".join(textos)
    except ImportError:
        return "[pytesseract não instalado]"
    except Exception as e:
        return f"[Erro OCR PDF: {e}]"

def ler_pdf_texto(pdf_b64: str) -> str:
    """Extrai texto de PDF com texto selecionável (sem OCR)."""
    try:
        import pdfplumber
        pdf_data = base64.b64decode(pdf_b64)
        with pdfplumber.open(io.BytesIO(pdf_data)) as pdf:
            textos = []
            for i, page in enumerate(pdf.pages, 1):
                txt = page.extract_text() or ""
                if txt.strip():
                    textos.append(f"--- Página {i} ---\n{txt.strip()}")
            return "\n\n".join(textos) if textos else ""
    except ImportError:
        # Fallback com PyPDF2
        try:
            import PyPDF2
            pdf_data = base64.b64decode(pdf_b64)
            reader = PyPDF2.PdfReader(io.BytesIO(pdf_data))
            textos = []
            for i, page in enumerate(reader.pages, 1):
                txt = page.extract_text() or ""
                if txt.strip():
                    textos.append(f"--- Página {i} ---\n{txt.strip()}")
            return "\n\n".join(textos)
        except:
            return ""
    except Exception as e:
        return f"[Erro ao ler PDF: {e}]"

def ler_docx(docx_b64: str) -> str:
    """Extrai texto de documento Word."""
    try:
        docx_data = base64.b64decode(docx_b64)
        doc = Document(io.BytesIO(docx_data))
        partes = []
        for para in doc.paragraphs:
            if para.text.strip():
                partes.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                linha = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                if linha: partes.append(linha)
        return "\n".join(partes)
    except Exception as e:
        return f"[Erro ao ler Word: {e}]"

def ler_excel(xlsx_b64: str) -> str:
    """Extrai conteúdo de planilha Excel."""
    try:
        data = base64.b64decode(xlsx_b64)
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        resultado = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            resultado.append(f"=== Aba: {sheet_name} ===")
            for row in ws.iter_rows(values_only=True):
                valores = [str(v) if v is not None else "" for v in row]
                if any(v.strip() for v in valores):
                    resultado.append(";".join(valores))
        return "\n".join(resultado)
    except Exception as e:
        return f"[Erro ao ler Excel: {e}]"


# ══════════════════════════════════════════════════════════════════
#  GERADOR PDF PROFISSIONAL
# ══════════════════════════════════════════════════════════════════
import re as _re

def _pdf_parse_md(content):
    blocks = []
    lines = content.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if _re.match(r'^!\[.*\]\(.*\)$', stripped): i += 1; continue
        if '|' in stripped and i+1 < len(lines) and _re.match(r'^\|[-| :]+\|', lines[i+1].strip()):
            headers = [c.strip() for c in stripped.split('|') if c.strip()]
            i += 2
            rows = []
            while i < len(lines) and '|' in lines[i]:
                r = [c.strip() for c in lines[i].split('|') if c.strip()]
                if r: rows.append(r)
                i += 1
            blocks.append({'type':'table','headers':headers,'rows':rows}); continue
        m = _re.match(r'^(#{1,4})\s+(.+)', stripped)
        if m:
            level = len(m.group(1))
            text = _re.sub(r'\*\*(.+?)\*\*', r'\1', m.group(2).strip())
            text = _re.sub(r'\*(.+?)\*', r'\1', text)
            blocks.append({'type':f'h{level}','text':text}); i += 1; continue
        if _re.match(r'^---+$', stripped): blocks.append({'type':'hr'}); i += 1; continue
        if _re.match(r'^[-*•]\s+', stripped):
            items = []
            while i < len(lines) and _re.match(r'^[-*•]\s+', lines[i].strip()):
                items.append(lines[i].strip()[2:].strip()); i += 1
            blocks.append({'type':'bullets','items':items}); continue
        if _re.match(r'^\d+\.\s+', stripped):
            items = []
            while i < len(lines) and _re.match(r'^\d+\.\s+', lines[i].strip()):
                items.append(_re.sub(r'^\d+\.\s+', '', lines[i].strip())); i += 1
            blocks.append({'type':'numbered','items':items}); continue
        if not stripped: blocks.append({'type':'space'}); i += 1; continue
        para_lines = []
        while i < len(lines):
            l = lines[i].strip()
            if (not l or l.startswith('#') or '|' in l and i+1 < len(lines) and _re.match(r'^\|[-| :]+\|', lines[i+1].strip() if i+1 < len(lines) else '') or
                _re.match(r'^[-*•]\s+', l) or _re.match(r'^\d+\.\s+', l) or _re.match(r'^---+$', l)):
                break
            para_lines.append(l); i += 1
        text = ' '.join(para_lines)
        if text.strip(): blocks.append({'type':'para','text':text})
    return blocks

def _md_to_rl(text):
    text = _re.sub(r'&', '&amp;', text)
    text = _re.sub(r'<(?!/?[biu])', '&lt;', text)
    text = _re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    text = _re.sub(r'\*(.+?)\*',     r'<i>\1</i>', text)
    text = _re.sub(r'`(.+?)`',       r'<font name="Courier" size="9">\1</font>', text)
    text = _re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    return text

def gen_pdf(titulo, content, imagens=None):
    buf = io.BytesIO()
    C_DARK   = colors.HexColor("#1e2532")
    C_NAVY   = colors.HexColor("#1a365d")
    C_ACCENT = colors.HexColor("#2b6cb0")
    C_ACCENT2= colors.HexColor("#3182ce")
    C_LIGHT  = colors.HexColor("#ebf8ff")
    C_BG     = colors.HexColor("#f7fafc")
    C_BORDER = colors.HexColor("#bee3f8")
    C_MUTED  = colors.HexColor("#718096")
    C_STRIPE = colors.HexColor("#f0f7ff")
    C_TEXT   = colors.HexColor("#2d3748")
    C_WHITE  = colors.white
    PAGE_W, PAGE_H = A4
    LM = RM = 2.2*cm; TM = 2.0*cm; BM = 2.2*cm
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=TM, bottomMargin=BM,
                             leftMargin=LM, rightMargin=RM, title=titulo)
    W = doc.width
    def sty(name, **kw): return ParagraphStyle(name, **kw)
    s_cv_title = sty('cvt', fontSize=24, fontName='Helvetica-Bold', textColor=C_WHITE,
                     alignment=TA_CENTER, leading=32, spaceAfter=4)
    s_cv_sub   = sty('cvs', fontSize=10, textColor=colors.HexColor("#90cdf4"),
                     alignment=TA_CENTER, leading=15)
    s_h1  = sty('h1p', fontSize=13, fontName='Helvetica-Bold', textColor=C_NAVY,
                spaceBefore=16, spaceAfter=6, leading=18, keepWithNext=1)
    s_h2  = sty('h2p', fontSize=11.5, fontName='Helvetica-Bold', textColor=C_ACCENT,
                spaceBefore=12, spaceAfter=4, leading=16, keepWithNext=1)
    s_h3  = sty('h3p', fontSize=10.5, fontName='Helvetica-Bold', textColor=C_TEXT,
                spaceBefore=8, spaceAfter=3, leading=15, keepWithNext=1)
    s_h4  = sty('h4p', fontSize=10, fontName='Helvetica-BoldOblique', textColor=C_MUTED,
                spaceBefore=6, spaceAfter=2, leading=14)
    s_body = sty('bdyp', fontSize=10, leading=17, spaceAfter=7,
                 alignment=TA_JUSTIFY, textColor=C_TEXT)
    s_bul  = sty('bulp', fontSize=10, leading=15, leftIndent=18, spaceAfter=3, textColor=C_TEXT)
    s_num  = sty('nump', fontSize=10, leading=15, leftIndent=22, firstLineIndent=-14,
                 spaceAfter=3, textColor=C_TEXT)
    s_foot = sty('footp', fontSize=7.5, textColor=C_MUTED, alignment=TA_CENTER)
    story = []
    if titulo:
        capa = Table([[Paragraph(titulo, s_cv_title)],[Paragraph(now_str(), s_cv_sub)]],
                     colWidths=[W])
        capa.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),C_NAVY),
            ('TOPPADDING',(0,0),(0,0),28),('BOTTOMPADDING',(0,0),(0,0),8),
            ('TOPPADDING',(0,1),(0,1),4),('BOTTOMPADDING',(0,1),(0,1),24),
            ('LEFTPADDING',(0,0),(-1,-1),20),('RIGHTPADDING',(0,0),(-1,-1),20),
        ]))
        story.append(capa)
        bar = Table([['']], colWidths=[W], rowHeights=[5])
        bar.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),C_ACCENT2),
                                  ('TOPPADDING',(0,0),(-1,-1),0),
                                  ('BOTTOMPADDING',(0,0),(-1,-1),0)]))
        story.append(bar)
        story.append(Spacer(1, 0.4*cm))
    blocks = _pdf_parse_md(content)
    headings = [(b['type'],b['text']) for b in blocks
                if b['type'] in ('h1','h2') and b.get('text','').strip()
                and b['text'].strip().lower() != (titulo or '').strip().lower()]
    if len(headings) >= 3:
        toc_h = sty('toch', fontSize=8, fontName='Helvetica-Bold', textColor=C_MUTED, spaceAfter=4)
        toc1  = sty('toc1', fontSize=10, textColor=C_NAVY, leading=18, fontName='Helvetica-Bold', spaceAfter=2)
        toc2  = sty('toc2', fontSize=9.5, textColor=C_ACCENT, leading=16, leftIndent=16, spaceAfter=1)
        toc_rows = [[Paragraph('SUMÁRIO', toc_h)]]
        for h, t in headings:
            prefix = '▶  ' if h == 'h1' else '     · '
            toc_rows.append([Paragraph(prefix + t, toc1 if h=='h1' else toc2)])
        toc = Table(toc_rows, colWidths=[W])
        toc.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),C_BG),
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#e2e8f0")),
            ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
            ('LEFTPADDING',(0,0),(-1,-1),14),('RIGHTPADDING',(0,0),(-1,-1),14),
            ('BOX',(0,0),(-1,-1),0.5,C_BORDER),
        ]))
        story.append(toc)
        story.append(Spacer(1, 0.5*cm))
    for block in blocks:
        btype = block.get('type'); txt = block.get('text','')
        if btype in ('h1',) and txt.strip().lower() == (titulo or '').strip().lower(): continue
        if btype == 'h1':
            box = Table([[Paragraph(_md_to_rl(txt), s_h1)]], colWidths=[W])
            box.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,-1),C_LIGHT),('LINEBEFORE',(0,0),(0,-1),5,C_ACCENT2),
                ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
                ('LEFTPADDING',(0,0),(-1,-1),12),('RIGHTPADDING',(0,0),(-1,-1),12),
            ]))
            story.append(Spacer(1,6)); story.append(box); story.append(Spacer(1,4))
        elif btype == 'h2':
            story.append(Spacer(1,4))
            story.append(Paragraph(_md_to_rl(txt), s_h2))
            story.append(HRFlowable(width='35%', thickness=2, color=C_ACCENT2, spaceAfter=4))
        elif btype == 'h3': story.append(Spacer(1,2)); story.append(Paragraph(_md_to_rl(txt), s_h3))
        elif btype == 'h4': story.append(Paragraph(_md_to_rl(txt), s_h4))
        elif btype == 'para': story.append(Paragraph(_md_to_rl(txt), s_body))
        elif btype == 'bullets':
            for item in block['items']:
                story.append(Paragraph('<font color="#2b6cb0" size="12">•</font>  ' + _md_to_rl(item), s_bul))
        elif btype == 'numbered':
            for n, item in enumerate(block['items'], 1):
                story.append(Paragraph(f'<b><font color="#2b6cb0">{n}.</font></b>  ' + _md_to_rl(item), s_num))
        elif btype == 'table':
            headers = block.get('headers',[]); rows = block.get('rows',[])
            if not headers: continue
            ncols = len(headers)
            def _cc(v):
                v = str(v)
                v = _re.sub(r'\*\*(.+?)\*\*', r'\1', v)
                v = _re.sub(r'\*(.+?)\*', r'\1', v)
                return v.strip()
            all_rows = [[_cc(c) for c in headers]] + [[_cc(c) for c in r] for r in rows]
            col_lens = [max(len(all_rows[ri][j]) if j<len(all_rows[ri]) else 3
                            for ri in range(len(all_rows))) for j in range(ncols)]
            total = sum(col_lens) or 1
            col_widths = [max(W*cl/total, W*0.08) for cl in col_lens]
            sw = sum(col_widths); col_widths = [w*W/sw for w in col_widths]
            is_wide = ncols > 5; fs = 7.5 if is_wide else 9; pad = 3 if is_wide else 5
            uid = str(abs(hash(str(block))))[-6:]
            s_th = sty('TH'+uid, fontSize=fs, fontName='Helvetica-Bold', textColor=C_WHITE,
                       alignment=TA_CENTER, leading=fs+3)
            s_td = sty('TD'+uid, fontSize=fs, leading=fs+5, textColor=C_TEXT, alignment=TA_LEFT)
            s_tn = sty('TN'+uid, fontSize=fs, leading=fs+5, textColor=C_TEXT, alignment=TA_RIGHT)
            def _is_num(v):
                v2 = v.replace('.','').replace(',','.').replace('R$','').replace('%','').replace(' ','')
                try: float(v2); return True
                except: return False
            num_cols = set()
            for j in range(ncols):
                vals = [all_rows[ri][j] for ri in range(1,len(all_rows)) if j<len(all_rows[ri])]
                if vals and sum(1 for v in vals if _is_num(v)) >= len(vals)*0.6:
                    num_cols.add(j)
            tdata = [[Paragraph(c, s_th) for c in all_rows[0]]]
            for ri in range(1, len(all_rows)):
                tdata.append([Paragraph(all_rows[ri][j] if j<len(all_rows[ri]) else '',
                                         s_tn if j in num_cols else s_td)
                               for j in range(ncols)])
            tbl = Table(tdata, colWidths=col_widths, repeatRows=1, hAlign='CENTER', splitByRow=True)
            tbl.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0),C_NAVY),
                ('LINEBELOW',(0,0),(-1,0),3,C_ACCENT2),
                ('ROWBACKGROUNDS',(0,1),(-1,-1),[C_WHITE,C_STRIPE]),
                ('GRID',(0,0),(-1,-1),0.5,C_BORDER),
                ('TOPPADDING',(0,0),(-1,-1),pad),('BOTTOMPADDING',(0,0),(-1,-1),pad),
                ('LEFTPADDING',(0,0),(-1,-1),pad+2),('RIGHTPADDING',(0,0),(-1,-1),pad+2),
            ]))
            story.append(Spacer(1,4)); story.append(tbl); story.append(Spacer(1,4))
        elif btype == 'hr': story.append(HRFlowable(width='100%', thickness=0.5, color=C_BORDER, spaceAfter=4))
        elif btype == 'space': story.append(Spacer(1, 0.2*cm))
    story.append(Spacer(1, 1*cm))
    rod = Paragraph(f"Gerado por Master IA · {now_str()}", s_foot)
    story.append(HRFlowable(width='100%', thickness=0.5, color=C_BORDER, spaceBefore=4, spaceAfter=4))
    story.append(rod)
    doc.build(story)
    buf.seek(0); return buf.read()


# ══════════════════════════════════════════════════════════════════
#  GERADOR WORD PROFISSIONAL
# ══════════════════════════════════════════════════════════════════
def _word_add_border_bottom(paragraph, color_hex, size):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom); pPr.append(pBdr)

def _word_set_cell_bg(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def gen_word(titulo, content):
    doc = Document()
    # Configurações da página
    for section in doc.sections:
        section.page_width  = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin = section.right_margin = Cm(2.5)
        section.top_margin = section.bottom_margin = Cm(2.5)
    # Cabeçalho
    if titulo:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(titulo.upper())
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0x1a, 0x36, 0x5d)
        _word_add_border_bottom(p, '3182ce', 8)
        p2 = doc.add_paragraph()
        run2 = p2.add_run(now_str())
        run2.font.size = Pt(8)
        run2.font.color.rgb = RGBColor(0x71, 0x80, 0x96)
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
    blocks = _pdf_parse_md(content)
    for block in blocks:
        btype = block.get('type'); txt = block.get('text','')
        if btype == 'h1':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(txt.upper())
            run.bold = True; run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0x1a, 0x36, 0x5d)
            _word_add_border_bottom(p, '3182ce', 8)
        elif btype == 'h2':
            p = doc.add_paragraph()
            run = p.add_run(txt)
            run.bold = True; run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(0x2b, 0x6c, 0xb0)
        elif btype == 'h3':
            p = doc.add_paragraph()
            run = p.add_run(txt)
            run.bold = True; run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0x2d, 0x37, 0x48)
        elif btype == 'h4':
            p = doc.add_paragraph()
            run = p.add_run(txt)
            run.bold = True; run.italic = True; run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)
        elif btype == 'para':
            raw = block.get('text','')
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            parts = _re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)', raw)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    r = p.add_run(part[2:-2]); r.bold = True
                elif part.startswith('*') and part.endswith('*'):
                    r = p.add_run(part[1:-1]); r.italic = True
                elif part.startswith('`') and part.endswith('`'):
                    r = p.add_run(part[1:-1])
                    r.font.name = 'Courier New'; r.font.size = Pt(9)
                elif part: p.add_run(part)
            for run in p.runs: run.font.size = Pt(10.5)
        elif btype == 'bullets':
            for item in block['items']:
                p = doc.add_paragraph(style='List Bullet')
                item_clean = _re.sub(r'\*\*(.+?)\*\*', r'\1', item)
                item_clean = _re.sub(r'\*(.+?)\*', r'\1', item_clean)
                run = p.add_run(item_clean); run.font.size = Pt(10.5)
        elif btype == 'numbered':
            for item in block['items']:
                p = doc.add_paragraph(style='List Number')
                item_clean = _re.sub(r'\*\*(.+?)\*\*', r'\1', item)
                run = p.add_run(item_clean); run.font.size = Pt(10.5)
        elif btype == 'table':
            headers = block.get('headers',[]); rows = block.get('rows',[])
            if not headers: continue
            ncols = len(headers)
            tbl = doc.add_table(rows=1+len(rows), cols=ncols)
            tbl.style = 'Table Grid'
            for j, h in enumerate(headers):
                cell = tbl.rows[0].cells[j]
                cell.text = _re.sub(r'\*\*(.+?)\*\*', r'\1', h)
                run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(cell.text)
                run.bold = True; run.font.color.rgb = RGBColor(0xff,0xff,0xff); run.font.size = Pt(10)
                _word_set_cell_bg(cell, '1a365d')
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for ri, row in enumerate(rows):
                bg = 'f7fafc' if ri%2==0 else 'ffffff'
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
        elif btype == 'space': doc.add_paragraph()
    doc.add_paragraph()
    pf = doc.add_paragraph()
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _word_add_border_bottom(pf, 'cbd5e0', 4)
    rf = pf.add_run(f"Gerado por Master IA · {now_str()}")
    rf.font.size = Pt(8); rf.font.color.rgb = RGBColor(0x71,0x80,0x96)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf.read()


# ══════════════════════════════════════════════════════════════════
#  GERADOR EXCEL PROFISSIONAL
# ══════════════════════════════════════════════════════════════════
def gen_excel(titulo, content):
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.cell.cell import MergedCell
    wb = openpyxl.Workbook()
    C_HEADER = "1a365d"; C_ACCENT = "2b6cb0"; C_ACCENT2 = "ebf4ff"
    C_STRIPE = "f7fafc"; C_WHITE  = "FFFFFF"; C_MUTED = "718096"
    C_BORDER = "cbd5e0"; C_GREEN = "c6f6d5"; C_RED = "fed7d7"
    f_hdr   = Font(bold=True, color="FFFFFF", size=10, name='Calibri')
    f_title = Font(bold=True, color=C_HEADER, size=14, name='Calibri')
    f_sub   = Font(bold=True, color=C_ACCENT, size=11, name='Calibri')
    f_body  = Font(size=10, name='Calibri')
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
    def safe_title(t):
        return _re.sub(r'[:\\/?*\[\]]', '-', str(t or 'Dados'))[:28]
    def try_num(v):
        if isinstance(v, (int, float)): return v
        s = str(v).replace(' ','').replace('.','').replace(',','.')
        s = _re.sub(r'[R$%]', '', s).strip()
        try: return float(s)
        except: return v
    def is_num_col(rows, col_idx):
        vals = [try_num(r[col_idx]) for r in rows if col_idx < len(r)]
        nums = [v for v in vals if isinstance(try_num(v), float)]
        return len(nums) >= len(vals) * 0.6 and len(nums) > 0
    blocks = _pdf_parse_md(content)
    tables = [(b['headers'], b['rows']) for b in blocks if b['type'] == 'table']
    texts  = [(b['type'], b.get('text','')) for b in blocks
              if b['type'] in ('h1','h2','h3','para','bullets','numbered')]
    ws = wb.active
    ws.title = safe_title(titulo)
    row = 1
    ws.merge_cells(f"A{row}:H{row}")
    c = ws[f"A{row}"]; c.value = titulo; c.font = f_title
    c.fill = PatternFill("solid", fgColor="f0f4ff")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 28; row += 1
    ws.merge_cells(f"A{row}:H{row}")
    c2 = ws[f"A{row}"]; c2.value = now_str()
    c2.font = f_muted; c2.alignment = Alignment(horizontal="center")
    ws.row_dimensions[row].height = 14; row += 2
    for ttype, ttext in texts:
        if ttype in ('h1','h2','h3'):
            ws.merge_cells(f"A{row}:H{row}")
            c = ws[f"A{row}"]; c.value = ttext
            c.font = f_title if ttype == 'h1' else f_sub if ttype == 'h2' else Font(bold=True, color="2d3748", size=10, name='Calibri')
            ws.row_dimensions[row].height = 20 if ttype == 'h1' else 18; row += 1
    row += 1
    for ti, (headers, rows) in enumerate(tables):
        if ti > 0: row += 2
        ncols = len(headers)
        for j, h in enumerate(headers, 1):
            c = ws.cell(row=row, column=j, value=_re.sub(r'\*\*(.+?)\*\*',r'\1',h))
            c.font = f_hdr; c.fill = fill_hdr; c.border = brd
            c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        ws.row_dimensions[row].height = 20; row += 1
        num_cols = {j for j in range(len(headers)) if is_num_col(rows, j)}
        for ri, r in enumerate(rows):
            bg = fill_alt if ri%2==0 else fill_white
            is_total = any(str(v).upper().strip() in ('TOTAL','SUBTOTAL','SOMA') for v in r[:2])
            for j, val in enumerate(r[:ncols], 1):
                cell_val = val
                if j-1 in num_cols:
                    nv = try_num(val)
                    if isinstance(nv, float): cell_val = nv
                c = ws.cell(row=row, column=j, value=cell_val)
                c.font = Font(bold=True, color="FFFFFF", size=10, name='Calibri') if is_total else f_body
                c.fill = fill_total if is_total else bg
                c.border = brd
                is_n = isinstance(cell_val, (int, float))
                c.alignment = Alignment(horizontal="right" if is_n else "left",
                                        wrap_text=True, vertical="center")
                if is_n and not is_total: c.number_format = '#,##0.00'
            ws.row_dimensions[row].height = 18; row += 1
    for col in ws.columns:
        ml = 10; col_letter = None
        for c in col:
            if isinstance(c, MergedCell): continue
            if col_letter is None:
                try: col_letter = c.column_letter
                except: pass
            try:
                if c.value: ml = max(ml, len(str(c.value)))
            except: pass
        if col_letter:
            ws.column_dimensions[col_letter].width = min(max(ml + 2, 12), 45)
    ws.freeze_panes = "A4"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

# ── Rota gerar doc (PDF/Word/Excel) ──────────────────────────────
@app.route("/gerar", methods=["POST"])
def gerar():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    d = request.get_json() or {}
    tipo    = (d.get("tipo") or "pdf").lower()
    titulo  = d.get("titulo") or "Documento"
    content = d.get("content") or ""
    try:
        if tipo == "pdf":
            data = gen_pdf(titulo, content)
            mime = "application/pdf"
            ext  = "pdf"
        elif tipo == "word":
            data = gen_word(titulo, content)
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ext  = "docx"
        elif tipo == "excel":
            data = gen_excel(titulo, content)
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ext  = "xlsx"
        else:
            return jsonify({"erro": f"Tipo '{tipo}' não suportado"}), 400
        fname = f"{safe_name(titulo)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        registrar_evento("documento", user, tipo)
        return send_file(io.BytesIO(data), as_attachment=True,
                         download_name=fname, mimetype=mime)
    except Exception as e:
        return jsonify({"erro": str(e), "detalhe": tb.format_exc()}), 500


# ══════════════════════════════════════════════════════════════════
#  FERRAMENTAS — GRÁFICOS + PYTHON
# ══════════════════════════════════════════════════════════════════
MASTER_COLORS = ["#3d7eff","#00c2a8","#ff6b6b","#ffd93d","#a78bfa",
                 "#fb923c","#34d399","#60a5fa","#f472b6","#a3e635"]

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
    plt.close(fig); buf.seek(0)
    return base64.b64encode(buf.read()).decode()

def _safe_floats(v): 
    r = []
    for x in v:
        try: r.append(float(x))
        except: r.append(0.0)
    return r

def _safe_labels(l): return [str(x) for x in l]

def tool_grafico_barras(labels, valores, titulo="", xlabel="", ylabel="", horizontal=False):
    labels = _safe_labels(labels); valores = _safe_floats(valores)
    n = min(len(labels), len(valores)); labels, valores = labels[:n], valores[:n]
    if n == 0: raise ValueError("Dados vazios")
    fig, ax = plt.subplots(figsize=(max(7, n*0.8), 4))
    _apply_master_style(fig, [ax])
    cores = [MASTER_COLORS[i % len(MASTER_COLORS)] for i in range(n)]
    y = np.arange(n); max_v = max(valores) if valores else 1
    if horizontal:
        bars = ax.barh(y, valores, color=cores, edgecolor="#2a3045", linewidth=0.5)
        ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=9)
        for bar, val in zip(bars, valores):
            ax.text(bar.get_width()+max_v*0.01, bar.get_y()+bar.get_height()/2,
                    f"{val:,.0f}".replace(",","."), va="center", ha="left",
                    color="#e8eaf0", fontsize=8)
    else:
        bars = ax.bar(y, valores, color=cores, edgecolor="#2a3045", linewidth=0.5, width=0.65)
        ax.set_xticks(y); ax.set_xticklabels(labels, rotation=25, ha="right", fontsize=8)
        for bar, val in zip(bars, valores):
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+max_v*0.01,
                    f"{val:,.0f}".replace(",","."), ha="center", va="bottom",
                    color="#e8eaf0", fontsize=8)
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10, color="#e8eaf0")
    if xlabel: ax.set_xlabel(xlabel, color="#8892a4")
    if ylabel: ax.set_ylabel(ylabel, color="#8892a4")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f"{x:,.0f}".replace(",",".")))
    fig.tight_layout(); return _fig_to_b64(fig)

def tool_grafico_linhas(series, titulo="", xlabel="", ylabel=""):
    if not series: raise ValueError("Series vazio")
    series_clean = []
    for s in series:
        if not isinstance(s, dict): continue
        nome = str(s.get("nome","Série"))
        dados = _safe_floats(s.get("dados") or s.get("values") or s.get("data") or [])
        labels = _safe_labels(s.get("labels") or s.get("x") or [])
        if dados: series_clean.append({"nome":nome,"dados":dados,"labels":labels})
    if not series_clean: raise ValueError("Nenhuma série válida")
    fig, ax = plt.subplots(figsize=(8, 4))
    _apply_master_style(fig, [ax])
    for i, s in enumerate(series_clean):
        cor = MASTER_COLORS[i % len(MASTER_COLORS)]
        xs = range(len(s["dados"]))
        ax.plot(xs, s["dados"], color=cor, linewidth=2.5, marker="o", markersize=5, label=s["nome"])
        ax.fill_between(xs, s["dados"], alpha=0.08, color=cor)
    all_labels = series_clean[0].get("labels",[])
    if all_labels:
        ax.set_xticks(range(len(all_labels)))
        ax.set_xticklabels(all_labels, rotation=20, ha="right", fontsize=8)
    if len(series_clean) > 1:
        ax.legend(fontsize=8, facecolor="#1e2330", edgecolor="#2a3045", labelcolor="#e8eaf0")
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", pad=10, color="#e8eaf0")
    if xlabel: ax.set_xlabel(xlabel, color="#8892a4")
    if ylabel: ax.set_ylabel(ylabel, color="#8892a4")
    fig.tight_layout(); return _fig_to_b64(fig)

def tool_grafico_pizza(labels, valores, titulo=""):
    labels = _safe_labels(labels); valores = _safe_floats(valores)
    n = min(len(labels), len(valores)); labels, valores = labels[:n], valores[:n]
    if n == 0: raise ValueError("Dados vazios")
    labels_short = [l[:18]+'…' if len(l) > 18 else l for l in labels]
    cores = [MASTER_COLORS[i % len(MASTER_COLORS)] for i in range(n)]
    fig, ax = plt.subplots(figsize=(7, 5.5))
    _apply_master_style(fig, [ax])
    wedges, texts, autotexts = ax.pie(valores, labels=None, colors=cores, autopct="%1.1f%%",
        startangle=90, wedgeprops=dict(edgecolor="#181c26", linewidth=1.5), pctdistance=0.75)
    for at in autotexts: at.set_color("#e8eaf0"); at.set_fontsize(9); at.set_fontweight("bold")
    ax.legend(wedges, labels_short, loc="lower center", bbox_to_anchor=(0.5,-0.18),
              ncol=min(3,n), fontsize=8, facecolor="#1e2330", edgecolor="#2a3045",
              labelcolor="#e8eaf0", framealpha=0.9)
    if titulo: ax.set_title(titulo, fontsize=12, fontweight="bold", color="#e8eaf0", pad=12)
    fig.subplots_adjust(bottom=0.25); return _fig_to_b64(fig)

def tool_grafico_dispersao(series, titulo="", xlabel="", ylabel=""):
    if not series: raise ValueError("Series vazio")
    series_clean = []
    for s in series:
        if not isinstance(s, dict): continue
        nome = str(s.get("nome",""))
        x = _safe_floats(s.get("x",[])); y = _safe_floats(s.get("y",[]))
        n = min(len(x), len(y))
        if n > 0: series_clean.append({"nome":nome,"x":x[:n],"y":y[:n]})
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
    fig.tight_layout(); return _fig_to_b64(fig)

# ── Execução Python nível máximo ──────────────────────────────────
EXEC_TIMEOUT = 60

def tool_executar_python(codigo: str):
    ns = {
        "io":io, "os":os, "re":re, "json":json, "datetime":datetime,
        "math":__import__("math"), "base64":base64, "textwrap":textwrap,
        "pd":pd, "np":np, "plt":plt, "matplotlib":matplotlib,
        "openpyxl":openpyxl, "Document":Document, "Pt":Pt,
        "RGBColor":RGBColor, "WD_ALIGN_PARAGRAPH":WD_ALIGN_PARAGRAPH,
        "Font":Font, "PatternFill":PatternFill, "Alignment":Alignment,
        "Border":Border, "Side":Side,
        "requests":_requests, "now_str":now_str, "safe_name":safe_name,
        "__arquivo__": {},
    }
    stdout_buf = io.StringIO()
    stderr_buf = io.StringIO()
    imagem_b64 = None
    erro = False
    _orig_show = plt.show
    def _capture_show(*args, **kwargs):
        nonlocal imagem_b64
        try:
            buf = io.BytesIO()
            plt.savefig(buf, format="png", dpi=130, bbox_inches="tight",
                        facecolor=plt.gcf().get_facecolor())
            plt.close(); buf.seek(0)
            imagem_b64 = base64.b64encode(buf.read()).decode()
        except Exception as e:
            stderr_buf.write(f"[warn] plt.show: {e}\n")
    plt.show = _capture_show
    try:
        result = [None]; exc_holder = [None]
        def _run():
            try:
                with contextlib.redirect_stdout(stdout_buf), \
                     contextlib.redirect_stderr(stderr_buf):
                    exec(compile(codigo, "<master_ia>", "exec"), ns)
            except Exception as e:
                exc_holder[0] = e
        t = threading.Thread(target=_run, daemon=True)
        t.start(); t.join(timeout=EXEC_TIMEOUT)
        if t.is_alive():
            return {"stdout":stdout_buf.getvalue().strip(),
                    "stderr":f"Timeout: execução excedeu {EXEC_TIMEOUT}s",
                    "erro":True,"imagem_b64":None,
                    "arquivo_b64":None,"arquivo_nome":None,"arquivo_tipo":None}
        if exc_holder[0]:
            erro = True; stderr_buf.write(tb.format_exc())
        if plt.get_fignums() and imagem_b64 is None: _capture_show()
        arq = ns.get("__arquivo__") or {}
        arq_b64 = arq.get("b64")
        arq_nome = arq.get("nome")
        arq_tipo = arq.get("tipo")
        if not arq_b64:
            for var_name in ("resultado_arquivo","output_file","arquivo_final"):
                val = ns.get(var_name)
                if isinstance(val, io.BytesIO):
                    val.seek(0); arq_b64 = base64.b64encode(val.read()).decode()
                    arq_nome = var_name + ".bin"; arq_tipo = "application/octet-stream"; break
                elif isinstance(val, dict) and val.get("b64"):
                    arq_b64 = val["b64"]; arq_nome = val.get("nome",var_name)
                    arq_tipo = val.get("tipo","application/octet-stream"); break
        return {"stdout":stdout_buf.getvalue().strip(),
                "stderr":stderr_buf.getvalue().strip(),
                "erro":erro,"imagem_b64":imagem_b64,
                "arquivo_b64":arq_b64,"arquivo_nome":arq_nome,"arquivo_tipo":arq_tipo}
    finally:
        plt.show = _orig_show; plt.close("all")

def tool_excel_avancado(titulo, colunas, linhas, grafico_tipo=None, grafico_series=None):
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.cell.cell import MergedCell
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = _re.sub(r'[:\\/?*\[\]]', '-', titulo or 'Dados')[:28]
    az    = PatternFill("solid", fgColor="1a365d")
    az2   = PatternFill("solid", fgColor="2b6cb0")
    alt   = PatternFill("solid", fgColor="f5f7ff")
    white = PatternFill("solid", fgColor="FFFFFF")
    brd = Border(left=Side(style="thin",color="d0d8f0"),right=Side(style="thin",color="d0d8f0"),
                 top=Side(style="thin",color="d0d8f0"),bottom=Side(style="thin",color="d0d8f0"))
    ncols = len(colunas)
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    c = ws["A1"]; c.value = titulo; c.font = Font(bold=True,color="FFFFFF",size=13)
    c.fill = az; c.alignment = Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
    c2 = ws["A2"]; c2.value = now_str()
    c2.font = Font(color="AAAAAA",size=8,italic=True); c2.alignment = Alignment(horizontal="center")
    for j, col in enumerate(colunas, 1):
        c = ws.cell(row=3, column=j, value=col)
        c.font = Font(bold=True,color="FFFFFF",size=10); c.fill = az2
        c.border = brd; c.alignment = Alignment(horizontal="center",wrap_text=True,vertical="center")
    ws.row_dimensions[3].height = 20
    for ri, row in enumerate(linhas):
        bg = alt if ri%2==0 else white
        for j, val in enumerate(row[:ncols], 1):
            cell_val = val
            try:
                if isinstance(val, str):
                    v = val.replace(".","").replace(",",".")
                    cell_val = float(v) if "." in val or val.lstrip("-").isdigit() else val
            except: pass
            c = ws.cell(row=4+ri, column=j, value=cell_val)
            c.font = Font(size=10); c.fill = bg; c.border = brd
            is_num = isinstance(cell_val,(int,float))
            c.alignment = Alignment(horizontal="right" if is_num else "left",
                                    wrap_text=True,vertical="center")
            if is_num: c.number_format = '#,##0.00'
        ws.row_dimensions[4+ri].height = 18
    for col in ws.columns:
        ml = 10; col_letter = None
        for c in col:
            if isinstance(c, MergedCell): continue
            if col_letter is None:
                try: col_letter = c.column_letter
                except: pass
            try:
                if c.value: ml = max(ml, len(str(c.value)))
            except: pass
        if col_letter: ws.column_dimensions[col_letter].width = min(max(ml+2,12),40)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()


# ══════════════════════════════════════════════════════════════════
#  TOOLS — DEFINIÇÃO PARA A API
# ══════════════════════════════════════════════════════════════════
TOOLS = [
    {
        "name": "buscar_web",
        "description": (
            "Busca informações ATUAIS na internet. Use quando o usuário perguntar sobre: "
            "prazos fiscais atuais, legislação recente, notícias tributárias, taxas SELIC, "
            "índices econômicos, novidades do eSocial/EFD-Reinf, qualquer informação que pode "
            "ter mudado recentemente. SEMPRE use antes de responder sobre prazos e alíquotas."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Termos de busca em português"},
                "num_resultados": {"type": "integer", "description": "Número de resultados (padrão 5)"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "salvar_memoria",
        "description": (
            "Salva uma informação importante na memória permanente do usuário para uso futuro. "
            "Use quando o usuário informar: nome do escritório, clientes recorrentes, preferências, "
            "configurações, qualquer contexto que ele vai querer lembrar nas próximas conversas. "
            "Exemplos de chaves: 'escritorio', 'cidade', 'clientes_principais', 'regime_padrao'."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "chave": {"type": "string", "description": "Nome da informação (ex: escritorio, cidade)"},
                "valor": {"type": "string", "description": "Valor a salvar"}
            },
            "required": ["chave", "valor"]
        }
    },
    {
        "name": "ler_memoria",
        "description": "Lê toda a memória salva do usuário. Use no início de conversas importantes ou quando o usuário pedir para lembrar de algo.",
        "input_schema": {
            "type": "object",
            "properties": {},
            "required": []
        }
    },
    {
        "name": "executar_python",
        "description": (
            "Executa código Python completo. Use para: análise de dados com pandas, "
            "cálculos, cruzamento de planilhas, geração de Excel/Word/PDF via código, "
            "gráficos matplotlib avançados, automação, processamento de texto.\n"
            "Bibliotecas disponíveis sem import: pd, np, plt, io, os, re, json, datetime, "
            "math, base64, openpyxl, Document, Pt, RGBColor, WD_ALIGN_PARAGRAPH, "
            "Font, PatternFill, Alignment, Border, Side, requests, now_str, safe_name.\n"
            "Para PDF: importe reportlab diretamente (from reportlab.lib... import ...).\n"
            "Para Word: use Document() do python-docx.\n"
            "Para Excel: use openpyxl.\n"
            "REGRA ABSOLUTA PARA RETORNAR ARQUIVOS — use EXATAMENTE este padrão:\n"
            "  buf = io.BytesIO()\n"
            "  # salve no buf: doc.save(buf) ou wb.save(buf) ou doc_pdf.build(buf)\n"
            "  buf.seek(0)\n"
            "  __arquivo__['b64'] = base64.b64encode(buf.read()).decode()\n"
            "  __arquivo__['nome'] = 'arquivo.pdf'  # nome com extensao correta\n"
            "  __arquivo__['tipo'] = 'application/pdf'  # MIME type correto\n"
            "NUNCA use print() para retornar arquivo — SEMPRE use __arquivo__.\n"
            "NUNCA chame gen_pdf(), gen_word() ou gen_excel() — essas funcoes NAO estao disponiveis no codigo.\n"
            "Tipos MIME: PDF=application/pdf | Excel=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet | Word=application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "codigo": {"type": "string", "description": "Código Python completo para executar"}
            },
            "required": ["codigo"]
        }
    },
    {
        "name": "gerar_grafico_barras",
        "description": "Gera gráfico de barras. Use para comparativos, rankings, valores por categoria.",
        "input_schema": {
            "type": "object",
            "properties": {
                "labels":     {"type":"array","items":{"type":"string"}},
                "valores":    {"type":"array","items":{"type":"number"}},
                "titulo":     {"type":"string"},
                "xlabel":     {"type":"string"},
                "ylabel":     {"type":"string"},
                "horizontal": {"type":"boolean","description":"True para barras horizontais"}
            },
            "required": ["labels","valores"]
        }
    },
    {
        "name": "gerar_grafico_linhas",
        "description": "Gera gráfico de linhas. Use para séries temporais, evolução, tendências.",
        "input_schema": {
            "type": "object",
            "properties": {
                "series": {
                    "type":"array",
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
        "description": "Gera gráfico de pizza. Use para proporções, participação percentual.",
        "input_schema": {
            "type": "object",
            "properties": {
                "labels":  {"type":"array","items":{"type":"string"}},
                "valores": {"type":"array","items":{"type":"number"}},
                "titulo":  {"type":"string"}
            },
            "required": ["labels","valores"]
        }
    },
    {
        "name": "gerar_grafico_dispersao",
        "description": "Gera scatter plot. Use para correlação entre variáveis.",
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
        "name": "gerar_excel_avancado",
        "description": "Gera planilha Excel formatada com gráfico opcional.",
        "input_schema": {
            "type": "object",
            "properties": {
                "titulo":       {"type":"string"},
                "colunas":      {"type":"array","items":{"type":"string"}},
                "linhas":       {"type":"array","items":{"type":"array"}},
                "grafico_tipo": {"type":"string","enum":["barras","linhas","pizza"]},
                "grafico_series": {"type":"array","items":{"type":"integer"}}
            },
            "required": ["titulo","colunas","linhas"]
        }
    },
    {
        "name": "ler_arquivo",
        "description": (
            "Lê o conteúdo de um arquivo enviado pelo usuário. "
            "Use quando o usuário enviar um arquivo para análise. "
            "Suporta: PDF (texto e OCR), Excel (xlsx/xls), Word (docx), imagem (OCR), CSV/TXT."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "arquivo_b64":  {"type":"string","description":"Conteúdo do arquivo em base64"},
                "arquivo_nome": {"type":"string","description":"Nome do arquivo com extensão"},
                "usar_ocr":     {"type":"boolean","description":"True para forçar OCR em PDF"}
            },
            "required": ["arquivo_b64","arquivo_nome"]
        }
    }
]

def tool_ler_arquivo(arquivo_b64: str, arquivo_nome: str, usar_ocr: bool = False) -> str:
    """Lê qualquer arquivo e retorna o conteúdo como texto."""
    ext = arquivo_nome.rsplit('.', 1)[-1].lower() if '.' in arquivo_nome else ''
    try:
        if ext in ('png','jpg','jpeg','gif','webp','bmp'):
            texto = ocr_imagem(arquivo_b64)
            return f"[OCR da imagem '{arquivo_nome}']\n\n{texto}"
        elif ext == 'pdf':
            # Tenta texto primeiro, OCR se vazio
            texto = ler_pdf_texto(arquivo_b64)
            if not texto.strip() or usar_ocr:
                texto_ocr = ocr_pdf(arquivo_b64)
                if texto_ocr.strip(): texto = texto_ocr
            if not texto.strip(): return f"[PDF '{arquivo_nome}' sem texto extraível]"
            return f"[Conteúdo do PDF '{arquivo_nome}']\n\n{texto}"
        elif ext in ('xlsx','xls'):
            texto = ler_excel(arquivo_b64)
            return f"[Conteúdo da planilha '{arquivo_nome}']\n\n{texto}"
        elif ext == 'docx':
            texto = ler_docx(arquivo_b64)
            return f"[Conteúdo do Word '{arquivo_nome}']\n\n{texto}"
        elif ext in ('txt','csv','py','js','html','json','xml','md'):
            data = base64.b64decode(arquivo_b64)
            for enc in ('utf-8','latin-1','cp1252'):
                try: return f"[Conteúdo de '{arquivo_nome}']\n\n{data.decode(enc)}"
                except: pass
            return f"[Conteúdo de '{arquivo_nome}' — encoding não detectado]"
        else:
            return f"[Arquivo '{arquivo_nome}' recebido — tipo .{ext} não suportado para leitura direta]"
    except Exception as e:
        return f"[Erro ao ler '{arquivo_nome}': {e}]"


# ══════════════════════════════════════════════════════════════════
#  PONTO 1 — BUSCA NA WEB (DuckDuckGo, sem API key)
# ══════════════════════════════════════════════════════════════════
def tool_buscar_web(query: str, num_resultados: int = 5) -> str:
    """Busca informações atuais na web usando DuckDuckGo."""
    try:
        resp = _requests.get(
            "https://api.duckduckgo.com/",
            params={"q": query, "format": "json", "no_html": 1,
                    "skip_disambig": 1, "no_redirect": 1},
            timeout=15, headers={"User-Agent": "Master-IA/3.0"}
        )
        data = resp.json()
        resultados = []
        if data.get("AbstractText"):
            resultados.append(f"Resumo: {data['AbstractText']}")
            if data.get("AbstractSource"):
                resultados.append(f"Fonte: {data['AbstractSource']} — {data.get('AbstractURL','')}")
        for topic in data.get("RelatedTopics", [])[:num_resultados]:
            if isinstance(topic, dict) and topic.get("Text"):
                resultados.append(f"• {topic['Text']}")
        if not resultados:
            return f"Busca realizada para '{query}' — sem resultados diretos. Tente uma query mais específica."
        return f"Resultados para '{query}':\n\n" + "\n".join(resultados)
    except Exception as e:
        return f"[Erro na busca web: {e}]"

# ══════════════════════════════════════════════════════════════════
#  PONTO 2 — MEMÓRIA PERSISTENTE DO USUÁRIO
# ══════════════════════════════════════════════════════════════════
def get_perfil(username: str) -> dict:
    f = DATA_DIR / "perfis" / f"{username}.json"
    f.parent.mkdir(exist_ok=True, parents=True)
    if f.exists():
        try: return json.loads(f.read_text())
        except: pass
    return {}

def save_perfil(username: str, perfil: dict):
    f = DATA_DIR / "perfis" / f"{username}.json"
    f.parent.mkdir(exist_ok=True, parents=True)
    f.write_text(json.dumps(perfil, ensure_ascii=False, indent=2))

def tool_salvar_memoria(username: str, chave: str, valor: str) -> str:
    try:
        perfil = get_perfil(username)
        perfil[chave] = valor
        perfil["_atualizado"] = now_str()
        save_perfil(username, perfil)
        return f"Memória salva: {chave} = {valor}"
    except Exception as e:
        return f"[Erro ao salvar memória: {e}]"

def tool_ler_memoria(username: str) -> str:
    try:
        perfil = get_perfil(username)
        if not perfil:
            return "Nenhuma memória salva ainda."
        linhas = [f"{k}: {v}" for k, v in perfil.items() if not k.startswith("_")]
        return "Memória do usuário:\n\n" + "\n".join(linhas)
    except Exception as e:
        return f"[Erro ao ler memória: {e}]"

@app.route("/perfil", methods=["GET"])
def get_perfil_route():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    return jsonify(get_perfil(user))

@app.route("/perfil", methods=["POST"])
def save_perfil_route():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    d = request.get_json() or {}
    perfil = get_perfil(user)
    perfil.update(d)
    perfil["_atualizado"] = now_str()
    save_perfil(user, perfil)
    return jsonify({"ok": True})

# ══════════════════════════════════════════════════════════════════
#  SYSTEM PROMPT — COMPLETO
# ══════════════════════════════════════════════════════════════════
SYSTEM_PROMPT = """Você é o Master IA — assistente de altíssimo nível para escritório contábil.

═══════════════════════════════════════════════════════
CAPACIDADES — USE ATIVAMENTE
═══════════════════════════════════════════════════════
1. LER ARQUIVOS: quando o usuário enviar PDF, Excel, Word, imagem ou CSV, use a ferramenta ler_arquivo para ler o conteúdo antes de responder.

2. PYTHON AVANÇADO: use executar_python para:
   - Cruzar planilhas, comparar listas, identificar divergências
   - Cálculos tributários, DAS, PGDAS, alíquotas Simples Nacional
   - Gerar relatórios Excel/Word/PDF via código
   - Qualquer análise que exija lógica ou comparação de dados

3. GRÁFICOS: use as ferramentas de gráfico quando pedido ou quando ajudar a visualizar dados.

4. DOCUMENTOS: gere PDF/Word quando pedido — o sistema converte automaticamente.

5. BUSCA WEB: use buscar_web para qualquer informação que pode ter mudado — prazos, legislação nova, taxas, índices. SEMPRE busque antes de responder sobre prazos fiscais atuais.

6. MEMÓRIA: no início de cada conversa, use ler_memoria para carregar o contexto do usuário. Quando o usuário informar algo importante (nome do escritório, clientes, preferências), use salvar_memoria imediatamente.

═══════════════════════════════════════════════════════
FLUXO PARA DOCUMENTOS (declarações, ofícios, cartas)
═══════════════════════════════════════════════════════
REGRA MÁXIMA: NUNCA invente dados. Se faltam informações essenciais, PERGUNTE antes de gerar.

DADOS ESSENCIAIS para qualquer documento:
- Nome(s) e CPF(s) das pessoas — NUNCA invente
- Nome da instituição e CNPJ — NUNCA invente
- Cidade/data — NUNCA invente

Se algum desses falta → PERGUNTE PRIMEIRO com lista clara das informações necessárias.
Se todos estão presentes → GERE DIRETO, limpo, sem placeholder, sem repetição.

Documento correto: 1 página, estrutura simples (cabeçalho → título → corpo → assinatura → carimbo).
PROIBIDO: [placeholders], repetição de parágrafos, "Fim do Documento", aviso legal, burocracia desnecessária.

═══════════════════════════════════════════════════════
QUANDO RESPONDER EM TEXTO
═══════════════════════════════════════════════════════
- Vá DIRETO ao conteúdo. NUNCA comece com "Vou criar", "Vou preparar", "Vou gerar".
- NUNCA diga "não posso gerar PDF/Word/Excel" — o sistema faz isso automaticamente.
- NUNCA peça para o usuário copiar e colar em outro programa.
- Para perguntas fiscais: responda completo, com base legal, exemplos práticos.
- Use markdown: # títulos, **negrito**, tabelas, listas quando útil.

═══════════════════════════════════════════════════════
PERSONALIDADE
═══════════════════════════════════════════════════════
- Direto, confiante, sem frases vazias ("Claro!", "Ótima pergunta!")
- Tom formal para documentos, conversacional para dúvidas
- Se o usuário errar ortografia ou mandar áudio transcrito confuso, entenda o contexto

═══════════════════════════════════════════════════════
CONHECIMENTO FISCAL BRASILEIRO — ESPECIALIDADE MÁXIMA
═══════════════════════════════════════════════════════
Simples Nacional, Lucro Presumido, Lucro Real, MEI, EIRELI, SLU
SPED, EFD-ICMS/IPI, EFD-Contribuições, EFD-Reinf, eSocial, DCTFWeb
DARF, DAS, PGDAS-D, DEFIS, DASN-SIMEI
ICMS, PIS, COFINS, ISS, IRPJ, CSLL, INSS, FGTS, IRRF
NF-e, NFC-e, NFS-e, CT-e, MDF-e, CF-e
DIFAL, CPRB, CEST, NCM, CNAE
MROSC, Lei 13.019/2014, prestação de contas terceiro setor
Rondônia: SEFIN-RO, SEFISC, Domínio Sistemas, Fiscontech

Responda sempre em português brasileiro."""

# ══════════════════════════════════════════════════════════════════
#  ROTA PRINCIPAL — CHAT COM FERRAMENTAS
# ══════════════════════════════════════════════════════════════════
@app.route("/chat_tools", methods=["POST"])
def chat_tools():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401

    d        = request.get_json() or {}
    messages = d.get("messages") or []
    api_key  = d.get("api_key") or ""
    model    = d.get("model") or "claude-opus-4-7"

    if not api_key: return jsonify({"erro":"api_key obrigatória"}), 400
    if not messages: return jsonify({"erro":"messages vazio"}), 400

    # ── Limpa histórico ──────────────────────────────────────────
    msgs_loop = []
    for m in messages:
        role    = m.get("role","")
        content = m.get("content","")
        if role not in ("user","assistant"): continue
        if isinstance(content, str) and content.strip():
            content = content.split("<tool_call>")[0].strip()
            if content: msgs_loop.append({"role":role,"content":content})
        elif isinstance(content, list):
            clean = []
            for b in content:
                if b.get("type") == "text" and b.get("text","").strip():
                    clean.append({"type":"text","text":b["text"].split("<tool_call>")[0].strip()})
                elif b.get("type") in ("image","document"):
                    clean.append(b)
            if clean: msgs_loop.append({"role":role,"content":clean})

    if not msgs_loop: return jsonify({"erro":"messages vazio"}), 400

    def executar_ferramenta(tool_name, tool_input):
        try:
            if tool_name == "gerar_grafico_barras":
                img = tool_grafico_barras(
                    tool_input["labels"], tool_input["valores"],
                    tool_input.get("titulo",""), tool_input.get("xlabel",""),
                    tool_input.get("ylabel",""), tool_input.get("horizontal",False))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Gráfico")}
                result_text = f"Gráfico '{tool_input.get('titulo','')}' gerado."

            elif tool_name == "gerar_grafico_linhas":
                img = tool_grafico_linhas(tool_input["series"],
                    tool_input.get("titulo",""), tool_input.get("xlabel",""), tool_input.get("ylabel",""))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Gráfico")}
                result_text = f"Gráfico de linhas gerado."

            elif tool_name == "gerar_grafico_pizza":
                img = tool_grafico_pizza(tool_input["labels"], tool_input["valores"],
                    tool_input.get("titulo",""))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Pizza")}
                result_text = "Gráfico de pizza gerado."

            elif tool_name == "gerar_grafico_dispersao":
                img = tool_grafico_dispersao(tool_input["series"],
                    tool_input.get("titulo",""), tool_input.get("xlabel",""), tool_input.get("ylabel",""))
                block = {"tipo":"imagem","b64":img,"legenda":tool_input.get("titulo","Dispersão")}
                result_text = "Gráfico de dispersão gerado."

            elif tool_name == "executar_python":
                import concurrent.futures as _cf
                with _cf.ThreadPoolExecutor(max_workers=1) as _ex:
                    _fut = _ex.submit(tool_executar_python, tool_input["codigo"])
                    try: res = _fut.result(timeout=65)
                    except _cf.TimeoutError:
                        res = {"stdout":"","stderr":"Timeout: execução excedeu 60s","erro":True,
                               "imagem_b64":None,"arquivo_b64":None,"arquivo_nome":None,"arquivo_tipo":None}
                block = {
                    "tipo":"codigo_resultado",
                    "codigo":tool_input["codigo"],
                    "stdout":res["stdout"],
                    "stderr":res["stderr"],
                    "erro":res["erro"],
                    "imagem_b64":res.get("imagem_b64"),
                    "arquivo_b64":res.get("arquivo_b64"),
                    "arquivo_nome":res.get("arquivo_nome"),
                    "arquivo_tipo":res.get("arquivo_tipo"),
                }
                out = res["stdout"] or ""
                err = res["stderr"] or ""
                img_info = " [imagem gerada]" if res.get("imagem_b64") else ""
                arq_info = f" [arquivo: {res.get('arquivo_nome')}]" if res.get("arquivo_b64") else ""
                result_text = f"Executado.{img_info}{arq_info}\nSaída: {out[:3000]}"
                if err and res["erro"]: result_text += f"\nErro: {err[:500]}"

            elif tool_name == "gerar_excel_avancado":
                xls_bytes = tool_excel_avancado(
                    tool_input["titulo"], tool_input["colunas"],
                    tool_input["linhas"], tool_input.get("grafico_tipo"),
                    tool_input.get("grafico_series"))
                xls_b64 = base64.b64encode(xls_bytes).decode()
                fname   = safe_name(tool_input["titulo"]) + ".xlsx"
                block   = {"tipo":"excel","b64":xls_b64,"nome":fname,"legenda":tool_input["titulo"]}
                registrar_evento("documento", user, "excel")
                result_text = f"Planilha '{tool_input['titulo']}' gerada com {len(tool_input['linhas'])} linhas."

            elif tool_name == "buscar_web":
                resultado = tool_buscar_web(
                    tool_input["query"],
                    tool_input.get("num_resultados", 5))
                block = {"tipo":"texto","conteudo": resultado}
                result_text = resultado

            elif tool_name == "salvar_memoria":
                resultado = tool_salvar_memoria(
                    user,
                    tool_input["chave"],
                    tool_input["valor"])
                block = {"tipo":"texto","conteudo": resultado}
                result_text = resultado

            elif tool_name == "ler_memoria":
                resultado = tool_ler_memoria(user)
                block = {"tipo":"texto","conteudo": resultado}
                result_text = resultado

            elif tool_name == "ler_arquivo":
                texto = tool_ler_arquivo(
                    tool_input["arquivo_b64"],
                    tool_input["arquivo_nome"],
                    tool_input.get("usar_ocr", False))
                block = {"tipo":"texto","conteudo": f"[Arquivo lido: {tool_input['arquivo_nome']}]"}
                result_text = texto[:8000]  # Limita para não estourar contexto

            else:
                block = {"tipo":"erro_ferramenta","ferramenta":tool_name,"mensagem":"Ferramenta desconhecida"}
                result_text = f"Ferramenta '{tool_name}' não reconhecida."

            return block, result_text

        except Exception as e:
            block = {"tipo":"erro_ferramenta","ferramenta":tool_name,
                     "mensagem":str(e),"trace":tb.format_exc()}
            return block, f"Erro em {tool_name}: {str(e)}"

    # ── Carrega memória e monta system prompt ────────────────────
    perfil = get_perfil(user)
    memoria_txt = ""
    if perfil:
        itens = [f"{k}: {v}" for k, v in perfil.items() if not k.startswith("_")]
        if itens:
            memoria_txt = ("\n\nMEMORIA PERMANENTE DO USUARIO:\n" + 
                          "\n".join(itens))
    system_com_memoria = SYSTEM_PROMPT + memoria_txt

    # ── Loop multi-step ──────────────────────────────────────────
    result_blocks = []
    MAX_ITER = 8
    stop_reason = ""

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
                    "system": system_com_memoria,
                    "tools": TOOLS,
                    "messages": msgs_loop,
                },
                timeout=180
            )
            if not resp.ok:
                return jsonify({"erro": resp.text}), resp.status_code
            data = resp.json()
        except Exception as e:
            return jsonify({"erro": str(e), "detalhe": tb.format_exc()}), 500

        stop_reason = data.get("stop_reason","")
        content     = data.get("content", [])
        tool_uses   = []

        for block in content:
            btype = block.get("type")
            if btype == "text":
                txt = block.get("text","").split("<tool_call>")[0].strip()
                if txt:
                    result_blocks.append({"tipo":"texto","conteudo":txt})
            elif btype == "tool_use":
                tool_uses.append(block)

        if stop_reason != "tool_use" or not tool_uses:
            break

        tool_results = []
        for tu in tool_uses:
            tool_name  = tu.get("name","")
            tool_input = tu.get("input",{})
            tool_id    = tu.get("id", f"tool_{iteration}_{tool_name}")
            rb, result_text = executar_ferramenta(tool_name, tool_input)
            result_blocks.append(rb)
            tool_results.append({
                "type": "tool_result",
                "tool_use_id": tool_id,
                "content": result_text
            })

        msgs_loop.append({"role":"assistant","content":content})
        msgs_loop.append({"role":"user","content":tool_results})

    registrar_evento("mensagem", user)
    return jsonify({"blocos": result_blocks, "stop_reason": stop_reason})


# ══════════════════════════════════════════════════════════════════
#  ROTA COMPARATIVO PGDASD / PLANILHA
# ══════════════════════════════════════════════════════════════════
@app.route("/comparativo", methods=["POST"])
def comparativo():
    user = auth(request)
    if not user: return jsonify({"erro":"Não autenticado"}), 401
    d = request.get_json()
    dados_a = d.get("dados_a") or ""; dados_b = d.get("dados_b") or ""
    label_a = d.get("label_a") or "Sistema A"; label_b = d.get("label_b") or "Sistema B"
    if not dados_a or not dados_b: return jsonify({"erro":"Envie dados_a e dados_b"}), 400
    def parsecsv(txt):
        rows = []
        for line in txt.strip().split("\n"):
            if not line.strip(): continue
            for sep in (";",",","\t"):
                parts = line.split(sep)
                if len(parts) >= 2: rows.append([p.strip() for p in parts]); break
            else: rows.append([line.strip()])
        return rows
    rows_a = parsecsv(dados_a); rows_b = parsecsv(dados_b)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Comparativo"
    az = PatternFill("solid",fgColor="1a365d"); verde = PatternFill("solid",fgColor="e6faf7")
    amarelo = PatternFill("solid",fgColor="fff8e1"); vermelho = PatternFill("solid",fgColor="ffe5e5")
    alt = PatternFill("solid",fgColor="eef2ff"); white = PatternFill("solid",fgColor="FFFFFF")
    f_hdr = Font(bold=True,color="FFFFFF",size=11); f_sub = Font(bold=True,color="2b6cb0",size=11)
    f_ok = Font(color="1a7a5c",size=10); f_diff = Font(bold=True,color="c0392b",size=10)
    brd = Border(left=Side(style="thin",color="d0d8f0"),right=Side(style="thin",color="d0d8f0"),
                 top=Side(style="thin",color="d0d8f0"),bottom=Side(style="thin",color="d0d8f0"))
    ws.merge_cells("A1:H1")
    c = ws["A1"]; c.value = f"Comparativo {label_a} × {label_b}"
    c.font = Font(bold=True,color="1a365d",size=14)
    c.fill = PatternFill("solid",fgColor="f5f7ff"); c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 24
    row = 3
    ws.merge_cells(f"A{row}:H{row}")
    ca = ws[f"A{row}"]; ca.value = f"▌ {label_a}"; ca.font = f_sub
    ca.fill = PatternFill("solid",fgColor="dbeafe"); ca.alignment = Alignment(indent=1); row += 1
    for ri, r in enumerate(rows_a):
        for ci, val in enumerate(r[:8], 1):
            c = ws.cell(row=row,column=ci,value=val)
            c.font = Font(bold=(ri==0),color="FFFFFF" if ri==0 else "2c3347",size=10)
            c.fill = az if ri==0 else (alt if ri%2==0 else white); c.border = brd
            c.alignment = Alignment(wrap_text=True)
        row += 1
    row += 1
    ws.merge_cells(f"A{row}:H{row}")
    cb = ws[f"A{row}"]; cb.value = f"▌ {label_b}"; cb.font = f_sub
    cb.fill = PatternFill("solid",fgColor="dcfce7"); cb.alignment = Alignment(indent=1); row += 1
    for ri, r in enumerate(rows_b):
        for ci, val in enumerate(r[:8], 1):
            c = ws.cell(row=row,column=ci,value=val)
            c.font = Font(bold=(ri==0),color="FFFFFF" if ri==0 else "2c3347",size=10)
            c.fill = PatternFill("solid",fgColor="166534") if ri==0 else (verde if ri%2==0 else white)
            c.border = brd; c.alignment = Alignment(wrap_text=True)
        row += 1
    from openpyxl.cell.cell import MergedCell
    for col in ws.columns:
        ml = 10; col_letter = None
        for c in col:
            if isinstance(c, MergedCell): continue
            if col_letter is None:
                try: col_letter = c.column_letter
                except: pass
            try:
                if c.value: ml = max(ml, len(str(c.value)))
            except: pass
        if col_letter: ws.column_dimensions[col_letter].width = min(max(ml+2,12),50)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); data_bytes = buf.read()
    registrar_evento("documento", user, "comparativo")
    fname = f"Comparativo_{safe_name(label_a)}_{safe_name(label_b)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(io.BytesIO(data_bytes), as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
