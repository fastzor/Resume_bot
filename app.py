"""
БизнесМатика Resume Formatter
Использует оригинальный файл как базу — шрифты Raleway встроены автоматически.
"""
import os, json, tempfile, subprocess, shutil, copy
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
import anthropic

app = Flask(__name__)

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
TEMPLATE     = os.path.join(BASE_DIR, "template.docx")   # оригинальный файл-шаблон

# Шрифты точно как в оригинале
F_MEDIUM  = "Raleway Medium"   # заголовки, метки
F_REGULAR = "Raleway"          # обычный текст
COLOR_BLACK = RGBColor(0x00, 0x00, 0x00)
COLOR_GRAY  = RGBColor(0x59, 0x59, 0x59)   # дата, подзаголовок должности

# ── Извлечение текста ──────────────────────────────────────────────
def extract_text(filepath, filename):
    ext = filename.lower().rsplit(".", 1)[-1]
    if ext in ("docx", "doc"):
        result = subprocess.run(["pandoc", filepath, "-t", "plain"],
                                capture_output=True, text=True)
        if result.stdout.strip():
            return result.stdout
        # Fallback через LibreOffice для .doc
        tmp_dir = os.path.dirname(filepath)
        subprocess.run(["libreoffice", "--headless", "--convert-to", "docx",
                        "--outdir", tmp_dir, filepath], capture_output=True)
        docx_path = filepath.rsplit(".", 1)[0] + ".docx"
        if os.path.exists(docx_path):
            return subprocess.run(["pandoc", docx_path, "-t", "plain"],
                                  capture_output=True, text=True).stdout
        return result.stdout
    elif ext == "pdf":
        result = subprocess.run(["pdftotext", filepath, "-"],
                                capture_output=True, text=True)
        return result.stdout
    with open(filepath, "r", errors="ignore") as f:
        return f.read()

# ── Claude: структурированные данные ──────────────────────────────
PROMPT = """Ты — парсер резюме. Верни ТОЛЬКО валидный JSON без markdown.

{
  "name": "Имя Ф.",
  "position": "Должность",
  "grade": "Junior | Middle | Senior",
  "languages": [{"lang": "Русский", "level": "Родной"}],
  "skills": ["Python", "SQL"],
  "about": "3-5 предложений",
  "experience": [{"role":"","period":"","project_name":"","project_desc":"","tasks":[],"achievements":[],"team":"","stack":""}],
  "education": [{"year":"","institution":"","specialty":""}],
  "courses": [{"year":"","type":"курс | сертификат | тест | экзамен","platform":"","name":""}]
}

Имя: только Имя + первая буква фамилии с точкой (пример: Алексей О.).
Грейд определи по опыту: до 2 лет Junior, 2-5 Middle, 5+ Senior.
Поле courses — собирай ВСЁ что находится под образованием:
  - курсы (любые платформы: Stepik, Coursera, Яндекс Практикум, Udemy и др.)
  - сертификаты (профессиональные, вендорские: AWS, Google, Microsoft и др.)
  - тесты и экзамены (любые профессиональные тесты)
  - дополнительное обучение (любые программы)
  Поле type заполни одним из: курс, сертификат, тест, экзамен
Только JSON, никакого другого текста."""

def parse_with_claude(text):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    resp = client.messages.create(
        model="claude-sonnet-4-20250514", max_tokens=4000,
        messages=[{"role":"user","content":f"{PROMPT}\n\nРЕЗЮМЕ:\n{text}"}]
    )
    raw = resp.content[0].text.strip().replace("```json","").replace("```","").strip()
    return json.loads(raw)

# ── Хелперы форматирования ────────────────────────────────────────
def _font(run, name, size_pt, color):
    run.font.name  = name
    run.font.size  = Pt(size_pt)
    run.font.bold  = False
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rf  = OxmlElement('w:rFonts')
    for attr in ('w:ascii','w:hAnsi','w:cs','w:eastAsia'):
        rf.set(qn(attr), name)
    rPr.insert(0, rf)

def med(para, text, size=10):
    """Raleway Medium — метки, заголовки"""
    r = para.add_run(text)
    _font(r, F_MEDIUM, size, COLOR_BLACK)
    return r

def reg(para, text, size=9, color=COLOR_BLACK):
    """Raleway Regular — обычный текст"""
    r = para.add_run(text)
    _font(r, F_REGULAR, size, color)
    return r

def sp(p, before=0, after=0):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)


def bullet_para(doc_or_cell, text, size=9):
    """Параграф с буллетом через XML (не зависит от стилей)"""
    p = doc_or_cell.add_paragraph()
    sp(p, before=1, after=1)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), '360')
    ind.set(qn('w:hanging'), '180')
    pPr.append(ind)
    run = p.add_run("• ")
    _font(run, F_REGULAR, size, COLOR_BLACK)
    run2 = p.add_run(text)
    _font(run2, F_REGULAR, size, COLOR_BLACK)
    return p

def no_borders(table):
    tbl   = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
    bdr = OxmlElement('w:tblBorders')
    for s in ('top','left','bottom','right','insideH','insideV'):
        b = OxmlElement(f'w:{s}'); b.set(qn('w:val'),'none'); bdr.append(b)
    tblPr.append(bdr)

def thin_border(cell, clr="C8C8C8"):
    tcPr = cell._tc.get_or_add_tcPr()
    bdr  = OxmlElement('w:tcBorders')
    for s in ('top','left','bottom','right'):
        b = OxmlElement(f'w:{s}')
        b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'4')
        b.set(qn('w:space'),'0');    b.set(qn('w:color'),clr)
        bdr.append(b)
    tcPr.append(bdr)

def cell_w(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    w = OxmlElement('w:tcW')
    w.set(qn('w:w'), str(int(cm*567))); w.set(qn('w:type'),'dxa')
    tcPr.append(w)

def cell_mar(cell, t=60, b=60, l=100, r=100):
    tcPr = cell._tc.get_or_add_tcPr()
    mar  = OxmlElement('w:tcMar')
    for s, v in (('top',t),('bottom',b),('left',l),('right',r)):
        m = OxmlElement(f'w:{s}')
        m.set(qn('w:w'),str(v)); m.set(qn('w:type'),'dxa')
        mar.append(m)
    tcPr.append(mar)

# ── Основная генерация ────────────────────────────────────────────
def build_resume(data, output_path):
    # Берём оригинал как базу — шрифты, логотипы, настройки сохраняются
    shutil.copy(TEMPLATE, output_path)
    doc = Document(output_path)

    # Очищаем тело документа (оставляем только структуру)
    body = doc.element.body
    for child in list(body):
        tag = child.tag.split('}')[-1]
        if tag not in ('sectPr',):   # sectPr = поля страницы — оставляем
            body.remove(child)

    ns = doc.styles['Normal']
    ns.font.name = F_REGULAR
    ns.font.size = Pt(9)
    ns.paragraph_format.space_before = Pt(0)
    ns.paragraph_format.space_after  = Pt(0)

    def add_p(alignment=None, before=0, after=0):
        p = doc.add_paragraph()
        sp(p, before, after)
        if alignment:
            p.alignment = alignment
        return p

    # ── ФИО — по центру, Raleway Medium 20pt ─────────────────────
    p = add_p(WD_ALIGN_PARAGRAPH.CENTER, before=8, after=2)
    med(p, data.get("name",""), size=20)

    # ── Должность — по центру, Raleway Regular 10pt серый ────────
    p = add_p(WD_ALIGN_PARAGRAPH.CENTER, before=0, after=8)
    reg(p, data.get("position",""), size=10, color=COLOR_GRAY)

    # ── Грейд ─────────────────────────────────────────────────────
    p = add_p(before=4, after=6)
    med(p, "Грейд: ", size=10)
    reg(p, data.get("grade",""), size=9)

    # ── Языки ─────────────────────────────────────────────────────
    langs = data.get("languages",[])
    if langs:
        p = add_p(before=8, after=4)
        med(p, "Языки:", size=10)
        for lang in langs:
            bullet_para(doc, f"{lang.get('lang','')} - {lang.get('level','')}")

    # ── Навыки — таблица 3 колонки, тонкие рамки ─────────────────
    skills = data.get("skills",[])
    if skills:
        p = add_p(before=10, after=4)
        med(p, "Навыки:", size=10)

        cols = 3
        rows = (len(skills)+cols-1)//cols
        padded = skills + [""]*( rows*cols - len(skills))
        tbl = doc.add_table(rows=rows, cols=cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        no_borders(tbl)
        for ri in range(rows):
            for ci in range(cols):
                c = tbl.cell(ri, ci)
                cell_w(c, 4.7)
                thin_border(c, "C8C8C8")
                cell_mar(c, t=70, b=70, l=110, r=110)
                cp = c.paragraphs[0]
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sp(cp)
                reg(cp, padded[ri*cols+ci], size=9)
        add_p(before=0, after=4)

    # ── О себе — метка Medium inline ─────────────────────────────
    about = data.get("about","")
    if about:
        p = add_p(before=8, after=6)
        med(p, "О себе: ", size=10)
        reg(p, about, size=9)

    # ── Опыт работы ───────────────────────────────────────────────
    exp = data.get("experience",[])
    if exp:
        p = add_p(before=10, after=6)
        med(p, "Опыт работы:", size=10)

        for job in exp:
            tbl = doc.add_table(rows=1, cols=2)
            tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
            no_borders(tbl)
            lc = tbl.cell(0,0); rc = tbl.cell(0,1)
            cell_w(lc, 4.4); cell_w(rc, 10.5)
            lc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            rc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell_mar(lc, t=40, b=40, l=0, r=140)
            cell_mar(rc, t=40, b=40, l=0, r=0)

            # Левая: должность + дата
            p1 = lc.paragraphs[0]; sp(p1, before=0, after=3)
            med(p1, job.get("role",""), size=9)
            p2 = lc.add_paragraph(); sp(p2, before=0, after=6)
            reg(p2, job.get("period",""), size=9, color=COLOR_GRAY)

            # Правая: детали
            rp = rc.paragraphs[0]; sp(rp, before=0, after=2)
            med(rp, "Название проекта: ", size=9)
            reg(rp, job.get("project_name",""), size=9)

            if job.get("project_desc"):
                p2 = rc.add_paragraph(); sp(p2, before=0, after=4)
                med(p2, "Описание проекта: ", size=9)
                reg(p2, job.get("project_desc",""), size=9)

            if job.get("tasks"):
                ph = rc.add_paragraph(); sp(ph, before=6, after=2)
                med(ph, "Задачи:", size=9)
                for t in job["tasks"]:
                    bullet_para(rc, t)

            if job.get("achievements"):
                ph = rc.add_paragraph(); sp(ph, before=6, after=2)
                med(ph, "Достижения:", size=9)
                for a in job["achievements"]:
                    bullet_para(rc, a)

            if job.get("team"):
                pt = rc.add_paragraph(); sp(pt, before=6, after=2)
                med(pt, "Команда: ", size=9)
                reg(pt, job["team"], size=9)

            if job.get("stack"):
                ps = rc.add_paragraph(); sp(ps, before=2, after=6)
                med(ps, "Стек: ", size=9)
                reg(ps, job["stack"], size=9)

            add_p(before=0, after=10)

    # ── Образование ───────────────────────────────────────────────
    edu = data.get("education",[])
    if edu:
        p = add_p(before=10, after=4)
        med(p, "Образование:", size=10)
        for e in edu:
            parts = [x for x in [e.get("year"),e.get("institution"),e.get("specialty")] if x]
            bp = doc.add_paragraph(style=None)
            sp(bp, before=1, after=1)
            reg(bp, " | ".join(parts), size=9)

    # ── Курсы ─────────────────────────────────────────────────────
    courses = data.get("courses",[])
    if courses:
        p = add_p(before=10, after=4)
        med(p, "Курсы, дополнительное обучение:", size=10)
        for c in courses:
            # Собираем строку: год | платформа | название [тип если сертификат]
            ctype = c.get("type","").lower()
            label = ""
            if ctype in ("сертификат", "certificate"):
                label = " [Сертификат]"
            elif ctype in ("тест", "test"):
                label = " [Тест]"
            elif ctype in ("экзамен", "exam"):
                label = " [Экзамен]"
            parts = [x for x in [c.get("year"), c.get("platform"), c.get("name")] if x]
            bullet_para(doc, " | ".join(parts) + label)

    doc.save(output_path)

# ── Flask ──────────────────────────────────────────────────────────
@app.route("/format-resume", methods=["POST"])
def format_resume():
    if "file" not in request.files:
        return jsonify({"error":"No file"}), 400
    file = request.files["file"]
    fname = file.filename or "resume.docx"
    with tempfile.TemporaryDirectory() as tmp:
        inp = os.path.join(tmp, fname)
        out = os.path.join(tmp, "result.docx")
        file.save(inp)
        text = extract_text(inp, fname)
        if not text.strip():
            return jsonify({"error":"Cannot extract text"}), 400
        data = parse_with_claude(text)
        build_resume(data, out)
        slug = data.get("name","resume").replace(" ","_").replace(".","")
        return send_file(out, as_attachment=True,
                         download_name=f"{slug}_БизнесМатика.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.route("/health")
def health():
    return jsonify({"status":"ok"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)
