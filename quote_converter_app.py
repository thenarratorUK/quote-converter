# quote_converter_app_pdf2docx_final_globclean_v3.py
import io, os, re, tempfile, streamlit as st

try:
    from docx import Document
except Exception:
    Document = None

try:
    from pdf2docx import Converter as PDF2DOCXConverter
except Exception:
    PDF2DOCXConverter = None

_ASCII_CTRL = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')

def _drop_nonchars(s: str) -> str:
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFDD0 <= code <= 0xFDEF or (code & 0xFFFE) == 0xFFFE:
            continue
        if 0xD800 <= code <= 0xDFFF:
            out.append('\uFFFD'); continue
        out.append(ch)
    return ''.join(out)

def _xml10_filter(text: str) -> str:
    if not text:
        return text
    out = []
    for ch in text:
        code = ord(ch)
        if code in (0x9, 0xA, 0xD) or 0x20 <= code <= 0xD7FF or 0xE000 <= code <= 0xFFFD or 0x10000 <= code <= 0x10FFFF:
            out.append(ch)
    return ''.join(out)

def sanitize_for_docx(text: str) -> str:
    if not text:
        return text
    text = _ASCII_CTRL.sub('', text)
    text = _drop_nonchars(text)
    return _xml10_filter(text)

def _detect_primary_style(text: str) -> str:
    if not text:
        return "UNKNOWN"
    singles_open = len(re.findall(r'(^|[\s(\[{<])‘', text))
    doubles_open = len(re.findall(r'(^|[\s(\[{<])“', text))
    singles_total = text.count("‘") + text.count("’")
    doubles_total = text.count("“") + text.count("”")
    if singles_open >= doubles_open * 1.5 and singles_open >= 4:
        return "UK"
    if doubles_open >= singles_open * 1.2 and doubles_open >= 4:
        return "US"
    if doubles_total > singles_total * 1.2 and doubles_open >= 2:
        return "US"
    if singles_total > doubles_total * 1.5 and singles_open >= 2:
        return "UK"
    return "UNKNOWN"

def normalize_quotes_to_us(text: str) -> str:
    if not text:
        return text
    APOS = "<<APOS>>"
    text = re.sub(r"(?<=\w)[’'](?=\w)", APOS, text)
    style = _detect_primary_style(text)
    if style == "UK":
        OPEN_S, CLOSE_S, OPEN_D, CLOSE_D = "<<OPEN_S>>", "<<CLOSE_S>>", "<<OPEN_D>>", "<<CLOSE_D>>"
        t = (text.replace("‘", OPEN_S)
                 .replace("’", CLOSE_S)
                 .replace("“", OPEN_D)
                 .replace("”", CLOSE_D))
        t = re.sub(r'(?<=\w)'+re.escape(CLOSE_S)+r'(?=\w)', APOS, t)
        for w in ("em","cause","til","tis","twas","sup","round","clock"):
            t = re.sub(r'\b'+re.escape(CLOSE_S)+w+r'\b', APOS+w, t, flags=re.IGNORECASE)
        t = re.sub(re.escape(CLOSE_S)+r'(?=\d{2}s\b)', APOS, t)
        t = (t.replace(OPEN_S,"“").replace(CLOSE_S,"”").replace(OPEN_D,"‘").replace(CLOSE_D,"’"))
        text = t
    else:
        def smarten_line(line: str) -> str:
            out, open_d = [], True
            for ch in line:
                if ch == '"':
                    out.append("“" if open_d else "”"); open_d = not open_d
                elif ch == "'":
                    out.append("’")
                else:
                    out.append(ch)
            return "".join(out)
        text = "\n".join(smarten_line(ln) for ln in text.split("\n"))
    return text.replace(APOS, "’")

def convert_docx_runs_to_us(doc: Document) -> None:
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = normalize_quotes_to_us(sanitize_for_docx(r.text))
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = normalize_quotes_to_us(sanitize_for_docx(r.text))

def _remove_global_shapes_all_parts(doc: Document) -> None:
    """
    Delete drawing/pict/object/sym/txbx/wsp elements from the main document and all related parts
    (headers/footers), then remove empty runs and paragraphs.
    """
    pkg = doc.part.package
    for part in pkg.parts:
        elt = getattr(part, 'element', None)
        if elt is None:
            continue
        # 1) Remove drawings/picts/objects/symbols and deep textbox containers
        nodes = list(elt.xpath(
            './/*[local-name()="drawing" or local-name()="pict" or local-name()="object" or local-name()="sym" or local-name()="wsp" or local-name()="txbx" or local-name()="txbxContent"]'
        ))
        for n in nodes:
            parent = n.getparent()
            if parent is not None:
                parent.remove(n)
        # 2) Remove empty runs
        for r in list(elt.xpath('.//*[local-name()="r"]')):
            has_text = bool(r.xpath('.//*[local-name()="t" and normalize-space(text())]'))
            has_children = len(r) > 0
            if not has_text and not has_children:
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)
        # 3) Remove paragraphs that are now empty or whitespace-only
        for p in list(elt.xpath('.//*[local-name()="p"]')):
            has_text = bool(p.xpath('.//*[local-name()="t" and normalize-space(text())]'))
            has_draw = bool(p.xpath('.//*[local-name()="drawing" or local-name()="pict" or local-name()="object" or local-name()="sym" or local-name()="wsp" or local-name()="txbx" or local-name()="txbxContent"]'))
            if not has_text and not has_draw:
                parent = p.getparent()
                if parent is not None:
                    parent.remove(p)

def convert_docx_bytes_to_us(docx_bytes: bytes) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx required.")
    doc = Document(io.BytesIO(docx_bytes))
    fix_dropcaps_acbd(doc)
    convert_docx_runs_to_us(doc)
    out = io.BytesIO(); doc.save(out)
    return out.getvalue()

def pdf_bytes_to_docx_using_pdf2docx(pdf_bytes: bytes) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx required.")
    if PDF2DOCXConverter is None:
        raise RuntimeError("pdf2docx required.")
    with tempfile.TemporaryDirectory() as tmpd:
        pdf_path = os.path.join(tmpd, "in.pdf")
        out_path = os.path.join(tmpd, "out.docx")
        with open(pdf_path, "wb") as f: f.write(pdf_bytes)
        cv = PDF2DOCXConverter(pdf_path)
        cv.convert(out_path, start=0, end=None)
        cv.close()
        doc = Document(out_path)
        fix_dropcaps_acbd(doc)

        # 1) Deep removal across all parts (fix persistent squares)
        _remove_global_shapes_all_parts(doc)

        # 2) Run-level cleanup and cautious mid-sentence blank removal
        paras = doc.paragraphs
        for i, p in enumerate(paras):
            for r in p.runs:
                if r.text:
                    r.text = (r.text.replace("\uFFFC","")
                                   .replace("\u00A0"," ")
                                   .replace("\u000c",""))
            if p.text.strip() in {"", "\u00A0"} and 0 < i < len(paras)-1:
                prev = paras[i-1].text.strip()
                nxt  = paras[i+1].text.strip()
                if prev and nxt and not re.search(r'[.!?]"?$', prev):
                    p.text = ""

        # 3) Normalize quotes to US
        convert_docx_runs_to_us(doc)

        buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

st.set_page_config(page_title="Quote Style Converter (Global Clean v3)", page_icon="📝", layout="centered")

CSS = """
:root { --primary-color:#008080;--primary-hover:#006666;--bg-1:#0b0f14;--bg-2:#11161d;
--card:#0f141a;--text-1:#e8eef5;--text-2:#b2c0cf;--muted:#8aa0b5;--accent:#e0f2f1;--ring:rgba(0,128,128,0.5);}
html,body,[data-testid="stAppViewContainer"]{
  background:linear-gradient(180deg,var(--bg-1),var(--bg-2))!important;
  color:var(--text-1)!important;
}
a{color:var(--accent)!important;}
div.stButton>button{
  background-color:var(--primary-color);
  color:#e8eef5;
  border:none;
  border-radius:.6rem;
  padding:.6rem 1rem;
}
div.stButton>button:hover{background-color:var(--primary-hover);}
body{font-family:Avenir,sans-serif;line-height:1.65;}
"""
st.markdown("<style>\n"+CSS+"\n</style>", unsafe_allow_html=True)

st.title("Quote Style Converter (pdf2docx – Global Clean v3)")
st.caption("Layout-preserving PDF→DOCX with US quotes and deepest cleanup of page-join squares.")

with st.container():
    mode = st.radio("Choose input type", ["DOCX → DOCX (UK → US)", "PDF → DOCX (pdf2docx → US quotes)"])
    uploaded = st.file_uploader("Upload file", type=["docx","pdf"])

if uploaded is not None:
    if mode.startswith("DOCX"):
        if not uploaded.name.lower().endswith(".docx"):
            st.error("Please upload a .docx file for this mode.")
        elif st.button("Convert DOCX to US quotes"):
            try:
                out_bytes = convert_docx_bytes_to_us(uploaded.read())
                st.success("Converted. Download below.")
                st.download_button("Download DOCX (US quotes)", out_bytes,
                    file_name=uploaded.name.rsplit(".",1)[0]+" (US Quotes).docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"Conversion failed: {e}")
    else:
        if not uploaded.name.lower().endswith(".pdf"):
            st.error("Please upload a .pdf file for this mode.")
        elif st.button("Convert PDF → DOCX (pdf2docx → US quotes)"):
            try:
                out_bytes = pdf_bytes_to_docx_using_pdf2docx(uploaded.read())
                st.success("Converted. Download below.")
                st.download_button("Download DOCX (US quotes)", out_bytes,
                    file_name=uploaded.name.rsplit(".",1)[0]+" (US Quotes).docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"Conversion failed: {e}")


# === ACBD drop-cap fixer (A=big letter+space, B=normal runs, C=from ALL-CAPS word to next paragraph that has <w:widowControl/>, D=that next paragraph) ===
def _acbd_pt(val, default=None):
    try:
        return float(val.pt) if hasattr(val, "pt") else (float(val) if val is not None else default)
    except Exception:
        return default

def _acbd_run_size_pt(run, para, default=11.0):
    sz = _acbd_pt(getattr(run.font, "size", None))
    if sz is not None:
        return sz
    try:
        psz = _acbd_pt(para.style.font.size, None)
        if psz is not None:
            return psz
    except Exception:
        pass
    return default

def _acbd_run_text(run):
    try:
        return "".join(t.text or "" for t in run._element.xpath(".//w:t", namespaces=run._element.nsmap))
    except Exception:
        return getattr(run, "text", "") or ""

def _acbd_first_word(s):
    m = re.search(r"[A-Za-z]+", s or "")
    return m.group(0) if m else ""

def _acbd_is_all_caps_word(w):
    return len(w) >= 2 and w.isalpha() and w.upper() == w

def _acbd_para_has_widowcontrol(para):
    try:
        el = para._element
        return bool(el.xpath(".//w:pPr/w:widowControl", namespaces=el.nsmap))
    except Exception:
        return False

def _acbd_fix_once_in_paragraph(doc, p_index):
    paras = doc.paragraphs
    if p_index < 0 or p_index >= len(paras):
        return False
    p = paras[p_index]
    runs = list(p.runs)
    if not runs:
        return False

    sizes = [_acbd_run_size_pt(r, p) for r in runs]
    vals = [v for v in sizes if v is not None]
    if not vals:
        return False
    try:
        import statistics as _stats
        majority = float(_stats.median(vals))
    except Exception:
        majority = sum(vals)/len(vals)
    threshold = 1.5 * majority

    # A
    A_idx = None
    A_char = None
    for i, r in enumerate(runs):
        txt = _acbd_run_text(r)
        alpha = [ch for ch in txt if ch.isalpha()]
        if len(alpha) == 1 and sizes[i] is not None and sizes[i] >= threshold:
            A_idx = i
            A_char = alpha[0]
            break
    if A_idx is None:
        return False

    # C-start: first run after A whose FIRST WORD is ALL-CAPS
    C_start = None
    for j in range(A_idx+1, len(runs)):
        fw = _acbd_first_word(_acbd_run_text(runs[j]))
        if _acbd_is_all_caps_word(fw):
            C_start = j
            break
    if C_start is None:
        return False

    # D begins at next paragraph which must have widowControl (and exists)
    if p_index + 1 >= len(paras):
        return False
    if not _acbd_para_has_widowcontrol(paras[p_index+1]):
        return False

    # B text: runs between A+1 and C_start
    B_text = "".join(_acbd_run_text(runs[t]) for t in range(A_idx+1, C_start)).strip()
    # C text: runs from C_start to end of paragraph p_index
    C_text = "".join(_acbd_run_text(runs[t]) for t in range(C_start, len(runs))).strip()
    if not B_text or not C_text:
        return False

    new_text = (A_char.upper() + C_text).strip()
    if B_text:
        new_text += " " + B_text

    if new_text and new_text != (p.text or "").strip():
        p.text = new_text
        return True
    return False

def fix_dropcaps_acbd(doc, max_passes=50):
    passes = 0
    while passes < max_passes:
        changes = 0
        for i in range(len(doc.paragraphs)):
            if _acbd_fix_once_in_paragraph(doc, i):
                changes += 1
        if changes == 0:
            break
        passes += 1
    return doc
# === end ACBD fixer ===

