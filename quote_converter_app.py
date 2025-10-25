import io, os, re, tempfile, streamlit as st


# === ACBD drop-cap fixer (internal logic; no UI changes) ===
# Pattern:
#   A = single big glyph run (≥1.5× paragraph median size), usually letter + space
#   B = subsequent normal-sized runs (same paragraph) until C-start
#   C = runs starting at the first ALL-CAPS word (same paragraph), and logically ending before the next paragraph that has <w:widowControl/>
#   D = the next paragraph (which has <w:widowControl/>)
# Reorder paragraph text to: A + C + " " + B (leave D untouched). Multi-pass until stable.
import statistics as _acbd_stats

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
    import re as _re
    m = _re.search(r"[A-Za-z]+", s or "")
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
        majority = float(_acbd_stats.median(vals))
    except Exception:
        majority = sum(vals)/len(vals)
    threshold = 1.5 * majority

    # A: first run that's a single alphabetic glyph (optionally with space) and large
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

    # Next paragraph must exist and have widowControl (terminates C; D begins there)
    if p_index + 1 >= len(paras):
        return False
    if not _acbd_para_has_widowcontrol(paras[p_index+1]):
        return False

    # B: text between A+1 and C_start
    B_text = "".join(_acbd_run_text(runs[t]) for t in range(A_idx+1, C_start)).strip()
    # C: text from C_start to end of this paragraph
    C_text = "".join(_acbd_run_text(runs[t]) for t in range(C_start, len(runs))).strip()
    if not B_text or not C_text:
        return False

    # New paragraph text: A + C + " " + B  (no space between A and start of C)
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

CSS = """:root {
  --primary-color: #008080;      /* Teal */
  --primary-hover: #007070;
  --background-color: #fdfdfd;
  --text-color: #222222;
  --card-background: #ffffff;
  --card-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
  --border-radius: 10px;
  --font-family: 'Avenir', sans-serif;
  --accent-color: #ff9900;
}

/* Global Styles */
body {
  background-color: var(--background-color);
  font-family: var(--font-family);
  color: var(--text-color);
  margin: 0;
  padding: 0;
}

h1, h2, h3, h4, h5, h6 {
  color: var(--text-color);
  font-weight: 700;
  margin-bottom: 0.5em;
}

/* Button Styles */
div.stButton > button {
  background-color: var(--primary-color);
  color: #ffffff;
  border: none;
  padding: 0.75em 1.25em;
  border-radius: var(--border-radius);
  cursor: pointer;
  transition: background-color 0.3s ease, transform 0.2s;
}

div.stButton > button:hover {
  background-color: var(--primary-hover);
  transform: translateY(-2px);
}

/* Card/Container Styling */
.custom-container {
  background: var(--card-background);
  padding: 2em;
  border-radius: var(--border-radius);
  box-shadow: var(--card-shadow);
  margin-bottom: 2em;
}

.css-1d391kg {
  background: var(--card-background);
  padding: 1em;
  border-radius: var(--border-radius);
  box-shadow: var(--card-shadow);
}

/* Form Element Styling */
input, select, textarea {
  border: 1px solid #ddd;
  border-radius: 4px;
  padding: 0.5em;
  font-size: 1em;
}

input:focus, select:focus, textarea:focus {
  outline: none;
  border-color: var(--primary-color);
  box-shadow: 0 0 5px rgba(0, 128, 128, 0.3);
}

/* Enforce font in uploader */
.stFileUploader, .stFileUploader label, .stFileUploader div, .stFileUploader button, .stFileUploader *,
[data-testid="stFileUploader"], [data-testid="stFileUploader"] *, [data-testid="stFileUploadDropzone"], [data-testid="stFileUploadDropzone"] *,
input[type="file"] {
  font-family: var(--font-family) !important;
}

"""
st.markdown("<style>\n"+CSS+"\n</style>", unsafe_allow_html=True)

st.title("UK to US Quote Converter with Optional PDF to DOCX Conversion")
st.write("Please upload a docx using single-quotes dialogue for conversion to double-quotes dialogue, or upload a PDF of either type for conversion to double-quotes dialogue in a docx.")

uploaded = st.file_uploader(
    "Upload DOCX (single-quotes) or PDF",
    type=["docx", "pdf"],
    accept_multiple_files=False,
    key="file",
    label_visibility="collapsed"
)



if uploaded is not None:
    # Show a simple file summary and a Convert button
    st.write(f"Selected file: **{uploaded.name}**")
    if st.button("Convert"):
        name_lower = uploaded.name.lower()
        if name_lower.endswith(".docx"):
            if Document is None:
                st.error("python-docx not available; cannot process DOCX.")
            else:
                try:
                    raw = uploaded.read()
                    out_bytes = docx_bytes_to_us_quotes(raw) if 'docx_bytes_to_us_quotes' in globals() else convert_docx_bytes_to_us(raw)
                    st.success("Converted. Download below.")
                    st.download_button("Download File", out_bytes,
                        file_name=uploaded.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"Conversion failed: {e}")
        elif name_lower.endswith(".pdf"):
            if PDF2DOCXConverter is None:
                st.error("pdf2docx not available; cannot convert PDF to DOCX.")
            else:
                try:
                    out_bytes = pdf_bytes_to_docx_using_pdf2docx(uploaded.read())
                    st.success("Converted. Download below.")
                    base = uploaded.name.rsplit(".",1)[0]
                    st.download_button("Download File", out_bytes,
                        file_name=base + ".docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"Conversion failed: {e}")
        else:
            st.error("Unsupported file type. Please upload a .docx or .pdf.")