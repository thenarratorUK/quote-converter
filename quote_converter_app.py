import io, os, re, tempfile, streamlit as st

try:
    from docx import Document
except Exception:
    Document = None

try:
    from pdf2docx import Converter as PDF2DOCXConverter
except Exception:
    PDF2DOCXConverter = None

# === Drop-cap heuristic (strict, size-based) ===
def _dc_run_size_pt(run, para, default=11.0):
    def _pt(x):
        try:
            return float(x.pt) if hasattr(x, "pt") else (float(x) if x is not None else None)
        except Exception:
            return None
    sz = _pt(getattr(run.font, "size", None))
    if sz is not None:
        return sz
    try:
        psz = _pt(para.style.font.size)
        if psz is not None:
            return psz
    except Exception:
        pass
    return default

def _dc_alpha_positions_with_sizes(paragraph):
    for ri, run in enumerate(paragraph.runs):
        t = run.text or ""
        if not t:
            continue
        size = _dc_run_size_pt(run, paragraph)
        for ci, ch in enumerate(t):
            if ch.isalpha():
                yield ri, ci, ch, size

def _dc_first_alpha(paragraph):
    for item in _dc_alpha_positions_with_sizes(paragraph):
        return item
    return None

def _dc_next_alpha_in_flow(paragraphs, start_para_idx, start_pos):
    # Search within the same paragraph after the given (ri, ci), then across following paragraphs
    p = paragraphs[start_para_idx]
    ri0, ci0 = start_pos
    for ri, ci, ch, sz in _dc_alpha_positions_with_sizes(p):
        if (ri > ri0) or (ri == ri0 and ci > ci0):
            return ch, sz, start_para_idx
    for j in range(start_para_idx + 1, len(paragraphs)):
        for ri, ci, ch, sz in _dc_alpha_positions_with_sizes(paragraphs[j]):
            return ch, sz, j
    return None, None, None

def fix_dropcap_reordering_strict(doc):
    """
    Detect A (first alpha glyph) as a true drop cap only if its size >= 1.5× the next alpha glyph size.
    If pattern ABCDEFG is present across paragraphs, rewrite to AE C G (A concatenated to E; E, C, G separated by spaces).
    The heuristic collapses the affected paragraphs into a single paragraph and avoids UI changes.
    """
    paras = list(doc.paragraphs)
    i = 0
    while i < len(paras):
        p = paras[i]
        t = p.text or ""

        first = _dc_first_alpha(p)
        if first is None:
            i += 1; continue
        ri0, ci0, Achar, Asz = first

        # find next alpha glyph size in document flow
        Bchar, Bsz, _ = _dc_next_alpha_in_flow(paras, i, (ri0, ci0))
        if Bsz is None or Asz < 1.5 * Bsz:
            i += 1; continue

        # We expect paragraph i to effectively start with A then spaces then rest (C)
        # Extract C conservatively with a simple pattern on the paragraph's visible text
        mt = re.match(r'^[\s]*([A-Za-z])( +)(.+)$', t)
        if not mt:
            i += 1; continue
        C = mt.group(3).strip()

        # Detect gap D: one or more empty paragraphs after i
        j = i + 1
        had_gap = False
        while j < len(paras) and (paras[j].text or "").strip() == "":
            had_gap = True
            j += 1
        if not had_gap or j >= len(paras):
            i += 1; continue

        E = (paras[j].text or "").strip()

        # Next non-empty after E is G
        k = j + 1
        while k < len(paras) and (paras[k].text or "").strip() == "":
            k += 1
        if k >= len(paras):
            i += 1; continue
        G = (paras[k].text or "").strip()

        # Build AE C G (no space between A and E)
        new_text = (Achar.upper() + E).strip()
        if C:
            new_text += " " + C
        if G:
            new_text += " " + G

        # Replace paragraph i text, then remove paragraphs i+1..k at XML level
        p.text = new_text
        for _ in range(k - i):
            try:
                nxt = p._element.getnext()
                if nxt is not None:
                    p._element.getparent().remove(nxt)
            except Exception:
                break

        # refresh snapshot after structural change
        paras = list(doc.paragraphs)
        i += 1

    return doc
# === end drop-cap heuristic ===



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
    fix_dropcap_reordering_strict(doc)
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
        fix_dropcap_reordering_strict(doc)

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