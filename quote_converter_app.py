# quote_converter_app_hybrid_dropcap_v2.py
import io, os, re, tempfile, streamlit as st

# Optional deps handled gracefully
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from pdf2docx import Converter as PDF2DOCXConverter
except Exception:
    PDF2DOCXConverter = None

try:
    from docx import Document
except Exception:
    Document = None

# ---------------- Common utilities ----------------
def _smart_quotes_us(text: str) -> str:
    if not text:
        return text
    APOS = "<<APOS>>"
    text = re.sub(r"(?<=\\w)[’'](?=\\w)", APOS, text)
    def smarten(line: str) -> str:
        out = []
        open_d = True
        for ch in line:
            if ch == '"':
                out.append("“" if open_d else "”"); open_d = not open_d
            elif ch == "'":
                out.append("’")
            else:
                out.append(ch)
        return "".join(out)
    text = "\\n".join(smarten(ln) for ln in text.split("\\n"))
    return text.replace(APOS, "’")

def _remove_shapes_empty(doc: Document) -> None:
    pkg = doc.part.package
    for part in pkg.parts:
        elt = getattr(part, 'element', None)
        if elt is None:
            continue
        nodes = list(elt.xpath(
            './/*[local-name()="drawing" or local-name()="pict" or local-name()="object" or local-name()="sym" '
            'or local-name()="wsp" or local-name()="txbx" or local-name()="txbxContent"]'
        ))
        for n in nodes:
            parent = n.getparent()
            if parent is not None:
                parent.remove(n)
        # empty runs
        for r in list(elt.xpath('.//*[local-name()="r"]')):
            has_text = bool(r.xpath('.//*[local-name()="t" and normalize-space(text())]'))
            if not has_text and len(r) == 0:
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)
        # empty paragraphs
        for p in list(elt.xpath('.//*[local-name()="p"]')):
            has_text = bool(p.xpath('.//*[local-name()="t" and normalize-space(text())]'))
            has_draw = bool(p.xpath('.//*[local-name()="drawing" or local-name()="pict" or local-name()="object" '
                                    'or local-name()="sym" or local-name()="wsp" or local-name()="txbx" '
                                    'or local-name()="txbxContent"]'))
            if not has_text and not has_draw:
                parent = p.getparent()
                if parent is not None:
                    parent.remove(p)

def _docx_normalize(doc: Document) -> None:
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text:
                r.text = (r.text.replace("\\uFFFC","")
                               .replace("\\u00A0"," ")
                               .replace("\\u000c",""))
                r.text = _smart_quotes_us(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text:
                            r.text = (r.text.replace("\\uFFFC","")
                                           .replace("\\u00A0"," ")
                                           .replace("\\u000c",""))
                            r.text = _smart_quotes_us(r.text)

# ---------------- PyMuPDF paragraph extraction for ALL drop-caps ----------------
def _page_lines_with_spans(page: "fitz.Page"):
    raw = page.get_text("rawdict")
    lines = []
    for b in raw.get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            # collect spans as (text, size, bbox)
            spans = []
            for sp in ln.get("spans", []):
                text = sp.get("text", "")
                if not text:
                    continue
                size = float(sp.get("size", 0.0))
                bbox = sp.get("bbox", [0,0,0,0])
                spans.append((text, size, bbox))
            lines.append(spans)
    return lines

def _is_dropcap_line_from_spans(spans, factor=1.5):
    # Build per-character stream preserving sizes
    chars = []
    for text, sz, _ in spans:
        for ch in text:
            if ch == " " or ch == "\\t":
                continue
            chars.append((ch, sz))
    # find first two printable characters
    letters = [(ch, sz) for ch, sz in chars if ch.strip()]
    if len(letters) < 2:
        return False, None
    first_ch, first_sz = letters[0]
    second_ch, second_sz = letters[1]
    if first_ch.isalpha() and first_sz >= factor * max(second_sz, 0.1):
        return True, first_ch
    return False, None

def _collect_paragraph(page: "fitz.Page", line_index: int):
    raw = page.get_text("rawdict")
    # Rebuild linear list of lines (keep also indent x0)
    lines = []
    for b in raw.get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            text = "".join(sp.get("text", "") for sp in ln.get("spans", []))
            text = re.sub(r"\\s+", " ", text).strip()
            x0 = ln["spans"][0].get("bbox", [0,0,0,0])[0] if ln.get("spans") else 0
            lines.append((text, x0))

    if line_index >= len(lines):
        return ""

    base_indent = lines[line_index][1]
    out_lines = []
    for i in range(line_index, len(lines)):
        t, x0 = lines[i]
        if i > line_index and (t == "" or abs(x0 - base_indent) > 40):
            break
        if t:
            out_lines.append(t)
    paragraph = " ".join(out_lines)
    paragraph = re.sub(r"\\s{2,}", " ", paragraph).strip()
    return paragraph

def extract_all_dropcap_paragraphs(pdf_bytes: bytes, factor=1.5):
    """Return list of extracted paragraphs (strings) for *every* drop-cap encountered in reading order."""
    if fitz is None:
        raise RuntimeError("PyMuPDF not installed.")
    extracted = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for pno in range(len(doc)):
            page = doc[pno]
            lines = _page_lines_with_spans(page)
            # Build an index map to find line indices in raw order
            # We'll reconstruct raw lines again to convert spans index -> raw line index
            raw = page.get_text("rawdict")
            raw_lines = []
            for b in raw.get("blocks", []):
                if b.get("type") != 0:
                    continue
                for ln in b.get("lines", []):
                    raw_lines.append(ln)

            raw_idx = 0
            for i, spans in enumerate(lines):
                is_dc, letter = _is_dropcap_line_from_spans(spans, factor=factor)
                if not is_dc:
                    raw_idx += 1
                    continue
                # Collect the paragraph starting at raw_idx
                para = _collect_paragraph(page, raw_idx)
                if para:
                    extracted.append(para)
                raw_idx += 1
    return extracted

# ---------------- Replace ALL matching paragraphs in DOCX ----------------
DC_OPEN_RE = re.compile(r'^[A-Z]\\s+\\S')  # e.g., "M cave’s ..."
SINGLE_LETTER_RE = re.compile(r'^[A-Z]$')

def replace_all_dropcap_paragraphs(doc: Document, extracted_paras: list) -> int:
    """
    Walk through docx paragraphs; whenever we find a drop-cap form, replace sequentially
    with the next extracted paragraph (reading order). Two cases:
      (1) Paragraph starts "X " (single capital + space): replace that paragraph.
      (2) A single-letter paragraph followed by a lowercase paragraph: merge+replace both.
    Returns number of replacements.
    """
    if not extracted_paras:
        return 0
    i = 0
    replaced = 0
    j = 0  # index into extracted_paras

    while i < len(doc.paragraphs) and j < len(extracted_paras):
        p = doc.paragraphs[i]
        txt = (p.text or "").strip()

        # Case (2): single-letter paragraph with following lowercase-start
        if SINGLE_LETTER_RE.match(txt) and i+1 < len(doc.paragraphs):
            nxt = (doc.paragraphs[i+1].text or "").strip()
            if nxt and nxt[:1].islower():
                doc.paragraphs[i].text = _smart_quotes_us(extracted_paras[j])
                doc.paragraphs[i+1].text = ""  # consume the next paragraph
                replaced += 1
                j += 1
                i += 2
                continue

        # Case (1): inline 'X ' start
        if DC_OPEN_RE.match(txt):
            doc.paragraphs[i].text = _smart_quotes_us(extracted_paras[j])
            replaced += 1
            j += 1
            i += 1
            continue

        i += 1

    return replaced

# ---------------- Streamlit app ----------------
st.set_page_config(page_title="Hybrid Converter (All Drop-caps)", page_icon="📝", layout="centered")
st.title("Hybrid PDF→DOCX: pdf2docx for layout, PyMuPDF for ALL drop-cap paragraphs")
st.caption("Detect every drop cap (first letter ≥ factor × second letter) via PyMuPDF, "
           "extract the full paragraph in reading order, convert whole PDF with pdf2docx, "
           "and replace each matching drop-cap paragraph in the DOCX.")

with st.expander("Options"):
    factor = st.slider("Drop-cap size factor (first ≥ factor × second)", min_value=1.3, max_value=3.0, value=1.5, step=0.1)

uploaded = st.file_uploader("Upload a PDF", type=["pdf"])

if uploaded is not None and st.button("Convert PDF → DOCX (Hybrid, all drop-caps)"):
    try:
        pdf_bytes = uploaded.read()

        # 1) Extract *all* drop-cap paragraphs via PyMuPDF
        extracted = extract_all_dropcap_paragraphs(pdf_bytes, factor=factor)

        # 2) Convert complete PDF with pdf2docx (layout-preserving)
        if PDF2DOCXConverter is None:
            raise RuntimeError("pdf2docx is required.")
        with tempfile.TemporaryDirectory() as tmpd:
            in_pdf = os.path.join(tmpd, "in.pdf")
            out_docx = os.path.join(tmpd, "out.docx")
            with open(in_pdf, "wb") as f:
                f.write(pdf_bytes)
            cv = PDF2DOCXConverter(in_pdf)
            cv.convert(out_docx, start=0, end=None)
            cv.close()
            if Document is None:
                raise RuntimeError("python-docx is required.")
            doc = Document(out_docx)

        # 3) Global cleanup (remove drawing/textbox placeholders etc.)
        _remove_shapes_empty(doc)

        # 4) Replace ALL matching drop-cap paragraphs sequentially
        n_repl = replace_all_dropcap_paragraphs(doc, extracted)

        # 5) Normalise quotes & control chars
        _docx_normalize(doc)

        # 6) Deliver result
        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        st.success(f"Converted. Replaced {n_repl} drop-cap paragraph(s). Download below.")
        st.download_button("Download DOCX (Hybrid, US quotes)", buf.getvalue(),
                           file_name=uploaded.name.rsplit('.',1)[0] + " (Hybrid All Dropcaps US).docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        st.error(f"Conversion failed: {e}")
