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



def _run_size_pt(run, para, default=11.0):
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

def _alpha_positions_with_sizes(paragraph):
    for ri, run in enumerate(paragraph.runs):
        t = run.text or ""
        if not t:
            continue
        size = _run_size_pt(run, paragraph)
        for ci, ch in enumerate(t):
            if ch.isalpha():
                yield ri, ci, ch, size

def fix_dropcap_reordering_strict(doc):
    """
    Same ABCDEFG -> AE C G reordering as fix_dropcap_reordering, but guarded by a font-size check:
    Only trigger when the *first alphabetic char in the initial paragraph* is >= 1.5× the size of the *next* alphabetic char in the document flow.
    """
    paras = list(doc.paragraphs)
    i = 0
    while i < len(paras):
        p = paras[i]
        t = p.text or ""

        # Identify the first alphabetic char position in paragraph i and its size
        first_alpha = None
        for ri, ci, ch, sz in _alpha_positions_with_sizes(p):
            first_alpha = (ri, ci, ch, sz)
            break
        if first_alpha is None:
            i += 1
            continue

        ri0, ci0, Achar, Asz = first_alpha

        # Gather the next alphabetic char in the *document flow* (same paragraph after A, or subsequent paragraphs)
        next_alpha = None
        # First: same paragraph, after ci0
        for ri, ci, ch, sz in _alpha_positions_with_sizes(p):
            if (ri > ri0) or (ri == ri0 and ci > ci0):
                next_alpha = (ri, ci, ch, sz)
                break
        # If not found, look ahead across following paragraphs
        if next_alpha is None:
            for j in range(i + 1, len(paras)):
                for ri, ci, ch, sz in _alpha_positions_with_sizes(paras[j]):
                    next_alpha = (ri, ci, ch, sz)
                    break
                if next_alpha is not None:
                    break
        if next_alpha is None:
            i += 1
            continue

        _, _, _, Bsz = next_alpha

        # Size guard: A must be >= 1.5× B to qualify
        if Asz < 1.5 * Bsz:
            i += 1
            continue

        # Validate local paragraph layout: paragraph i should begin with A + space + rest-of-line (C)
        # We won't rely only on regex; we reconstruct C by slicing text from the first visible non-space char after A
        # However, to avoid being too invasive, require that the very first visible glyph is Achar and a following space exists nearby.
        starts_ok = False
        visible_seen = 0
        for ch in (p.text or ""):
            if not ch.isspace():
                visible_seen += 1
                if visible_seen == 1 and ch == Achar:
                    starts_ok = True
                break
        if not starts_ok:
            i += 1
            continue

        # Extract C string: from the first non-space char after the *first visible run char* (not strictly after ci0, but good-enough for split shapes)
        # We'll parse as: leading single glyph + optional spaces + remainder -> C
        mt = re.match(r'^([^\S\r\n]*)([A-Za-z])( +)(.+)$', p.text or "")
        if not mt:
            i += 1
            continue
        # groups: lead_ws, A, space, C
        C = mt.group(4).strip()

        # Detect gap D: one or more empty paragraphs immediately following
        j = i + 1
        had_gap = False
        while j < len(paras) and (paras[j].text or "").strip() == "":
            had_gap = True
            j += 1
        if not had_gap or j >= len(paras):
            i += 1
            continue

        # E = first non-empty after the gap
        E = (paras[j].text or "").strip()

        # G = next non-empty after E (skip any additional empties)
        k = j + 1
        while k < len(paras) and (paras[k].text or "").strip() == "":
            k += 1
        if k >= len(paras):
            i += 1
            continue
        G = (paras[k].text or "").strip()

        # Construct AE C G (no space between A and E)
        new_text = (Achar.upper() + E).strip()
        if C:
            new_text += " " + C
        if G:
            new_text += " " + G

        # Replace paragraphs i..k with a single paragraph containing new_text
        p.text = new_text
        # Remove paras i+1..k at XML level
        for _ in range(k - i):
            try:
                nxt = p._element.getnext()
                if nxt is not None:
                    p._element.getparent().remove(nxt)
            except Exception:
                break

        paras = list(doc.paragraphs)
        i += 1

    return doc

def fix_dropcap_reordering_strict(doc):
    """
    Heuristic for PDF->DOCX line misordering caused by drop caps.
    Pattern described by user:
      A = single larger letter (captured in first paragraph as a single letter + space prefix)
      B = space after it
      C = following text (continuation of first paragraph after "A ")
      D = a gap (one or more empty paragraphs)
      E = the sentence that actually belongs immediately after the drop cap
      F = a newline / next line break (typically next paragraph)
      G = the following line
    Transform:
      ABCDEFG  ->  A+E + " " + C + " " + G
    Implementation assumptions:
      - A and C are in paragraph i, with text starting "^[A-Za-z]\\s+..."
      - D are one or more empty paragraphs i+1..j-1
      - E is the first non-empty paragraph at index j
      - G is the first non-empty paragraph at index k > j (skipping any further empty paragraphs)
    The function replaces paragraphs [i..k] with a single paragraph "AE C G".
    """
    paras = list(doc.paragraphs)
    i = 0
    while i < len(paras):
        p = paras[i]
        t = p.text or ""
        m = re.match(r'^([A-Za-z])\s+(.+)$', t)
        if not m:
            i += 1
            continue

        A = m.group(1)
        C = m.group(2).strip()

        # Find first non-empty paragraph after optional gap(s)
        j = i + 1
        had_gap = False
        while j < len(paras) and (paras[j].text or "").strip() == "":
            had_gap = True
            j += 1

        if j >= len(paras):
            i += 1
            continue

        E = (paras[j].text or "").strip()

        # Find next non-empty after E (skip additional gaps)
        k = j + 1
        while k < len(paras) and (paras[k].text or "").strip() == "":
            k += 1
        if k >= len(paras):
            i += 1
            continue

        G = (paras[k].text or "").strip()

        # We only apply if we actually observed a gap D; this reduces false positives
        if not had_gap:
            i += 1
            continue

        # Construct new paragraph text
        new_text = (A + E).strip()  # AE, no space between A and E
        if C:
            new_text += " " + C
        if G:
            new_text += " " + G

        # Replace paragraphs i..k with a single paragraph containing new_text
        p.text = new_text
        # Remove paragraphs i+1..k from the document
        for _ in range(k - i):
            # Note: python-docx requires removing from the underlying XML
            # We'll remove the next sibling paragraph element
            try:
                nxt = p._element.getnext()
                if nxt is not None:
                    p._element.getparent().remove(nxt)
            except Exception:
                break

        # Refresh local reference to paragraphs by re-snapshotting (structure changed)
        paras = list(doc.paragraphs)
        # Move past the modified paragraph
        i += 1

    return doc



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
