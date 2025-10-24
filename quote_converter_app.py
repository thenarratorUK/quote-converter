# quote_converter_app_pdf2docx_dropcap_v6.py
import io, os, re, tempfile, streamlit as st

try:
    from docx import Document
    from docx.shared import Pt
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
            out.append('\\uFFFD'); continue
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
    singles_open = len(re.findall(r'(^|[\\s(\\[{<])‘', text))
    doubles_open = len(re.findall(r'(^|[\\s(\\[{<])“', text))
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
    text = re.sub(r"(?<=\\w)[’'](?=\\w)", APOS, text)
    style = _detect_primary_style(text)
    if style == "UK":
        OPEN_S, CLOSE_S, OPEN_D, CLOSE_D = "<<OPEN_S>>", "<<CLOSE_S>>", "<<OPEN_D>>", "<<CLOSE_D>>"
        t = (text.replace("‘", OPEN_S)
                 .replace("’", CLOSE_S)
                 .replace("“", OPEN_D)
                 .replace("”", CLOSE_D))
        t = re.sub(r'(?<=\\w)'+re.escape(CLOSE_S)+r'(?=\\w)', APOS, t)
        for w in ("em","cause","til","tis","twas","sup","round","clock"):
            t = re.sub(r'\\b'+re.escape(CLOSE_S)+w+r'\\b', APOS+w, t, flags=re.IGNORECASE)
        t = re.sub(re.escape(CLOSE_S)+r'(?=\\d{2}s\\b)', APOS, t)
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
        text = "\\n".join(smarten_line(ln) for ln in text.split("\\n"))
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
    pkg = doc.part.package
    for part in pkg.parts:
        elt = getattr(part, 'element', None)
        if elt is None:
            continue
        nodes = list(elt.xpath(
            './/*[local-name()="drawing" or local-name()="pict" or local-name()="object" or local-name()="sym" or local-name()="wsp" or local-name()="txbx" or local-name()="txbxContent"]'
        ))
        for n in nodes:
            parent = n.getparent()
            if parent is not None:
                parent.remove(n)
        # empty runs
        for r in list(elt.xpath('.//*[local-name()="r"]')):
            has_text = bool(r.xpath('.//*[local-name()="t" and normalize-space(text())]'))
            has_children = len(r) > 0
            if not has_text and not has_children:
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)
        # empty paragraphs
        for p in list(elt.xpath('.//*[local-name()="p"]')):
            has_text = bool(p.xpath('.//*[local-name()="t" and normalize-space(text())]'))
            has_draw = bool(p.xpath('.//*[local-name()="drawing" or local-name()="pict" or local-name()="object" or local-name()="sym" or local-name()="wsp" or local-name()="txbx" or local-name()="txbxContent"]'))
            if not has_text and not has_draw:
                parent = p.getparent()
                if parent is not None:
                    parent.remove(p)

# === Drop-cap reconstruction (replacement) ===
def _median_font_size(p):
    sizes = []
    for r in p.runs:
        if r.font.size:
            sizes.append(r.font.size.pt)
    return sorted(sizes)[len(sizes)//2] if sizes else None

def _detect_dropcap(paras, start_idx=0):
    # Mode A: paragraph begins with 'X ' (letter + space)
    for j in range(start_idx, min(start_idx+6, len(paras))):
        txt = (paras[j].text or "").lstrip()
        if re.match(r'^[A-Z]\\s\\S', txt):
            return j, txt[0].upper()
    # Mode B: oversized single-letter run
    for j in range(start_idx, min(start_idx+6, len(paras))):
        p = paras[j]
        if not p.runs:
            continue
        med = _median_font_size(p) or 0
        k = next((i for i,r in enumerate(p.runs) if r.text and r.text.strip()), None)
        if k is None:
            continue
        r = p.runs[k]
        t = r.text.strip()
        if len(t) == 1 and t.isalpha():
            size = (r.font.size.pt if r.font.size else med or 0)
            if size >= max(20, 1.6*(med or 12)):
                return j, t.upper()
    return None, None

def _strip_leading_same_letter(text, letter):
    if not text:
        return text
    left_spaces = len(text) - len(text.lstrip())
    s = text[left_spaces:]
    if not s:
        return text
    if s[0].upper() == letter:
        s = s[1:]
        return (" " * left_spaces) + s
    return text

def reconstruct_dropcap_block(paras, start_window=10, max_lines=4):
    if not paras:
        return False
    idx, letter = _detect_dropcap(paras, 0)
    if letter is None:
        return False

    lines, line_idxs = [], []
    j = idx

    first_txt = (paras[idx].text or "").strip()
    if re.match(r'^[A-Z]\\s\\S', first_txt):
        lines.append(first_txt)
        line_idxs.append(idx)
        j = idx + 1
    else:
        j = idx + 1

    while j < min(idx + 1 + start_window, len(paras)) and len(lines) < max_lines:
        t = (paras[j].text or "").strip()
        if not t:
            j += 1; continue
        # Skip short all-caps headings without punctuation (e.g., TOC/headers)
        if len(t) <= 50 and t.isupper() and not re.search(r'[.!?…]', t):
            j += 1; continue
        if len(t) <= 180:
            lines.append(t)
            line_idxs.append(j)
            if re.search(r'[.!?…]"?$', t):
                break
            j += 1; continue
        break

    if not lines:
        return False

    first_raw = lines[0]
    if first_raw.lstrip()[:1].upper() == letter:
        merged = first_raw
    else:
        merged = letter + _strip_leading_same_letter(first_raw, letter)

    for t in lines[1:]:
        merged += " " + _strip_leading_same_letter(t, letter)

    paras[line_idxs[0]].text = merged
    for k in line_idxs[1:]:
        paras[k].text = ""
    if idx not in line_idxs:
        paras[idx].text = ""

    return True

def convert_docx_bytes_to_us(docx_bytes: bytes) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx required.")
    doc = Document(io.BytesIO(docx_bytes))
    convert_docx_runs_to_us(doc)
    out = io.BytesIO(); doc.save(out)
    return out.getvalue()

def pdf_bytes_to_docx_using_pdf2docx(pdf_bytes: bytes, fix_dropcaps: bool=True) -> bytes:
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

        _remove_global_shapes_all_parts(doc)

        paras = doc.paragraphs
        for i, p in enumerate(paras):
            for r in p.runs:
                if r.text:
                    r.text = (r.text.replace("\\uFFFC","")
                                   .replace("\\u00A0"," ")
                                   .replace("\\u000c",""))
        for i, p in enumerate(paras):
            if p.text.strip() in {"", "\\u00A0"} and 0 < i < len(paras)-1:
                prev = paras[i-1].text.strip()
                nxt  = paras[i+1].text.strip()
                if prev and nxt and not re.search(r'[.!?…]"?$', prev):
                    p.text = ""

        if fix_dropcaps:
            reconstruct_dropcap_block(doc.paragraphs, start_window=10, max_lines=4)

        convert_docx_runs_to_us(doc)

        buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

st.set_page_config(page_title="Quote Style Converter (Drop-cap v6)", page_icon="📝", layout="centered")

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
st.markdown("<style>\\n"+CSS+"\\n</style>", unsafe_allow_html=True)

st.title("Quote Style Converter (pdf2docx – Drop-cap v6)")
st.caption("PDF→DOCX with US quotes, global square cleanup, and replaced multi-line drop-cap reconstruction.")

with st.container():
    mode = st.radio("Choose input type", ["DOCX → DOCX (UK → US)", "PDF → DOCX (pdf2docx → US quotes)"])
    uploaded = st.file_uploader("Upload file", type=["docx","pdf"])
    fix_dc = st.checkbox("Fix decorative drop caps", value=True)

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
                out_bytes = pdf_bytes_to_docx_using_pdf2docx(uploaded.read(), fix_dropcaps=fix_dc)
                st.success("Converted. Download below.")
                st.download_button("Download DOCX (US quotes)", out_bytes,
                    file_name=uploaded.name.rsplit(".",1)[0]+" (US Quotes).docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"Conversion failed: {e}")
