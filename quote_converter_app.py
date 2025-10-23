# quote_converter_app_clean.py
import io
import re
import sys
import streamlit as st

# ===== Dependencies =====
try:
    from docx import Document
except Exception:
    Document = None

# Optional PDF extractors (use first available)
_PDF_EXTRACTORS = []
try:
    from pdfminer.high_level import extract_text as _pdfminer_extract
    def _extract_pdf_pdfminer(data: bytes) -> str:
        return _pdfminer_extract(io.BytesIO(data))
    _PDF_EXTRACTORS.append(("pdfminer.six", _extract_pdf_pdfminer))
except Exception:
    pass

try:
    import PyPDF2
    def _extract_pdf_pypdf2(data: bytes) -> str:
        reader = PyPDF2.PdfReader(io.BytesIO(data))
        chunks = []
        for page in reader.pages:
            t = page.extract_text() or ""
            chunks.append(t)
        return "\n".join(chunks)
    _PDF_EXTRACTORS.append(("PyPDF2", _extract_pdf_pypdf2))
except Exception:
    pass

try:
    import fitz  # PyMuPDF
    def _extract_pdf_pymupdf(data: bytes) -> str:
        doc = fitz.open(stream=data, filetype="pdf")
        chunks = []
        for page in doc:
            chunks.append(page.get_text())
        return "\n".join(chunks)
    _PDF_EXTRACTORS.append(("PyMuPDF", _extract_pdf_pymupdf))
except Exception:
    pass

# ===== XML-safe sanitization (NO recursion) =====
# Precompiled patterns for speed and safety
_ASCII_CTRL = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')  # keep \t,\n,\r
# Unicode noncharacters U+FDD0..U+FDEF or any ...FFFE/FFFF
def _drop_nonchars(s: str) -> str:
    out_chars = []
    for ch in s:
        code = ord(ch)
        if 0xFDD0 <= code <= 0xFDEF:
            continue
        if (code & 0xFFFE) == 0xFFFE:
            continue
        # Replace surrogate code points (should not appear in valid str)
        if 0xD800 <= code <= 0xDFFF:
            out_chars.append('\uFFFD')
            continue
        out_chars.append(ch)
    return ''.join(out_chars)

def _xml10_filter(text: str) -> str:
    """Keep only XML 1.0 valid chars: #x9 | #xA | #xD | #x20–#xD7FF | #xE000–#xFFFD | #x10000–#x10FFFF."""
    if not text:
        return text
    out_chars = []
    for ch in text:
        code = ord(ch)
        if code in (0x9, 0xA, 0xD):
            out_chars.append(ch); continue
        if 0x20 <= code <= 0xD7FF:
            out_chars.append(ch); continue
        if 0xE000 <= code <= 0xFFFD:
            out_chars.append(ch); continue
        if 0x10000 <= code <= 0x10FFFF:
            out_chars.append(ch); continue
        # else: drop
    return ''.join(out_chars)

def sanitize_for_docx(text: str) -> str:
    if not text:
        return text
    text = _ASCII_CTRL.sub('', text)
    text = _drop_nonchars(text)
    text = _xml10_filter(text)
    return text

# ===== UK → US quotes conversion (placeholder-based; NO recursion) =====
def uk_to_us_quotes(text: str) -> str:
    if not text:
        return text
    OPEN_S, CLOSE_S, OPEN_D, CLOSE_D, APOS = "<<OPEN_S>>", "<<CLOSE_S>>", "<<OPEN_D>>", "<<CLOSE_D>>", "<<APOS>>"
    # Normalize straight quotes to curly to simplify matching
    text = text.replace("'", "’").replace('"', '”')
    # Tokenize curly quotes
    text = (text.replace("‘", OPEN_S)
                .replace("’", CLOSE_S)
                .replace("“", OPEN_D)
                .replace("”", CLOSE_D))
    # Apostrophes inside words (letters ’ letters) => APOS
    text = re.sub(r'(?<=\w)'+re.escape(CLOSE_S)+r'(?=\w)', APOS, text)
    # Word-initial elisions (case-insensitive)
    for w in ("em","cause","til","tis","twas","sup","round","clock"):
        text = re.sub(r'\b'+re.escape(CLOSE_S)+w+r'\b', APOS+w, text, flags=re.IGNORECASE)
    # Decades like ’90s
    text = re.sub(re.escape(CLOSE_S)+r'(?=\d{2}s\b)', APOS, text)
    # Swap single↔double (primary↔nested)
    text = (text.replace(OPEN_S,"“")
                .replace(CLOSE_S,"”")
                .replace(OPEN_D,"‘")
                .replace(CLOSE_D,"’"))
    # Restore apostrophes
    return text.replace(APOS,"’")

def convert_docx_bytes_to_us(docx_bytes: bytes) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx is required. Add it to requirements.txt")
    doc = Document(io.BytesIO(docx_bytes))
    # Process paragraphs
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = uk_to_us_quotes(sanitize_for_docx(r.text))
    # Process tables
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = uk_to_us_quotes(sanitize_for_docx(r.text))
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def pdf_bytes_to_docx_us(pdf_bytes: bytes) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx is required. Add it to requirements.txt")
    # Extract text with first available backend
    text = None
    used = None
    for name, fn in _PDF_EXTRACTORS:
        try:
            text = fn(pdf_bytes)
            used = name
            break
        except Exception:
            continue
    if text is None:
        text = "[PDF text extraction failed. Install pdfminer.six, PyPDF2, or PyMuPDF and retry.]"
    # Sanitize then normalize quotes
    text = sanitize_for_docx(text)
    text = uk_to_us_quotes(text)
    # Build DOCX line-by-line with an additional XML 1.0 filter at write time
    doc = Document()
    for para in text.split("\n"):
        clean = _xml10_filter(para)
        doc.add_paragraph(clean)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ===== UI =====
st.set_page_config(page_title="Quote Style Converter", page_icon="📝", layout="centered")

CSS = """
:root { --primary-color: #008080; --primary-hover: #006666; --bg-1: #0b0f14; --bg-2: #11161d; --card: #0f141a; --text-1: #e8eef5; --text-2: #b2c0cf; --muted: #8aa0b5; --accent: #e0f2f1; --ring: rgba(0, 128, 128, 0.5); }
html, body, [data-testid="stAppViewContainer"] { background: linear-gradient(180deg, var(--bg-1), var(--bg-2)) !important; color: var(--text-1) !important; }
a { color: var(--accent) !important; }
.card { background: var(--card); border: 1px solid rgba(255,255,255,.08); border-radius: 1rem; padding: 1rem 1.25rem; margin: 0.5rem 0 1.25rem 0; box-shadow: 0 10px 25px rgba(0,0,0,.25); }
div.stButton > button { background-color: var(--primary-color); color: #e8eef5; border: none; border-radius: 0.6rem; padding: 0.6rem 1rem; }
div.stButton > button:hover { background-color: var(--primary-hover); }
body { font-family: Avenir, sans-serif; line-height: 1.65; }
.kicker { text-transform: uppercase; letter-spacing: .1em; font-weight: 600; color: #9cc; }
h1, h2, h3 { letter-spacing: .02em; }
.pill { display: inline-block; padding: .2rem .6rem; border: 1px solid rgba(255,255,255,.2); border-radius: 999px; font-size: .85rem; color: #e8eef5; }
.muted { color: #b2c0cf; }
.mono { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; }
"""
st.markdown("<style>\n" + CSS + "\n</style>", unsafe_allow_html=True)

st.title("Quote Style Converter")
st.caption("Upload a DOCX (UK quotes) → DOCX (US quotes). Or upload a PDF → DOCX (US quotes).")

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    mode = st.radio("Choose input type", ["DOCX → DOCX (UK → US)", "PDF → DOCX (→ US)"])
    uploaded = st.file_uploader("Upload file", type=["docx","pdf"])
    st.markdown('</div>', unsafe_allow_html=True)

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
        elif st.button("Convert PDF → DOCX (US quotes)"):
            try:
                out_bytes = pdf_bytes_to_docx_us(uploaded.read())
                st.success("Converted. Download below.")
                st.download_button("Download DOCX (US quotes)", out_bytes,
                    file_name=uploaded.name.rsplit(".",1)[0]+" (US Quotes).docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.caption("Note: PDF extraction quality depends on the PDF and installed libraries.")
            except Exception as e:
                st.error(f"Conversion failed: {e}")
