# quote_converter_app_pdf2docx.py
import io
import re
import os
import tempfile
import streamlit as st

# ===== Dependencies =====
try:
    from docx import Document
except Exception:
    Document = None

# pdf2docx for layout-preserving PDF→DOCX
try:
    from pdf2docx import Converter as PDF2DOCXConverter
except Exception:
    PDF2DOCXConverter = None

# ===== XML-safe sanitization =====
_ASCII_CTRL = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')  # keep \t,\n,\r

def _drop_nonchars(s: str) -> str:
    out_chars = []
    for ch in s:
        code = ord(ch)
        if 0xFDD0 <= code <= 0xFDEF:
            continue
        if (code & 0xFFFE) == 0xFFFE:
            continue
        if 0xD800 <= code <= 0xDFFF:
            out_chars.append('\uFFFD'); continue
        out_chars.append(ch)
    return ''.join(out_chars)

def _xml10_filter(text: str) -> str:
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
    return ''.join(out_chars)

def sanitize_for_docx(text: str) -> str:
    if not text:
        return text
    text = _ASCII_CTRL.sub('', text)
    text = _drop_nonchars(text)
    text = _xml10_filter(text)
    return text

# ===== UK → US quotes conversion (placeholder-based) =====

def _detect_primary_style(text: str) -> str:
    """Heuristic: detect whether text is UK-primary (‘…’) or US-primary (“…”).
    Returns "UK", "US", or "UNKNOWN".
    """
    if not text:
        return "UNKNOWN"
    # Count likely opening glyphs
    singles_open = len(re.findall(r'(^|[\s(\[{<])‘', text))
    doubles_open = len(re.findall(r'(^|[\s(\[{<])“', text))
    # Also check frequency overall (ignore apostrophes between letters)
    singles_total = text.count("‘") + text.count("’")
    doubles_total = text.count("“") + text.count("”")
    # Simple thresholds
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
    """Smart normalizer:
       - Preserve apostrophes inside words.
       - If text is already US-primary, do NOT flip it.
       - If text is UK-primary, perform UK→US swap (placeholder-based).
       - For straight quotes, 'smartify' them into US curly quotes using a simple state tracker.
    """
    if not text:
        return text

    # 0) Preserve apostrophes (straight or curly) inside words
    APOS = "<<APOS>>"
    text = re.sub(r"(?<=\w)[’'](?=\w)", APOS, text)

    # 1) Decide if UK→US needed based on existing curly quotes
    style = _detect_primary_style(text)

    # 2) If UK: perform UK→US using the robust placeholder method
    if style == "UK":
        OPEN_S, CLOSE_S, OPEN_D, CLOSE_D = "<<OPEN_S>>", "<<CLOSE_S>>", "<<OPEN_D>>", "<<CLOSE_D>>"
        t = (text.replace("‘", OPEN_S)
                 .replace("’", CLOSE_S)
                 .replace("“", OPEN_D)
                 .replace("”", CLOSE_D))
        # Word-initial elisions and decades: map closing-single placeholders to apostrophes
        # We'll just ensure we don't treat placeholders between letters as quotes
        t = re.sub(r'(?<=\w)'+re.escape(CLOSE_S)+r'(?=\w)', APOS, t)
        for w in ("em","cause","til","tis","twas","sup","round","clock"):
            t = re.sub(r'\b'+re.escape(CLOSE_S)+w+r'\b', APOS+w, t, flags=re.IGNORECASE)
        t = re.sub(re.escape(CLOSE_S)+r'(?=\d{2}s\b)', APOS, t)
        # Swap
        t = (t.replace(OPEN_S,"“")
               .replace(CLOSE_S,"”")
               .replace(OPEN_D,"‘")
               .replace(CLOSE_D,"’"))
        text = t
    else:
        # 3) If US or UNKNOWN: do NOT flip curly quotes.
        # Instead, smarten any remaining straight quotes "..." to US curly.
        # Simple open/close state per line.
        def smarten_line(line: str) -> str:
            out = []
            open_d = True  # assume double-quote starts as opening
            i = 0
            while i < len(line):
                ch = line[i]
                if ch == '"':
                    out.append("“" if open_d else "”")
                    open_d = not open_d
                elif ch == "'":
                    # standalone straight apostrophe -> curly apostrophe
                    out.append("’")
                else:
                    out.append(ch)
                i += 1
            return "".join(out)
        text = "\n".join(smarten_line(ln) for ln in text.split("\n"))

    # 4) Restore apostrophes
    text = text.replace(APOS, "’")
    return text


def uk_to_us_quotes(text: str) -> str:
    if not text:
        return text
    OPEN_S, CLOSE_S, OPEN_D, CLOSE_D, APOS = "<<OPEN_S>>", "<<CLOSE_S>>", "<<OPEN_D>>", "<<CLOSE_D>>", "<<APOS>>"
    text = text.replace("'", "’").replace('"', '”')
    text = (text.replace("‘", OPEN_S)
                .replace("’", CLOSE_S)
                .replace("“", OPEN_D)
                .replace("”", CLOSE_D))
    text = re.sub(r'(?<=\w)'+re.escape(CLOSE_S)+r'(?=\w)', APOS, text)
    for w in ("em","cause","til","tis","twas","sup","round","clock"):
        text = re.sub(r'\b'+re.escape(CLOSE_S)+w+r'\b', APOS+w, text, flags=re.IGNORECASE)
    text = re.sub(re.escape(CLOSE_S)+r'(?=\d{2}s\b)', APOS, text)
    text = (text.replace(OPEN_S,"“")
                .replace(CLOSE_S,"”")
                .replace(OPEN_D,"‘")
                .replace(CLOSE_D,"’"))
    return text.replace(APOS,"’")

def convert_docx_runs_to_us(doc: Document) -> None:
    # In-place conversion of all runs (paragraphs + tables)
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = normalize_quotes_to_us(sanitize_for_docx(r.text))
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = normalize_quotes_to_us(sanitize_for_docx(r.text))

def convert_docx_bytes_to_us(docx_bytes: bytes) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx is required. Add it to requirements.txt")
    doc = Document(io.BytesIO(docx_bytes))
    convert_docx_runs_to_us(doc)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def pdf_bytes_to_docx_using_pdf2docx(pdf_bytes: bytes) -> bytes:
    """Convert PDF→DOCX with pdf2docx (layout preserving), then post-process quotes to US."""
    if Document is None:
        raise RuntimeError("python-docx is required. Add it to requirements.txt")
    if PDF2DOCXConverter is None:
        raise RuntimeError("pdf2docx is required for layout-preserving PDF conversion. Add it to requirements.txt")

    # pdf2docx requires filesystem paths
    with tempfile.TemporaryDirectory() as tmpd:
        pdf_path = os.path.join(tmpd, "in.pdf")
        out_path = os.path.join(tmpd, "out.docx")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        cv = PDF2DOCXConverter(pdf_path)
        cv.convert(out_path, start=0, end=None)
        cv.close()

        # Open produced DOCX, normalize quotes
        doc = Document(out_path)
        convert_docx_runs_to_us(doc)

        # Write to memory
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

# ===== UI =====
st.set_page_config(page_title="Quote Style Converter (pdf2docx)", page_icon="📝", layout="centered")

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

st.title("Quote Style Converter (pdf2docx)")
st.caption("DOCX (UK→US) and PDF→DOCX (layout-preserving via pdf2docx) with US quote normalization.")

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    mode = st.radio("Choose input type", ["DOCX → DOCX (UK → US)", "PDF → DOCX (pdf2docx → US quotes)"])
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
        elif st.button("Convert PDF → DOCX (pdf2docx → US quotes)"):
            try:
                out_bytes = pdf_bytes_to_docx_using_pdf2docx(uploaded.read())
                st.success("Converted. Download below.")
                st.download_button("Download DOCX (US quotes)", out_bytes,
                    file_name=uploaded.name.rsplit(".",1)[0]+" (US Quotes).docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"Conversion failed: {e}")
