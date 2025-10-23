# quote_converter_app.py
import io, re, streamlit as st

# Try import docx
try:
    from docx import Document
except Exception:
    Document = None

_pdf_extractors = []
try:
    from pdfminer.high_level import extract_text as _pdfminer_extract
    def _extract_pdf_pdfminer(data: bytes) -> str:
        return _pdfminer_extract(io.BytesIO(data))
    _pdf_extractors.append(("pdfminer.six", _extract_pdf_pdfminer))
except Exception:
    pass

try:
    import PyPDF2
    def _extract_pdf_pypdf2(data: bytes) -> str:
        reader = PyPDF2.PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    _pdf_extractors.append(("PyPDF2", _extract_pdf_pypdf2))
except Exception:
    pass

try:
    import fitz  # PyMuPDF
    def _extract_pdf_pymupdf(data: bytes) -> str:
        doc = fitz.open(stream=data, filetype="pdf")
        return "\n".join(page.get_text() for page in doc)
    _pdf_extractors.append(("PyMuPDF", _extract_pdf_pymupdf))
except Exception:
    pass

def uk_to_us_quotes(text: str) -> str:
    if not text: return text
    OPEN_S, CLOSE_S, OPEN_D, CLOSE_D, APOS = "<<OPEN_S>>", "<<CLOSE_S>>", "<<OPEN_D>>", "<<CLOSE_D>>", "<<APOS>>"
    text = text.replace("'", "’").replace('"', '”')
    text = (text.replace("‘", OPEN_S).replace("’", CLOSE_S).replace("“", OPEN_D).replace("”", CLOSE_D))
    text = re.sub(r"(?<=\w)"+re.escape(CLOSE_S)+r"(?=\w)", APOS, text)
    for w in ["em","cause","til","tis","twas","sup","round","clock"]:
        text = re.sub(r"\b"+re.escape(CLOSE_S)+w+r"\b", APOS+w, text, flags=re.IGNORECASE)
    text = re.sub(re.escape(CLOSE_S)+r"(?=\d2s\b)", APOS, text)
    text = (text.replace(OPEN_S,"“").replace(CLOSE_S,"”").replace(OPEN_D,"‘").replace(CLOSE_D,"’"))
    return text.replace(APOS,"’")

def convert_docx_bytes_to_us(docx_bytes: bytes) -> bytes:
    if Document is None: raise RuntimeError("Install python-docx")
    doc = Document(io.BytesIO(docx_bytes))
    for p in doc.paragraphs:
        for r in p.runs: r.text = uk_to_us_quotes(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs: r.text = uk_to_us_quotes(r.text)
    out = io.BytesIO(); doc.save(out); return out.getvalue()

def pdf_bytes_to_docx_us(pdf_bytes: bytes) -> bytes:
    if Document is None: raise RuntimeError("Install python-docx")
    text=None
    for name,fn in _pdf_extractors:
        try: text=fn(pdf_bytes); break
        except Exception: continue
    if text is None: text="[PDF extraction failed. Install pdfminer.six / PyPDF2 / PyMuPDF.]"
    text = uk_to_us_quotes(text)
    doc = Document()
    for para in text.split("\n"): doc.add_paragraph(para)
    out = io.BytesIO(); doc.save(out); return out.getvalue()

st.set_page_config(page_title="Quote Style Converter", page_icon="📝", layout="centered")
st.markdown(f"""
<style>
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
</style>
""", unsafe_allow_html=True)

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
