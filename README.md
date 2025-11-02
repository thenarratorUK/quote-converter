## README: quotescript.streamlit.app

### PDF to DOCX Converter with Quote Correction 

#### Overview

This web app converts and cleans manuscripts in preparation for use with the Scripter app ([https://scripter.streamlit.app](https://scripter.streamlit.app))

As the Scripter app requires a DOCX using international (double) speech marks, if your manuscript is in PDF or uses British (single) speech marks (DOCX or PDF, this app must be used to convert the document from PDF to DOCX (correcting the quote style if necessary) or can be used to convert a British-styled DOCX to an American styled DOCX 

The app is available at [**https://quotescript.streamlit.app**](https://quotescript.streamlit.app) and runs directly in your browser.

---

### How to Use

#### 1. Upload a File

- Click **Browse files** or drag your file into the upload box.
- Supported inputs:
  - **Word document (.docx)** — runs quote-standardisation directly.
  - **PDF file (.pdf)** — converts to DOCX automatically, then cleans and repairs.
- Example input filenames:
  - `BookTitle.docx`
  - `BookTitle.pdf`

When processing, the app automatically creates a temporary working DOCX before cleaning.

---

#### 2. Quote Normalisation

- Converts all quotation marks to **US smart quotes** (“ ” and ‘ ’).
- Correctly distinguishes between:
  - Opening and closing quotes.
  - In-word apostrophes (e.g., *don’t*, *it’s*).
  - Leading apostrophes in contractions (e.g., *’em*, *’til*, *’tis*).
- Rewrites only the quote characters; other formatting (bold, italics, etc.) is preserved.
- Works across all text runs in the DOCX, ensuring uniform results.
- If converting from PDF to DOCX, controls for drop caps (a single larger letter at the start of a chapter) and the revealing of previously invisible artefacts that may otherwise interfere with readability. 
- Once finished - this can take a few minutes, during which an indicator at the top of the page will be the only sign something is still happening - the app displays a message:
  - **“Converted. Download below.”**
  - A **Download File** button appears to save your cleaned DOCX.

---

#### 3.  Download and Review

- Click **Download File** to save the cleaned document to your computer.
- The file will have the same base name as your input, ending in `.docx`.
  - Example:
    - Input: `BookTitle.pdf`
    - Output: `BookTitle.docx`
- Open the resulting file in Microsoft Word or your editor of choice to confirm:
  - Quotation marks are standardised.
  - Paragraph flow and punctuation are correct.
  - Decorative or boxed artefacts are removed.

---

### Key Behaviours

- Does not modify your original file; it always produces a new DOCX.
- Keeps all paragraph and style structure intact except for necessary repairs.
- Skips code blocks or other non-textual objects that might contain deliberate straight quotes.
- Attempts to preserve italics, headings, and inline formatting wherever possible.

---

### Practical Notes

- If your script contains poetry, song lyrics, or code snippets, skim the cleaned file to ensure intended straight quotes remain unchanged.
- Multi-column PDFs or those with complex layouts may still require manual correction after conversion.
- Decorative or embedded image shapes may remain; these are not removed automatically.
- For best results, always start with the highest-quality source file available.
