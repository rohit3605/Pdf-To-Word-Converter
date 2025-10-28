import streamlit as st
from docx import Document
import fitz  # PyMuPDF
import tempfile
import os
import time

# --------------------------- #
# PAGE CONFIGURATION
# --------------------------- #
st.set_page_config(page_title="PDF → Word Converter", page_icon="📄")

# --------------------------- #
# APP TITLE
# --------------------------- #
st.title("📄 PDF → Word Converter")
st.write("Easily convert your PDF files into editable Word documents (DOCX).")

# --------------------------- #
# FILE UPLOADER
# --------------------------- #
uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.read())
        pdf_path = temp_pdf.name

    if st.button("🚀 Convert to Word"):
        progress = st.progress(0)
        st.info("Converting your PDF to Word...")

        try:
            # Simulate progress bar
            for i in range(0, 100, 15):
                time.sleep(0.1)
                progress.progress(i + 10)

            # Open the PDF
            doc = fitz.open(pdf_path)
            word_doc = Document()

            for page_num, page in enumerate(doc, start=1):
                # Extract text
                text = page.get_text("text") or page.get_text("blocks") or ""

                # Clean non-printable characters
                clean_text = ''.join(ch for ch in text if ch.isprintable())
                clean_text = clean_text.replace('\x00', '').strip()

                if clean_text:
                    try:
                        word_doc.add_paragraph(clean_text)
                    except Exception as e:
                        st.warning(f"⚠️ Skipped a problematic section on page {page_num}: {e}")
                else:
                    word_doc.add_paragraph(f"[No readable text on page {page_num}]")

                word_doc.add_page_break()

            # Save Word file
            output_path = pdf_path.replace(".pdf", ".docx")
            word_doc.save(output_path)

            with open(output_path, "rb") as f:
                st.success("✅ Conversion complete!")
                st.download_button(
                    label="📥 Download Word File",
                    data=f,
                    file_name="converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"❌ Error: {e}")

        finally:
            # Safe cleanup
            try:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                if os.path.exists(output_path):
                    os.remove(output_path)
            except:
                pass
