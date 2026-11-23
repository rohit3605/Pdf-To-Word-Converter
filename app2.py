import streamlit as st

from converter.engine import convert_pdf_to_docx
from converter.exceptions import ConversionFailedError, EncryptedPDFError, InvalidPDFError

# --------------------------- #
# PAGE CONFIGURATION
# --------------------------- #

st.set_page_config(page_title="PDF → Word Converter", page_icon="📄")

# --------------------------- #
# APP TITLE
# --------------------------- #

st.title("📄 PDF → Word Converter with Images")
st.write("Convert your PDF into an editable Word document (including text and line formatting).")

# --------------------------- #
# FILE UPLOADER
# --------------------------- #

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file and st.button("🚀 Convert to Word"):
    st.info("Converting your PDF to Word (text + images)...")

    try:
        docx_bytes = convert_pdf_to_docx(uploaded_file.getvalue())
    except InvalidPDFError:
        st.error("❌ The uploaded file is empty or is not a valid PDF.")
    except EncryptedPDFError:
        st.error("❌ This PDF is password-protected. Remove the password and try again.")
    except ConversionFailedError:
        st.error("❌ The PDF could not be converted. Please try another file.")
    else:
        st.success("✅ Conversion completed successfully!")
        st.download_button(
            label="📥 Download Word File",
            data=docx_bytes,
            file_name="converted_document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
