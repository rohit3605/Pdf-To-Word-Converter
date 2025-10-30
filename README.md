📄 PDF → Word Converter (with Images)

A simple and powerful Streamlit web app that converts PDF files into editable Word (DOCX) documents — including both text and images.

⚙️ Built with Python, Streamlit, and PyMuPDF — lightweight, fast, and deployable anywhere!

🚀 Features

✅ Extracts clean text from PDF pages
✅ Preserves images inside the Word file
✅ Progress bar for conversion process
✅ Works on both desktop & mobile
✅ One-click download after conversion
✅ Safe cleanup of temporary files

🧠 Tech Stack

Python 3.8+

Streamlit – for web UI

PyMuPDF (fitz) – for PDF reading and image extraction

python-docx – for Word document creation

🏗️ How It Works

Upload your PDF file (.pdf)

Click Convert to Word

The app:

Reads each page of the PDF

Extracts text (cleaning invalid XML characters)

Extracts embedded images

Adds both text and images to a Word document

You can then download the converted Word file
