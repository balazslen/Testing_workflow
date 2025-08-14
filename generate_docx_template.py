from docx import Document
import os

def create_word_template():
    doc = Document()
    doc.add_heading('Executive Summary', level=1)
    doc.add_paragraph(
        "Paste your executive summary content here.\n\n"
        "This document was automatically generated using Python."
    )
    doc.add_heading('Key Points', level=2)
    doc.add_paragraph("- Point 1\n- Point 2\n- Point 3")

    filename = "Executive_Summary_Template.docx"
    doc.save(filename)
    print(f"✅ Word document created and saved as '{filename}' in: {os.getcwd()}")

if __name__ == "__main__":
    create_word_template()


