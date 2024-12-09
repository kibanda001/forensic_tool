import PyPDF2

# Ouvrir le fichier PDF en mode binaire
with open("SHORT COURSES.pdf", "rb") as file:
    pdf_metadata = PyPDF2.PdfReader(file)
    doc_info = pdf_metadata.metadata
    for key, value in doc_info.items():
        print(f"[+] {key}: {value}")

