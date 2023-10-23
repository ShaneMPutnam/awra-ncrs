import os
import glob
import win32com.client

word = win32com.client.Dispatch("Word.Application")

for pdf_path in glob.glob(os.path.join(os.getcwd(), "*.pdf")):
    doc_path = pdf_path.replace("pdf", "docx")
    document = word.Documents.Open(pdf_path)
    document.SaveAs(doc_path)
    document.Close()

word.Quit()
