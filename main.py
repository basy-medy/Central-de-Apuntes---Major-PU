import os
import comtypes.client

def convert_to_pdf(doc_path, pdf_path):
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # PDF format
        doc.Close()
        print(f"Converted: {doc_path} -> {pdf_path}")
    except Exception as e:
        print(f"Failed to convert {doc_path}: {e}")
    finally:
        word.Quit()

def find_and_convert_docs(base_folder):
    for foldername, subfolders, filenames in os.walk(base_folder):
        for filename in filenames:
            if filename.lower().endswith(('.docx', '.doc')):
                doc_path = os.path.join(foldername, filename)
                pdf_path = os.path.splitext(doc_path)[0] + ".pdf"
                
                if not os.path.exists(pdf_path):  # Skip if PDF already exists
                    convert_to_pdf(doc_path, pdf_path)
                else:
                    print(f"PDF already exists: {pdf_path}")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    find_and_convert_docs(current_dir)
