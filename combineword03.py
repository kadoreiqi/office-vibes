import os
import win32com.client as win32
from tkinter import filedialog, Tk
import traceback

def copy_doc_content(word, doc, output_doc):
    """Copy all content from input doc to output doc"""
    try:
        # Copy all content (text, tables, etc.)
        doc.Activate()
        word.Selection.WholeStory()
        word.Selection.Copy()

        # Paste into output document
        output_doc.Activate()
        word.Selection.Paste()
        word.Selection.InsertBreak(7)  # Insert page break (wdPageBreak)
    except Exception as e:
        print(f"❌ Failed to copy content from {doc.Name}: {e}")

def merge_docs(docx_paths, output_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0  # Suppress dialogs

    try:
        output_doc = word.Documents.Add()  # New blank document to combine pages into

        docs = []
        for path in docx_paths:
            print(f"📄 Opening: {os.path.basename(path)}")
            docs.append(word.Documents.Open(path))

        # Merge all documents (content-wise, no page selection)
        for doc in docs:
            copy_doc_content(word, doc, output_doc)

        for doc in docs:
            doc.Close(False)

        output_doc.SaveAs(output_path)
        print(f"✅ Merged document saved to: {output_path}")
        output_doc.Close()

    except Exception as e:
        print("❌ Error during merging:")
        traceback.print_exc()

    finally:
        word.Quit()

def main():
    root = Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Select folder with DOCX files")
    if not folder:
        print("No folder selected.")
        return

    docx_files = sorted([
        os.path.join(folder, f) for f in os.listdir(folder)
        if f.lower().endswith(".docx")
    ])

    if not docx_files:
        print("❌ No DOCX files found in folder.")
        return

    output_path = filedialog.asksaveasfilename(
        title="Save merged DOCX file as",
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not output_path:
        print("No output path selected.")
        return

    merge_docs(docx_files, output_path)

if __name__ == "__main__":
    main()
input()
