import docx, easygui

def sanitize_spaces(path):
    found = False
    doc = docx.Document(path)
    for p in doc.paragraphs:
        for r in p.runs:
            if '\xa0' in r.text:
                found = True
                r.text = r.text.replace('\xa0',' ')
    if found:
        doc.save(path)

filedoc = easygui.fileopenbox(
    title="Select a .docx file",
    filetypes=["*.docx"],
    default="*.docx"
)

sanitize_spaces(filedoc)