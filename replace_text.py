import easygui, os, docx
import ps_docx_utils as pdu

def find_docx_files(folder_path):
    """Finds all *.docx files in the specified folder and its subfolders.

    Args:
        folder_path: The path to the folder where the search should begin.

    Returns:
        A list of full paths to all the found *.docx files.
    """

    docx_files = []
    for root, _, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith(".docx"):
                full_path = os.path.join(root, filename)
                docx_files.append(full_path)
    return docx_files

def get_text_mapping(path):
    mapping = []
    state = None
    doc = docx.Document(path)
    for p in doc.paragraphs:
        if p.style.name == 'Heading 1':
            mapping.append({'header':p.text,'search':[],'replace':[]})
            state = None
            continue
            
        if p.text in ('Search','Replace'):
            state = p.text.lower()
            continue
        if state:
            mapping[-1][state].append(p)
    return mapping

def process_file(path):
    doc = docx.Document(path)

    for tm in text_mapping:
        found_start = False
        found_end = False
        i=0
        for p in doc.paragraphs:
            if i<len(tm["search"]) and p.text == tm["search"][i].text:
                if not found_start and i==0:
                    found_start = True
                if found_start:
                    pdu.delete_paragraph(p)
                i+=1
            elif found_start:
                found_start = False
                found_end = True

            if found_end:
                pdu.copy_paragraph_list_before(p,tm["replace"])
                found_end = False
                i=0
                print(f'{path}: inserted {tm["header"]}')

    doc.save(path)

folder_path = easygui.diropenbox(title="Select Folder")    # Open folder selection dialog

text_mapping = get_text_mapping('docx_resources\\заміна-рубрики.docx')

if folder_path:    # Check if a folder was selected
    docx_files = find_docx_files(folder_path)
    #if docx_files:
    #    print(docx_files)

for filename in docx_files:
    process_file(filename)

