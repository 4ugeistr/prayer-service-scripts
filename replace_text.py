import easygui, os, docx, logging
from datetime import datetime
import ps_docx_utils as pdu

start_time = datetime.now()


logging.basicConfig(filename=f'logs\\replace_text_{datetime.now().strftime("%H%M%S")}.log',filemode = 'w', level=logging.INFO)
logger = logging.getLogger(__name__)

'''
Finds all *.docx files in the specified folder and its subfolders.
'''
def find_docx_files(folder_path):
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
    print(f'Processing: {path}')
    logger.info(f'Processing: {path}')
    file_changed=False

    #dbg
    prok = False

    for tm in text_mapping:

        #dbg
        if 'Псалом' in tm['search'][0].text:
            logger.info(f"___Probing {tm['search'][0].text}")
            logger.info(f"___Strings qty = {len(tm['search'])}")
            prok = True
            #logger.info
        found_start = False
        found_end = False
        n=0
        i=0
        p_buffer=[]
        for p in doc.paragraphs:
            #logger.info(p.text)

            if prok and i ==0:
                logger.info(p.text.lower().replace('  ',' ').strip())
                logger.info(tm["search"][i].text.lower().replace('  ',' ').strip())
                logger.info(p.text.lower().replace('  ',' ').strip() == tm["search"][i].text.lower().replace('  ',' ').strip())


            if i<len(tm["search"]) and p.text.lower().replace('  ',' ').strip() == tm["search"][i].text.lower().replace('  ',' ').strip():
                if not found_start and i==0:
                    logger.info("NEW START"+p.text)
                    found_start = True
                if found_start:
                    p_buffer.append(p)
                i+=1
            elif found_start and i==len(tm["search"]):
                found_start = False
                found_end = True
            
            elif found_start:
                logger.info(f"WARNING! Not a full match for {tm['header']}. Found {i}/{len(tm['search'])}")
                #logger.info(f"{n} {p.text}")
                #logger.info(f"{n} {tm['search'][i].text}")
                found_start = False
                p_buffer=[]
                i=0

            if found_end:
                file_changed=True
                if not(i==1 and p.text == tm["search"][0].text):
                    for item in p_buffer:
                        pdu.delete_paragraph(item)
                    p_buffer=[]
                    pdu.copy_paragraph_list_before(p,tm["replace"])
                    found_end = False
                    i=0
                    logger.info(f'{path}: inserted {tm["header"]}')
            n+=1
        #dbg
        prok = False
    if file_changed:
        doc.save(path)

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

folder_path = easygui.diropenbox(title="Select Folder")    # Open folder selection dialog

text_mapping = get_text_mapping('docx_resources\\заміна-рубрики.docx')
for tm in text_mapping:
    #for p in tm['replace']:
    if tm['replace'][-1].text=='':
        print(tm)
        raise Exception

if folder_path:    # Check if a folder was selected
    docx_files = find_docx_files(folder_path)
    #if docx_files:
    #    print(docx_files)

for filename in docx_files:
    sanitize_spaces(filename)
    process_file(filename)

print('Done!')

end_time = datetime.now()
elapsed_time = (end_time - start_time).total_seconds()
print(f"Elapsed time : {elapsed_time}") 
