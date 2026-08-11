import easygui, os, docx, logging, csv
from datetime import datetime
import ps_docx_utils as pdu

start_time = datetime.now()


logging.basicConfig(filename=f'..\\ps_drafts\\logs\\replace_word_{datetime.now().strftime("%H%M%S")}.log',filemode = 'w', level=logging.INFO)
logger = logging.getLogger(__name__)

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


def get_matrix_full(csv_filename):
    matrix=[]
    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            matrix.append(row)
    return matrix

def process_file(path):
    doc = docx.Document(path)
    logger.info(f"Опрацьовуємо {path}")
    print(f'Processing: {path}')
    for tm in text_mapping:
        n=0
        for p in doc.paragraphs:
            for r in p.runs:
                if tm[0] in r.text:
                    logger.info(f"Заміна {tm[0]} на {tm[1]} в параграфі {n}:")
                    logger.info(p.text)
                    r.text = r.text.replace(tm[0], tm[1])
            n+=1
    doc.save(path)

folder_path = easygui.diropenbox(title="Select Folder")    # Open folder selection dialog

text_mapping = get_matrix_full('docx_resources\\заміна-слів.csv')

if folder_path:    # Check if a folder was selected
    docx_files = find_docx_files(folder_path)

for filename in docx_files:
    process_file(filename)

print('Done!')

end_time = datetime.now()
elapsed_time = (end_time - start_time).total_seconds()
print(f"Elapsed time : {elapsed_time}") 