import os
import ps_docx_utils as du
import easygui as eg

import docx
from docx.enum.text import WD_COLOR_INDEX

from sys import exit

day_name_sequence = ["НД","ПН","ВВ","СР","ЧТ","ПТ","СБ",]

day_short_dic={"ПН":1,
        "ВВ":2,
        "СР":3,
        "ЧТ":4,
        "ПТ":5,
        "СБ":6,
        "НД":7}
day_short_dic_reversed = {v:k for k,v in day_short_dic.items()}
day_short_dic_string='('+'|'.join(day_short_dic.keys())+')'

day_dic = {"Понеділок":1,
            "Вівторок":2,
            "Середа":3,
            "Четвер":4,
            "П'ятниця":5,
            "Субота":6,
            "Неділя":7}
day_dic_reversed = {v:k for k,v in day_dic.items()}
#day_dic["П’ятниця"]=5
#day_dic_string='('+'|'.join(day_dic.keys())+')'



def find_docx_files(start_dir):

    docx_files = []
    for root, dirs, files in os.walk(start_dir):
        for file in files:
            if file.endswith('.docx') and not file.startswith('~'):
                full_path = os.path.join(root, file)
                docx_files.append(full_path)
    return docx_files

def get_octoechos_file(docx_files, echos,weekday_abr):
    for file_path in docx_files:
        #print(file_path)
        file_name = os.path.basename(file_path)
        #print(file_name)
        if file_name[0] == str(echos) and file_name[2:4]==weekday_abr:
            return file_path
    else:
        print(f"File not found for {echos}, {weekday_abr}")
        raise Exception
        #exit(-1)
        #return
    #return docx_files[0]

def format_as_code(par):
    par.runs[0].font.highlight_color=WD_COLOR_INDEX.GRAY_50

def build_sources(docx_files):
    for echos in range(1,9):
        filename = f'Джерело. Глас {echos}.docx'
        new_doc = docx.Document()
        new_doc.add_heading(f'ГЛАС {echos}', level=1)

        for day in day_name_sequence:
            new_doc.add_heading(f'Глас {echos} - {day}', level=1)
            par = new_doc.add_paragraph(f"/октоїх/{echos}/{day}")
            format_as_code(par)
            
            old_template = get_octoechos_file(docx_files,echos,day)
            old_doc = docx.Document(old_template)

            new_doc.add_heading('ВЕЧІРНЯ', level=2)
            
            par = new_doc.add_heading(f"//вечірня/господи_взиваю", level=3)
            format_as_code(par)
            
            #[Вечірня] Стихири на Господи Взиваю
            paragraphs = get_stichera_gv(old_doc)
            for i,p in enumerate(paragraphs):
                par = new_doc.add_paragraph(f"//вечірня/господи_взиваю/стихири/{i+1}")
                format_as_code(par)
                du.copy_paragraph(new_doc,p)
            #ps.copy_paragraph_list(new_doc,paragraphs)

            #[Вечірня] Стихири на стиховні
            paragraphs = get_vespers_stichera_stkh(old_doc)
            for i,p in enumerate(paragraphs):
                par = new_doc.add_paragraph(f"//вечірня/стиховня/стихири/{i+1}")
                format_as_code(par)
                du.copy_paragraph(new_doc,p)
                
            new_doc.add_heading('УТРЕНЯ', level=2)
            
            #[Утреня] Сідальні після Першого славослов'я
            paragraphs = get_first_sessional_hymns(old_doc)
            for i,p in enumerate(paragraphs):
                par = new_doc.add_paragraph(f"//утреня/сідальні_по_першому_славословї/{i+1}")
                format_as_code(par)
                du.copy_paragraph(new_doc,p)

            #[Утреня] Сідальні після Другого славослов'я
            paragraphs = get_second_sessional_hymns(old_doc)
            for i,p in enumerate(paragraphs):
                par = new_doc.add_paragraph(f"//утреня/сідальні_по_другому_славословї/{i+1}")
                format_as_code(par)
                du.copy_paragraph(new_doc,p)

            #[Утреня] Світильні
            paragraphs = get_exapostolaria(old_doc)
            for i,p in enumerate(paragraphs):
                par = new_doc.add_paragraph(f"//утреня/світильні/{i+1}")
                format_as_code(par)
                du.copy_paragraph(new_doc,p)

            #[Утреня] Стихири хвалитні
            paragraphs = get_stichera_lauds(old_doc)
            if paragraphs:
                for i,p in enumerate(paragraphs):
                    par = new_doc.add_paragraph(f"//утреня/хвалитні/стихири/{i+1}")
                    format_as_code(par)
                    du.copy_paragraph(new_doc,p)
            elif day=="НД":
                print(f"Утреня, стихири на хвалитних не знайдено для {echos}-{day}")
              
   
            #[Утреня] Стихири на стиховні
            paragraphs = get_matins_stichera_stkh(old_doc)
            if paragraphs:
                for i,p in enumerate(paragraphs):
                    par = new_doc.add_paragraph(f"//утреня/стиховня/стихири/{i+1}")
                    format_as_code(par)
                    du.copy_paragraph(new_doc,p)
            #else:
            #    print(f"Утреня, стихири на стиховні не знайдено для {echos}-{day}")
            
        new_doc.save(f"Джерело. Глас {echos}.docx")

def process_template(filepath):
    doc = docx.Document(filepath)


def get_stichera_gv(doc):

    found_header = False
    paragraphs_found = []

    for p in doc.paragraphs:

        # Шукаємо фрагмент документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "Стихири":
            found_header = True
            continue

        # Задаємо строки, які ігноруємо
        if found_header and (p.text.startswith("Стих:") or p.text.startswith("Псалом")):
            continue

        # Обробляємо потрібні параграфи
        if found_header and p.text:
            paragraphs_found.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_header and (not p.text or p.text in ("Вхід","Світло тихе")):
            return paragraphs_found

def get_vespers_stichera_stkh(doc):

    found_header = False
    paragraphs_found = []

    for p in doc.paragraphs:

        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "Стихири на стиховні":
            found_header = True
            continue

        # Задаємо строки, які ігноруємо
        if found_header and (p.text.startswith("Стих:")):
            continue

        # Обробляємо потрібні параграфи
        if found_header and p.text:
            paragraphs_found.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_header and (not p.text or p.text in ("Пісня Симеона")):
            return paragraphs_found
        
def get_first_sessional_hymns(doc):

    found_header = False
    paragraphs_found = []

    for p in doc.paragraphs:

        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "Перший сідальний":
            found_header = True
            continue

        # Обробляємо потрібні параграфи
        if found_header and p.text:
            paragraphs_found.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_header and (not p.text or p.text in ("Мала єктенія") or p.text.startswith("Відтак")):
            return paragraphs_found

def get_second_sessional_hymns(doc):

    found_header = False
    paragraphs_found = []

    for p in doc.paragraphs:

        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "Другий сідальний":
            found_header = True
            continue

        # Обробляємо потрібні параграфи
        if found_header and p.text:
            paragraphs_found.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_header and (not p.text or p.text in ("Псалом 50")):
            return paragraphs_found

def get_exapostolaria(doc):

    found_header = False
    paragraphs_found = []

    for p in doc.paragraphs:

        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "Світильний":
            found_header = True
            continue

        # Обробляємо потрібні параграфи
        if found_header and p.text and not p.text.startswith("ВСТАВИТИ"):
            paragraphs_found.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_header and (not p.text or p.text in ("Хвалитні")):
            return paragraphs_found

def get_stichera_lauds(doc):

    found_header = service_matins = False
    paragraphs_found = []

    for p in doc.paragraphs:

        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "УТРЕНЯ":
            service_matins = True
            continue
        
        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if service_matins and p.text == "Стихири":
            found_header = True
            continue

        # Задаємо строки, які ігноруємо
        if found_header and (p.text.startswith("Стих:") or p.text.startswith("Псалом")):
            continue

        # Обробляємо потрібні параграфи
        if found_header and p.text:
            paragraphs_found.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_header and (not p.text or p.text in ("Стихира євангельська")):
            return paragraphs_found

def get_matins_stichera_stkh(doc):

    found_header = service_matins = False
    paragraphs_found = []

    for p in doc.paragraphs:

        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "УТРЕНЯ":
            service_matins = True
            continue
        
        # Шукаємо початок фрагменту документу, що нас цікавить. Наприклад - за заголовком
        if service_matins and p.text == "Стихири на стиховні":
            found_header = True
            continue

        # Обробляємо потрібні параграфи
        if found_header and p.text:
            paragraphs_found.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_header and (not p.text or p.text.startswith("Після стиховні")):
            return paragraphs_found



if __name__ == "__main__":
    target_dir = eg.diropenbox(title="Оберіть початкову директорію")
    
    docx_files = find_docx_files(target_dir)

    build_sources(docx_files)

    print("Done!")

    
    '''
    doc = docx.Document(get_octoechos_file(docx_files,1,"ВВ"))

    paragraphs_found = get_gv_stichera(doc)

    print(f"Found {len(paragraphs_found)} paragraphs with stichera:")
    for i, p in enumerate(paragraphs_found):
        print(f"[{i}]: {p.text}")
    '''
