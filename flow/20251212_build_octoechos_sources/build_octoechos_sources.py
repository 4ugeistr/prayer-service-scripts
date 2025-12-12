import ps_docx_utils as du
import os, docx
import easygui as eg

from sys import exit

day_name_sequence = ["НД","ПН","ВТ","СР","ЧТ","ПТ","СБ",]

day_short_dic={"ПН":1,
        "ВТ":2,
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

def build_sources(docx_files):
    for echos in range(1,9):
        filename = f'Джерело. Глас {echos}.docx'
        new_doc = docx.Document()
        new_doc.add_heading(f'ГЛАС {echos}', level=1)

        for day in day_name_sequence:
            new_doc.add_heading(f'Глас {echos} - {day}', level=1)
            new_doc.add_paragraph(f"/октоїх/1/{day}")
            new_doc.add_heading('ВЕЧІРНЯ', level=2)
            new_doc.add_heading(f"//господи_взиваю", level=3)
            old_template = get_octoechos_file(docx_files,echos,day)
            old_doc = docx.Document(old_template)
            paragraphs = get_gv_stichera(old_doc)

            for i,p in enumerate(paragraphs):
                new_doc.add_paragraph(f"//господи_взиваю/{i+1}")
                du.copy_paragraph(new_doc,p)
            #ps.copy_paragraph_list(new_doc,paragraphs)


        new_doc.save(f"Джерело. Глас {echos}.docx")

def process_template(filepath):
    doc = docx.Document(filepath)


def get_gv_stichera(doc):

    found_stichara_header = False
    stichera_paragraphs = []

    for p in doc.paragraphs:

        # Шукаємо фрагмент документу, що нас цікавить. Наприклад - за заголовком
        if p.text == "Стихири":
            found_stichara_header = True
            continue

        # Задаємо строки, які ігноруємо
        if found_stichara_header and (p.text.startswith("Стих:") or p.text.startswith("Псалом")):
            continue

        # Обробляємо потрібні параграфи
        if found_stichara_header and p.text:
            stichera_paragraphs.append(p)
            continue

        # Виходимо, коли закінчили опрацьовувати потрібний фрагмент
        if found_stichara_header and (not p.text or p.text in ("Вхід","Світло тихе")):
            return stichera_paragraphs

if __name__ == "__main__":
    target_dir = eg.diropenbox(title="Оберіть початкову директорію")
    docx_files = find_docx_files(target_dir)

    build_sources(docx_files)





    doc = docx.Document(get_octoechos_file(docx_files,1,"ВТ"))

    stichera_paragraphs = get_gv_stichera(doc)

    print(f"Found {len(stichera_paragraphs)} paragraphs with stichera:")
    for i, p in enumerate(stichera_paragraphs):
        print(f"[{i}]: {p.text}")