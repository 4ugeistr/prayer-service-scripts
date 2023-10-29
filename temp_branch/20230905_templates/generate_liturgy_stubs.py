import easygui, calendar, os, docx, re
from datetime import datetime, timedelta

mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])


#select month
#month_no = easygui.enterbox("Введіть номер місяця", "Місяць")
month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1
#select mode
#mode = easygui.choicebox('u - Новоюліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])

month_dic = {'Січень':1,
              'Лютий':2,
              'Березень':3,
              'Квітень':4,
              'Травень':5,
              'Червень':6,
              'Липень':7,
              'Серпень':8,
              'Вересень':9,
              'Жовтень':10,
              'Листопад':11,
              'Грудень':12}
month_dic_reversed = {v:k for k,v in month_dic.items()}
month_dic_string='('+'|'.join(month_dic.keys())+')'

day_dic = {"Понеділок":1,
            "Вівторок":2,
            "Середа":3,
            "Четвер":4,
            "П'ятниця":5,
            "Субота":6,
            "Неділя":7}
day_dic_reversed = {v:k for k,v in day_dic.items()}
day_dic_string='('+'|'.join(day_dic.keys())+')'

doc = docx.Document("Літургія - Мінея\\10-Літургія-шаблони.docx") 

def get_echos(date,mode):
    if mode=='u':
        return (int(date.strftime("%U"))-25) % 8 + 1
    elif mode=='g':
        return (int(date.strftime("%U"))-24) % 8 + 1

def copy_paragraph(target_doc,source_paragraph):
    target_paragraph = target_doc.add_paragraph()
    for run in source_paragraph.runs:
        new_run = target_paragraph.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        new_run.font.size = run.font.size
        new_run.font.name = run.font.name
        new_run.font.color.rgb = run.font.color.rgb
        new_run.font.highlight_color = run.font.highlight_color

def copy_paragraph_list(target_doc,paragraph_list):
    for p in paragraph_list:
        copy_paragraph(target_doc,p)
        
path = "Літургія - змінні частини - шаблони.docx"
def get_resurrection_template_texts(path):
    template_dic={}
    doc = docx.Document(path)
    rubric_found = cur_glas = None
    i=0
    for p in doc.paragraphs:
        #print(p.text)
        if not rubric_found and not p.text.startswith("Воскресна служба"):
            pass
        elif p.text.startswith("Воскресна служба"):
            rubric_found=True
        if rubric_found:
            re_result = re.search("Неділя – глас (\d)",p.text)
            if re_result:
                cur_glas = int(re_result.group(1))
                template_dic[cur_glas]=[]
                continue

        if cur_glas and p.text.startswith("</vidpust>"):
            template_dic[cur_glas].append(p)
            cur_glas=None

        if cur_glas:
            template_dic[cur_glas].append(p)
    return template_dic
            
def get_everyday_template_texts(path):
    template_dic={}
    doc = docx.Document(path)
    rubric_found = cur_glas = None
    i=0
    for p in doc.paragraphs:
        #print(p.text)
        if not rubric_found and not p.text.startswith("Повсякденна служба"):
            pass
        elif p.text.startswith("Повсякденна служба"):
            rubric_found=True
        if rubric_found:
            re_result = re.search(f"{day_dic_string}",p.text)
            if re_result:
                cur_glas = day_dic[re_result.group(1)]
                template_dic[cur_glas]=[]
                continue

        if cur_glas and p.text.startswith("</vidpust>"):
            template_dic[cur_glas].append(p)
            cur_glas=None

        if cur_glas:
            template_dic[cur_glas].append(p)
    return template_dic


           
        
def get_menaion_template_texts(doc):
    directory_path ="Літургія - Мінея"
    all_files = os.listdir(directory_path)
    docx_file = [file for file in all_files if file.endswith(".docx") and file[:2]==f'{month_no:02}' and len(file) >= 4][0]
    print(docx_file)
    
    template_found=False
    template_dic={}
    for p in doc.paragraphs:
        re_result=re.search(f'^{month_dic_string} (\d+)',p.text)
        if re_result:
            cur_date = int(re_result.group(2))
            template_dic[cur_date]=[]
            template_found = True
            
        #re_result=re.search(f'<ustav',p.text)
        #if re_result:
        #    template_found=True
        if template_found:
            template_dic[cur_date].append(p)

        re_result=re.search(f'</vidpust',p.text)
        if re_result:
            template_dic[cur_date]=template_dic[cur_date][1:]
            template_found=False
            cur_date=None
            
    
    #print(template_dic)
    return template_dic

templates_menaion = get_menaion_template_texts(doc)
templates_resurrection = get_resurrection_template_texts(path)
templates_everyday = get_everyday_template_texts(path)

new_doc = docx.Document()
'''
Get all days in the month
'''
for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
    print(d)
    heading_text =' '.join([month_dic_reversed[month_no],str(d)+',',day_dic_reversed[datetime(year_no, month_no, d).weekday()+1]])
    new_doc.add_heading(heading_text, level=2)
    if datetime(year_no, month_no, d).weekday()+1==7:
        #print(d, "copied sunday ", get_echos(datetime(year_no,month_no,d)))
        copy_paragraph_list(new_doc,templates_resurrection[get_echos(datetime(year_no,month_no,d),mode)])
    elif d in templates_menaion.keys():
        copy_paragraph_list(new_doc,templates_menaion[d])
    else:
        copy_paragraph_list(new_doc,templates_everyday[datetime(2023, month_no, d).weekday()+1])
        
    
mode_new = 'ГР' if mode=='g' else "НЮ"
new_doc.save(f'{month_no}-Літургія-{mode_new}.docx')
