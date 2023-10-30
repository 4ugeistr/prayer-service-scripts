import re, docx, csv, easygui, glob, os, calendar, shutil
from datetime import datetime

mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
#mode = 'u'
month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


def get_echos(date,mode):
    if mode=='u':
        return (int(date.strftime("%U"))-25) % 8 + 1
    elif mode=='g':
        return (int(date.strftime("%U"))-24) % 8 + 1


def get_matrix(csv_filename):
    matrix={}
    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            #print(row)
            #print(row[0],str(month_no))
            if row[0]==str(month_no):
                matrix[int(row[1].split('.')[2])]=row[2:]
    return matrix


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

def copy_paragraph_before(paragraph,source_paragraph):
    target_paragraph = paragraph.insert_paragraph_before()
    for run in source_paragraph.runs:
        new_run = target_paragraph.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        new_run.font.size = run.font.size
        new_run.font.name = run.font.name
        new_run.font.color.rgb = run.font.color.rgb
        new_run.font.highlight_color = run.font.highlight_color


def get_resurrection_template_texts(path):
    template_dic={}
    doc = docx.Document(path)
    #rubric_found = cur_glas = None
    #i=0
    key = None
    for p in doc.paragraphs:
        re_result = re.search("Глас (\d)", p.text)
        if re_result:
            key=int(re_result.group(1))
            if not key in template_dic:
                template_dic[key]=[]
        if key and p.text.find("Тропар")!=-1:
            template_dic[key].append(p)
        if key and p.text.find("Кондак")!=-1:
            template_dic[key].append(p)
    return template_dic

def get_menaion_template_texts(path):
    template_dic={}
    doc = docx.Document(path)
    key = None
    for p in doc.paragraphs:
        re_result = re.search("^(\d{1,2})",p.text)
        if re_result:
            key = int(re_result.group(1))
            if not key in template_dic:
                template_dic[key]=[]
        if key and p.text.find("Тропар")!=-1:
            template_dic[key].append(p)
        if key and p.text.find("Кондак")!=-1:
            template_dic[key].append(p)
    return template_dic

def get_feast_template_texts(path):
    template_dic={}
    doc = docx.Document(path)
    key = None
    for p in doc.paragraphs:
        re_result = re.search("^(\d{1,2})",p.text)
        if re_result:
            key = int(re_result.group(1))
            if not key in template_dic:
                template_dic[key]=[]
        if key and p.text.find("Тропар")!=-1:
            template_dic[key].append(p)
        if key and p.text.find("Кондак")!=-1:
            template_dic[key].append(p)
    return template_dic


def get_template_files(path):
    files = glob.glob(path)
    print(files)
    template_dic={}
    for f in files:
        re_result = re.search(r"docx_templates\\hours-template-(\d)",f)
        template_dic[int(re_result.group(1))] = f
    return template_dic

#import troparia - resurrection
#import troparia - mineion



#for all days in month, choose template
#for each doc
#   insert troparia
#   insert kondakion

ordo_matrix = get_matrix("Часи.csv")
templates_resurrection = get_resurrection_template_texts('воскресні.docx')
templates_menaion = get_menaion_template_texts('тропарі-11.docx')
templates_feast = get_feast_template_texts('свято-11.docx')
template_file_list = get_template_files('docx_templates/hours-template-*.docx')


folder_name=f'drafts\\{year_no}-{month_no}-{mode}'
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

    
print(year_no, month_no)
for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
    #print(d,datetime(year_no, month_no, d).weekday()+1)
    dest_filename=f'{folder_name}\\{d:02}.{month_no}.docx'

    if ordo_matrix[d][1]=='y' or datetime(year_no, month_no, d).weekday()+1==7:
        shutil.copy(template_file_list[7],dest_filename)
    else:
        shutil.copy(template_file_list[datetime(year_no, month_no, d).weekday()+1],dest_filename)
print("templates created")


hours_matrix={}
for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
    if ordo_matrix[d][3] =="":
        if len(templates_menaion[d])<=2:
            hours_matrix[d] = ["","","",
                            "",templates_menaion[d][0],templates_menaion[d][1],
                            "","","",
                            "",templates_menaion[d][0],templates_menaion[d][1]]
        else:
            #print(d)
            hours_matrix[d] = ["","","",
                            "",templates_menaion[d][0],templates_menaion[d][1],
                            "","","",
                            "",templates_menaion[d][2],templates_menaion[d][3]]
    else:
        hours_matrix[d]=[]
        i=0
        for i in range(12):
            o = ordo_matrix[d][2:][i]
            if o=="":
                hours_matrix[d].append("")
            elif o =='resurrection' and (i+1)%3!=0:
                hours_matrix[d].append(templates_resurrection[get_echos(datetime(year_no,month_no,d),mode)][0])
            elif o =='resurrection' and (i+1)%3==0:
                hours_matrix[d].append(templates_resurrection[get_echos(datetime(year_no,month_no,d),mode)][1])
            elif o =='feast' and (i+1)%3!=0:
                hours_matrix[d].append(templates_feast[d][0])
            elif o =='feast' and (i+1)%3==0:
                hours_matrix[d].append(templates_feast[d][1])
            
            elif o =='saint' and (i+1)%3!=0:
                hours_matrix[d].append(templates_menaion[d][0])
            elif o =='saint' and (i+1)%3==0:
                hours_matrix[d].append(templates_menaion[d][1])    

print("paragraphs gathered")          
'''
for p in hours_matrix[2]:
	if type(p)==str:
		print("")
	else:
		print(p.text[:40])
'''


hours_dic = {1:"ПЕРШИЙ",
             3:"ТРЕТІЙ",
             6:"ШОСТИЙ",
            9:"ДЕВ'ЯТИЙ"}
hours_dic_reversed = {v:k for k,v in hours_dic.items()}
hours_dic_reversed_string='('+'|'.join(hours_dic_reversed.keys())+')'

hours_translate = {1:1,
                   3:2,
                   6:3,
                   9:4}


for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
    #print(d)
#for d in range(1,2):
    dest_filename=f'{folder_name}\\{d:02}.{month_no}.docx'
    doc = docx.Document(dest_filename)
    hour=None
    for p in doc.paragraphs:
        
        re_result = re.search(f"^ЧАС {hours_dic_reversed_string}",p.text)
        if re_result:
            
            hour = int(hours_dic_reversed[re_result.group(1)])
            #print(d, "found hour", hour)
        re_result = re.search("^Слава Отцю, і Сину, і Святому Духові\.",p.text)
        if re_result and hour:
            #print("found СН")
            value=hours_matrix[d][3*hours_translate[hour]-3]
            if not type(value)==str:
                copy_paragraph_before(p,value)

        re_result = re.search("^Тропар(\s)?$",p.text)
        if re_result and hour:
            #print("found тропар")
            value=hours_matrix[d][3*hours_translate[hour]-2]
            copy_paragraph_before(p,value)
            delete_paragraph(p)

        re_result = re.search("^Кондак(\s)?$",p.text)
        if re_result and hour:
            #print("found кондак")
            value=hours_matrix[d][3*hours_translate[hour]-1]
            copy_paragraph_before(p,value)
            delete_paragraph(p)
    doc.save(dest_filename)
        
for k,v in hours_matrix.items():
    print(k,len(v))
