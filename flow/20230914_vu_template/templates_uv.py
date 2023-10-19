import re,glob,calendar,docx,os,easygui, shutil
from datetime import datetime

#filenames= glob.glob('*/*/*.txt')
mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])

old_files = glob.glob('2022/*/*.doc*')
filenames={}


month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1

day_dic={"ПН":1,
        "ВТ":2,
        "СР":3,
        "ЧТ":4,
        "ПТ":5,
        "СБ":6,
        "НД":7}
day_dic_reversed = {v:k for k,v in day_dic.items()}
day_dic_string='('+'|'.join(day_dic.keys())+')'

i=0
for f in old_files:
    re_result=re.search('2022\\\\2022.(\d{2})\\\\(\d{1,2})(?:-|_)(?:.*?)(?:-|_)(.*).docx?',f)
    if re_result:
        i+=1
        #print(i,re_result.group(1),re_result.group(2),re_result.group(3))
        if not int(re_result.group(1)) in filenames.keys():
            filenames[int(re_result.group(1))]={}
        if not int(re_result.group(2)) in filenames[int(re_result.group(1))].keys():
            filenames[int(re_result.group(1))][int(re_result.group(2))]={}
        filenames[int(re_result.group(1))][int(re_result.group(2))]=re_result.group(3)


for month, month_data in filenames.items():
    day_count=0
    for day in month_data.keys():
        day_count+=1
    print(month, day_count, calendar.monthrange(2023, month)[1])


for month in range(1,13):
    for day_no in range(1,calendar.monthrange(2023, month)[1]+1):
        if not day_no in filenames[month].keys():
            print("Пропущено:",month, day_no)

for m in range(1,13):
    for d, data in filenames[m].items():
        if re.search('Гл', data):
            print(m,d,data)
            filenames[m][d]=re.sub('Гл.\s?\d\s?-\s?','',data)
            print(m,d,filenames[m][d])

def get_echos(date,mode):
    if mode=='u':
        return (int(date.strftime("%U"))-25) % 8 + 1
    elif mode=='g':
        return (int(date.strftime("%U"))-24) % 8 + 1

def copy_paragraph(target_doc,source_paragraph):
    target_paragraph = target_doc.add_paragraph()
    target_paragraph.style = source_paragraph.style
    target_paragraph.alignment = source_paragraph.alignment
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

def get_octoechos_template_files():
    template_dic = {}
    all_files = glob.glob('01-Октоїх/*/*.docx')
    for f in all_files:
        re_result=re.search('01-Октоїх\\\\Глас_(\d)\\\\(\d)-(\w{2})-?.docx',f)
        if re_result:
            #print(re_result.group(2), re_result.group(3))
            doc = docx.Document(f)
            if int(re_result.group(2)) not in template_dic.keys():
                template_dic[int(re_result.group(2))]={}
            if re_result.group(3) in day_dic.keys():
                template_dic[int(re_result.group(2))][day_dic[re_result.group(3)]]=f
            else:
                print("Something wrong with", f, re_result.group(3))
        else:
            print("None found for:", f)
    return template_dic


def get_octoechos_templates():
    template_dic = {}
    all_files = glob.glob('01-Октоїх/*/*.docx')
    for f in all_files:
        re_result=re.search('01-Октоїх\\\\Глас_(\d)\\\\(\d)-(\w{2})-?.docx',f)
        if re_result:
            #print(re_result.group(2), re_result.group(3))
            doc = docx.Document(f)
            if int(re_result.group(2)) not in template_dic.keys():
                template_dic[int(re_result.group(2))]={}
            if re_result.group(3) in day_dic.keys():
                template_dic[int(re_result.group(2))][day_dic[re_result.group(3)]]=doc.paragraphs
            else:
                print("Something wrong with", f, re_result.group(3))
        else:
            print("None found for:", f)
    return template_dic

def get_menaion_template_files():
    filenames = glob.glob(f'В,У - Мінея/{month_no:02}*/*.docx')
    template_dic={}
    for f in filenames:
        print("Checking",f)
        pattern='(\d{2})-__-(.*?).docx'
        print(pattern)
        re_result = re.search(pattern,f)
        doc = docx.Document(f)
        template_dic[int(re_result.group(1))]=f
    
    return template_dic

def get_menaion_templates():
    filenames = glob.glob(f'В,У - Мінея/{month_no:02}*/*.docx')
    template_dic={}
    for f in filenames:
        print("Checking",f)
        pattern='(\d{2})-__-(.*?).docx'
        print(pattern)
        re_result = re.search(pattern,f)
        doc = docx.Document(f)
        template_dic[int(re_result.group(1))]=doc.paragraphs
    
    return template_dic

templates_octoechos_dic = get_octoechos_template_files()
templates_menaion_dic = get_menaion_template_files()

folder_name=f'drafts\\{year_no}-{month_no}-{mode}'
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

print(year_no, month_no)
for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
    print(d,datetime(year_no, month_no, d).weekday()+1)
    dest_filename=f'{folder_name}\\{d:02}-{day_dic_reversed[datetime(year_no, month_no, d).weekday()+1]}'
    if datetime(year_no, month_no, d).weekday()+1==7:
        #echos=
        dest_filename+=f'-Гл.{get_echos(datetime(year_no,month_no,d),mode)}'
    dest_filename+=f'-{filenames[month_no][d]}.docx'
    if d in templates_menaion_dic.keys() and datetime(year_no, month_no, d).weekday()+1!=7:
        src_filename=templates_menaion_dic[d]
    else:        
        src_filename=templates_octoechos_dic[get_echos(datetime(year_no,month_no,d),mode)][datetime(year_no, month_no, d).weekday()+1]
    shutil.copy(src_filename,dest_filename)        

