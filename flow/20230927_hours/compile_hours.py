import re, docx, csv, easygui
from datetime import datetime

#mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
mode = 'u'
month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1

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



def get_resurrection_troparia_texts(path):
    template_dic={}
    doc = docx.Document(path)
    echos = None
    #i=0
    for p in doc.paragraphs:
        re_result = re.search("Глас (\d)",p.text)
        if re_result:
            #print("found")
            echos=int(re_result.group(1))
            template_dic[echos]=[]
        #print(p.text)

        re_result = re.search("((Тропар|Кондак).*?\(г\. \d\)): (.*)",p.text)
        if echos and re_result:
            template_dic[echos].append([re_result.group(1),re_result.group(3)])
    return template_dic

def get_menaion_troparia_texts(path):
    print("in")
    template_dic={}
    doc = docx.Document(path)
    key = None #cur_day
    #i=0
    for p in doc.paragraphs:
        print(p.text[:10])
        re_result = re.search(f"^(\d+) (.*)",p.text)
        if re_result:
            print("found")
            key=int(re_result.group(1))
            cur_day_heading = re_result.group(2)
            template_dic[cur_day]=[]
            print(cur_day)
            
        #print(p.text)
        
        re_result_troparion = re.search("(Тропар.*?\(г\. \d\)): (.*)",p.text)
        re_result_kondakion = re.search("(Кондак.*?\(г\. \d\)): (.*)",p.text)
        if key:
            if re_result_troparion:
                pass
            elif re_result_kondakion:
                pass
            else:
                pass
            
        #    template_dic[echos].append([re_result.group(1),re_result.group(3)])
    return template_dic
            
#import troparia - resurrection
#import troparia - mineion



#for all days in month, choose template
#for each doc
#   insert troparia
#   insert kondakion

ordo_matrix = get_matrix("Часи.csv")
templates_resurrection = get_resurrection_troparia_texts('воскресні.docx')
templates_menaion = get_menaion_troparia_texts(f'тропарі-{month_no}.docx')


