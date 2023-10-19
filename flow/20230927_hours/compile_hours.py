import re, docx, csv, easygui
from datetime import datetime

#mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
mode = 'u'
month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1


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
            print(row)
            print(row[0],str(month_no))
            if row[0]==str(month_no):
                matrix[int(row[1].split('.')[2])]=row[2:]
    return matrix



def get_resurrection_template_texts(path):
    template_dic={}
    doc = docx.Document(path)
    #rubric_found = cur_glas = None
    #i=0
    for p in doc.paragraphs:
        re_result = re.search("Глас (\d)")
        if re_result:
            echos=int(re_result.group(1))
        
        if p
        





        
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


            
#import troparia - resurrection
#import troparia - mineion



#for all days in month, choose template
#for each doc
#   insert troparia
#   insert kondakion

ordo_matrix = get_matrix("ЧасиЮл.csv")
templates_resurrection = get_resurrection_template_texts('воскресні.docx')


