import re, docx, glob
from datetime import datetime

#cur_year
#cur_month


month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1



files = glob.glob(f'Мінея_{month_no:02}*.docx')
#docx_filename = f'Мінея_{}_Лютий.docx'
#doc = docx.Document(files[0])

def get_stichera_matrix(path):
    stichera_dic={}
    doc = docx.Document(path)

    day=None
    handle = None
    #gv = None
    #st_found = None
    for p in doc.paragraphs:

        re_result = re.search(r"^(\d{1,2}) (.*)",p.text)
        if re_result:
            day = int(re_result.group(1))
            stichera_dic[day] = {"label":re_result.group(2)}
            #do we need a flag?
            handle='gv'
            #st_dound=False
            stichera_dic[day]["gv_stichera"]=[]
            continue
        
        re_result = re.search(r"^Слава",p.text)
        if re_result:
            if handle=='gv':
                stichera_dic[day]["gv_doxa"]=p
            elif handle=='st':
                stichera_dic[day]["st_doxa"]=p
            continue
        
        re_result = re.search(r"^І нині",p.text)
        if re_result:
            if handle=='gv':
                stichera_dic[day]["gv_theotokion"]=p
            elif handle=='st':
                stichera_dic[day]["st_doxa"]=p
            continue

        if re.search("стиховні", p.text):
            handle = "st"
            continue

        
            
        if handle == 'gv':
            stichera_dic[day]["gv_stichera"].append(p)

    return stichera_dic

        

stichera_matrix = get_stichera_matrix(files[0])

'''
PLAN:
get
'''






