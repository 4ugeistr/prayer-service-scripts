import re, docx, glob
from datetime import datetime

#cur_year
#cur_month


month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1


#c:\Work\GitHub\prayer-service-scripts\flow\20230914_vu_template\Стихири - Мінея\Мінея_06_Червень.docx
files = glob.glob(f'Стихири - Мінея\\Мінея_{month_no:02}*.docx')
#docx_filename = f'Мінея_{}_Лютий.docx'
#doc = docx.Document(files[0])


gospel_books = {
    "Мт":"Матея",
    "Мр":"Марка",
    "Лк":"Луки",
    "Ів":"Івана",
}



def get_stichera_matrix(path):
    stichera_dic={}
    doc = docx.Document(path)

    day=None
    handle = None
    #gv = None
    #st_found = None
    for p in doc.paragraphs:
        #if day==26:
        #    print(p.text[:40])
        re_result = re.search(r"^(\d{1,2}) (.*)",p.text)
        if re_result:
            day = int(re_result.group(1))
            stichera_dic[day] = {"label":re_result.group(2)}
            #do we need a flag?
            handle='gv'
            #st_found=False
            #stichera_dic[day]["gv_stichera"]=[]
            continue

        re_result = re.search(r"богородичний",p.text)
        if re_result:
            #print("   found theo")
            stichera_dic[day].setdefault(f"{handle}_theotokion", []).append(p)
            continue
        
        re_result = re.search(r"^Слава",p.text)
        if re_result:
            #print("   found doxa")
            stichera_dic[day].setdefault(f"{handle}_doxa", []).append(p)
            continue
        

        if re.search("стиховні", p.text.lower()):
            handle = "st"
            continue

        if handle in ('gv','st') and p.text:
            #print("   found stichera",handle)
            #if not f"{handle}_stichera" in stichera_dic[day]:
            if 'тихири' in p.text:
                continue
            stichera_dic[day].setdefault(f"{handle}_stichera", []).append(p)
        
            
    return stichera_dic

def get_generic_stichera_matrix(path):
    stichera_dic={}
    doc = docx.Document(path)
    #day=None
    handle = None
    #gv = None
    #st_found = None
    for p in doc.paragraphs:
        #if day==26:
        #    print(p.text[:40])
        re_result = re.search(r"Служба (.*)",p.text)
        if re_result:
            saint_rank = re_result.group(1)
            stichera_dic[saint_rank] = {"label":p.text}
            #do we need a flag?
            handle=None
            #st_dound=False
            #stichera_dic[day]["gv_stichera"]=[]
            continue

        if re.search("Господи, взиваю",p.text):
            handle='gv'
            continue


        re_result = re.search(r"богородичний",p.text)
        if re_result:
            #print("   found theo")
            stichera_dic[saint_rank].setdefault(f"{handle}_theotokion", []).append(p)
            continue
        
        re_result = re.search(r"^Слава",p.text)
        if re_result:
            #print("   found doxa")
            stichera_dic[saint_rank].setdefault(f"{handle}_doxa", []).append(p)
            continue
        

        if re.search("стиховні", p.text.lower()):
            handle = "st"
            continue

        if handle in ('gv','st') and p.text:
            #print("   found stichera",handle)
            if 'тихири' in p.text:
                continue
            stichera_dic[saint_rank].setdefault(f"{handle}_stichera", []).append(p)
        
            
    return stichera_dic

def get_resurrection_gospel_matrix(path):
    dic={k:{} for k in range(1,12) }
    doc = docx.Document(path)
    gospel_no = None


    # flag = None | gospel | exapostolarion | stichera
    flag = None
    for p in doc.paragraphs:
        
        re_result = re.search(r"(\d{1,2})-\w{2} Євангеліє",p.text)
        if re_result:
            #print("Found new gospel_no in:", p.text)
            flag = None
            gospel_no = int(re_result.group(1))
            continue
        
        
        re_result = re.search(r"^(Євангеліє: |Світильний|Стихира євангельська)(.*)",p.text)
        if re_result:

            match re_result.group(1):
                case "Євангеліє: ":
                    flag = 'gospel'
                    #dic[gospel_no]['zachalo']=re_result.group(2)
                    zachalo=re_result.group(2)
                    dic[gospel_no]['gospel_apostle']=gospel_books[zachalo[:2]]
                case "Світильний":
                    flag = 'exapostolarion'
                case "Стихира євангельська":
                    flag = 'stichera'       

            dic[gospel_no][flag] = []
            continue

        if flag:
            dic[gospel_no][flag].append(p)
    return dic

        
if __name__ == "__main__":
    menaion_stichera_matrix = get_stichera_matrix(files[0])
    generic_stichera_matrix = get_generic_stichera_matrix("Стихири ГВ загальної служби.docx")
    resurrection_gospel_matrix = get_resurrection_gospel_matrix("11-ВОСКРЕСНІ ЄВАНГЕЛІЯ.docx")
    print("DONE.")





