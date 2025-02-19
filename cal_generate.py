import paschalia, calendar, docx, easygui, csv
from datetime import datetime

import ps_date_utils as du
from ps_docx_utils import format_line


#mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
#year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1
year_no = 2025

style_map = """
i => i
b => b
h2 => p
"""

def get_matrix_full(csv_filename):
    matrix=[]
    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            matrix.append(row)
    return matrix
'''
def insert_header_liturgy(doc,date):
    day_name = du.day_dic_reversed[date.weekday()+1]
    #month_name = month_dic_reversed[month_no]
    week_no=paschalia.get_week(date,"","")[:2]

    #Субота, Неділя, Тиждень etc...
    special_dates=[{"date":datetime(year_no,9,14),
                     "holiday":"Воздвиженні",
                     "holiday_locative":"Воздвиженні",
                     "holiday_instrumental":"Воздвиженям"},
                     {"date":datetime(year_no,12,25),
                     "holiday":"Різдво",
                     "holiday_locative":"Різдві",
                     "holiday_instrumental":"Різдвом"},
                    {"date":datetime(year_no,1,6),
                     "holiday":"Богоявленні",
                     "holiday_locative":"Богоявленні",
                     "holiday_instrumental":"Богоявленням"},
                     ]
    for sd in special_dates:
        diff = (sd["date"]-date).days
        if abs(diff)<=7 and date.weekday()+1 in [6,7]:
            if diff>0:
                p_new = doc.add_paragraph(f"{day_name} перед {sd['holiday_instrumental']}")
            else:
                p_new = doc.add_paragraph(f"{day_name} по {sd['holiday_locative']}")
            format_line(p_new, '')
    
    if date.weekday()+1 in [6,7]:
        #p_new = doc.add_paragraph(f"{day_name} {week_no} по П'ятидесятниці.")
        p_new = doc.add_paragraph(paschalia.get_day_label_legacy(date))
        if date.weekday()+1 == 7:
            p_new.text+= f" Гл. "+str(paschalia.get_echos(date,paschalia_dates))+"."
            format_line(p_new, 'r')
        else:
            format_line(p_new, '')

    if date.weekday()+1 in [1]:
        #p_new = doc.add_paragraph(f"Тиждень {week_no} по П'ятидесятниці.")
        p_new = doc.add_paragraph(paschalia.get_day_label_legacy(date))
        format_line(p_new, '')

    #перелік святих
    lst = filter(lambda l:int(l[0])==date.month and int(l[1])==date.day,day_headers_menaion[1:])
    for l in lst:
        p_new=doc.add_paragraph(l[9])
        format_line(p_new, ''.join(l[2:5]))
'''

def get_sunday_header(date,mode):

    week = paschalia.get_week_from_50(date,paschalia.get_prev_next_pascha(date,mode))
    #while True:
    ending =''
    for k,v in paschalia.ending_fem_dic.items():
        if str(week).endswith(k):
            ending = v
            break
    res = f"Неділя {week}-{ending} по Зісланні Святого Духа."
    return [{'text':f'{res}','format':'ri'}]



def get_special_day_strings(date):
    day_name = du.day_dic_reversed[date.weekday()+1]
    #month_name = month_dic_reversed[month_no]
    #week_no=paschalia.get_week(date,"","")[:2]

    #Субота, Неділя, Тиждень etc...
    special_dates=[{"date":datetime(year_no,9,14),
                     "holiday":"Воздвиженні",
                     "holiday_locative":"Воздвиженні",
                     "holiday_instrumental":"Воздвиженям"},
                     {"date":datetime(year_no,12,25),
                     "holiday":"Різдво",
                     "holiday_locative":"Різдві",
                     "holiday_instrumental":"Різдвом"},
                    {"date":datetime(year_no,1,6),
                     "holiday":"Богоявленні",
                     "holiday_locative":"Богоявленні",
                     "holiday_instrumental":"Богоявленням"},
                     ]
    res = []
    for sd in special_dates:
        diff = (sd["date"]-date).days
        if abs(diff)<=7 and abs(diff)>0 and date.weekday()+1 in [6,7]:
            if diff>0:
                txt = f"{day_name} перед {sd['holiday_instrumental']}"
            else:
                txt = f"{day_name} по {sd['holiday_locative']}"
            res.append({'text':txt,'format':'ir'})
        if res and res[-1]['text'] == 'Неділя перед Різдвом':
            res[-1]['text'] += ', святих Отців'
        if res and res[-1]['text'] == 'Неділя по Різдві':
            res[-1]['text'] += '. Пам’ять святих і праведних Йосифа Обручника, Давида, царя, і Якова, брата Божого'

    
    #Hardcode for 2025
    #Свято Матері Божої Неустанної Помочі
    if date == datetime(date.year,7,6):
        res.append({'text':"🕁 Свято Матері Божої Неустанної Помочі",'format':'ir','arbitrary_symbol':'*'})
    #Собори в липні, Жовтні
    if date == datetime(date.year,7,13):
        res.append({'text':"Неділя 5-та, святих отців шести Вселенських Соборів.",'format':'ir'})
    if date == datetime(date.year,10,12):        
        res.append({'text':"Неділя 18-та, cвятих отців Сьомого Вселенського Собору.",'format':'ir'})
    #if date == datetime(date.year,12,14):        
    #    res.append({'text':"Неділя 27-ма, святих Праотців",'format':'ir'})
        
    return res
    


def get_special_unimportant_day_strings(date):
    res = []
    #Hardcode for 2025
    if date == datetime(date.year,3,26):
        res.append({'text':"Віддання Благовіщення",'format':'ir'})
    #Hardcode for 2025
    if date == datetime(date.year,2,9):
        res.append({'text':"Віддання Стрітення",'format':'ir'})
    return res
    #іноді для Стрітення теж нестандартне віддання.


def get_triodion_strings(date,mode):
    lines = []
    paschalia_dates = paschalia.get_prev_next_pascha(date,mode)
    triodion_params = paschalia.get_day_details(date,paschalia_dates)
    
    matrix = []
    
    if triodion_params[0] in ('pentecost') and triodion_params[2]:
        matrix += list(filter(lambda l:l[0]==triodion_params[0] and l[2] and int(l[2])==triodion_params[2] and int(l[3])==triodion_params[3],day_headers_triodion))
    elif triodion_params[0] in ('lent','pascha'):
        matrix += list(filter(lambda l:l[0]==triodion_params[0] and int(l[1])==triodion_params[1] and int(l[3])==triodion_params[3],day_headers_triodion))
    else:
        matrix += list(filter(lambda l:l[0]==triodion_params[0] and l[1] and int(l[1])==triodion_params[1] and not (l[2]) and int(l[3])==triodion_params[3],day_headers_triodion))

    for l in matrix:
        lines.append({"text":l[10], "format":''.join(l[4:7]),'arbitrary_symbol':l[8]})
        #p_new=doc.add_paragraph(l[9])
        #format_line(p_new, ''.join(l[2:5]))
    return lines


def get_menaion_strings(date,short = False):
    fetch_index = 8 if short else 9
    lines = []
    matrix = filter(lambda l:int(l[0])==date.month and int(l[1])==date.day,day_headers_menaion[1:])
    for l in matrix:
        lines.append({"text":l[fetch_index], "format":''.join(l[2:5]),'arbitrary_symbol':l[6]})
        #p_new=doc.add_paragraph(l[9])
        #format_line(p_new, ''.join(l[2:5]))
    return lines

def convert_db_entries_to_paragraphs(doc, lst, mode = ''):
    for line in lst:
        #print(line["text"])
        p=doc.add_paragraph(line["text"])
        format_line(p,line["format"],mode)        

def compile_header(date,mode,short=True):
    lst = []

    strings = get_special_day_strings(date)
    if strings:
        lst += strings

    strings = get_triodion_strings(date,mode)
    if strings:
        lst += strings
    elif not lst and date.weekday()+1==7:
        lst += get_sunday_header(date,mode) 
    
    lst += get_menaion_strings(date,short)
    lst += get_special_unimportant_day_strings(date)

    #lst[0]['text'] = f'{date.day} '+lst[0]['text']
    return lst


def format_line_for_html(s, handle='',symbol=''):
    #handle = "bir"
    if symbol:
        s = symbol+s[2:]
        if symbol == '#':
            s=f"<strong>{s}</strong>"
            return s
    
    #додаткові строки до дванадесятого свята
    if 'r' in handle and 'b' in handle:
        s=f"<strong>{s}</strong>"
        return s

    #i - чорний курсив
    if not 'r' in handle and 'i' in handle and  not 'b' in handle:
        s=f"<em>{s}</em>"
        return s
    #r - червоний звичайний
    if 'r' in handle and not 'i' in handle and not 'b' in handle:
        s=f"<span>{s}</span>"
        return s
    #ri - червоний курсив
    if 'r' in handle and 'i' in handle and not 'b' in handle:
        s=f"<i>{s}</i>"
        return s
    
    

    if 'b' in handle:
        s=f"<b>{s}</b>"
    if 'r' in handle:
        s=f"<i>{s}</i>"
    return s


def prepare_header_for_docx(date,mode):
    lst = compile_header(date,mode)
    paragraph_list=[]
    doc=docx.Document()
    for item in lst:
        new_p = doc.add_paragraph(item['text'])
        format_line(new_p,handle=item['format'])
        paragraph_list.append(new_p)
    return paragraph_list

def prepare_header_for_html(date,mode):
    lst = compile_header(date,mode)
    res=""
    lst_formatted = []
    for item in lst:
        lst_formatted.append(format_line_for_html(item['text'],handle = item['format'],symbol=item.get('arbitrary_symbol')))
    
    res = '<br>'.join(lst_formatted)
    return res


'''
#поганий підхід

def convert_db_entries_to_html(lst):
    doc = docx.Document()
    convert_db_entries_to_paragraphs(doc, lst,mode='html')
    doc.save('tmp.docx')

    with open('tmp.docx', "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file, style_map=style_map).value
    result = result.replace("</p><p>","<br>")
    result = result.replace("<p>","")
    result = result.replace("</p>","")
    return result
'''

day_headers_menaion = get_matrix_full("matrices\\Місяцеслов-БД.csv")
day_headers_triodion = get_matrix_full("matrices\\Місяцеслов-БД-Тріодь.csv")


if __name__ == "__main__":

    #day_headers_menaion = get_matrix_full("Місяцеслов-БД.csv")
    #day_headers_triodion = get_matrix_full("Дні-Тріодь.csv")

    mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
    paschalia_dates = paschalia.get_prev_next_pascha(datetime.now(), mode)

    doc_filename = f"{year_no}_Календар_{mode}.docx"

    print(paschalia.get_day_details(datetime(2024,3,3),paschalia_dates))

    doc = docx.Document()
    doc.add_heading(f"{year_no}",0)
    for month_no in range(1,13):
        doc.add_heading(du.month_dic_reversed[month_no],level=1)
        
        for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
            #doc.add_heading(f"{d} "+du.day_dic_reversed[datetime(year_no,month_no,d).weekday()+1],level=2)

            doc.add_paragraph(str(du.day_dic_reversed[datetime(year_no,month_no,d).weekday()+1]))
            
            
           #doc.add_paragraph(f"{d}")
            
            #GET MENAION DAY HEADING
            first_line = f"{d} "
            for line in get_menaion_strings(datetime(year_no,month_no,d)):
                text = first_line+ line["text"]
                p=doc.add_paragraph(text)
                if first_line:
                    first_line=""
                format_line(p,line["format"])
            
            
            #GET SPECIAL DAY HEADING
            #GET TRIODION DAY HEADING
            
            doc.add_paragraph()
            doc.add_paragraph()

    doc_filename = f"drafts\\calendar\\{year_no}_Календар_{mode}.docx"
    doc.save(doc_filename)
    print(f"Generated {doc_filename}")

    
