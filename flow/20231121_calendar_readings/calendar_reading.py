import docx, re, csv, calendar

docx_filename = "2024NJUL_orig.docx"

month_list = {'Січень':1,
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
'''
month_list = dict(month_list_lower)
for k,v in month_list_lower.items():
    month_list[k.upper()]=v
'''

reading_indicator_string = '^(Єв\. –|'\
                                'Ап\. –|'\
                                'Єв\. -|'\
                                'Ап\. -|'\
                                'Читання на Шостому часі:|'\
                                'Час Шостий:|'\
                                'Літ.:|'\
                                'Вечірня з Літургією св. Василія Великого.|'\
                                'На Літургії св. Василія Великого з вечірнею.|'\
                                'Літургія Передосвячених Дарів.|'\
                                'Літургія св. Йоана Золотоустого.|'\
                                'Літургія св. Василія Великого.|'\
                                'На вечірні:|'\
                                'Вечірня:|'\
                                'Утр:|'\
                                'Утр.:|'\
                                'На освячення води|'\
                                'На вмиванні:|'\
                                'По вмиванні:|'\
                                'Ряд.:|'\
                                'Ап.:|'\
                                'Апп.:|'\
                                'Рівноап.:|'\
                                'Предтечі:|'\
                                'Новоліттю:|'\
                                'Св.:|'\
                                'Свв.:|'\
                                'Прп.:|'\
                                'Ісп.:|'\
                                'Мч.:|'\
                                'Мчч.:|'\
                                'Мчц.:|'\
                                'Свщнмч:|'\
                                'Свщмч:|'\
                                'Свщмч.:|'\
                                'Свщнмчч.:|'\
                                'Влкмч.:|'\
                                'На \d-му часі|'\
                                'Отців:|'\
                                'Отцям:|'\
                                'Богородиці:|'\
                                'Собору:|'\
                                'Оновлення:|'\
                                'Анни:|'\
                                'Св. Миколая:|'\
                                'Свята:|'\
                                'Суботи:|'\
                                'Всім святим:|'\
                                'За упокій:'\
                                ')'


month_list_string='('+'|'.join([x.lower() for x in month_list.keys()])+')'

day_list_string="(Понеділок|Вівторок|Середа|Четвер|П’ятниця|П'ятниця|Субота|Неділя)"
doc = docx.Document(docx_filename)

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


def get_matrix_full(csv_filename):
    matrix=[]
    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            matrix.append(row)
    return matrix

#matrix[0]["days"][0]["header"]
def get_day_dic(month,day):
    for m in matrix:
        if month == m["month"]:
            for d in m["days"]:
                if day == d["day"]:
                    return d

def get_saints(month,day):
    saint_string=""
    for row in saint_matrix[1:]:
        if int(row[0])==month and int(row[1])==day:
            if saint_string:
                saint_string+='\n'+row[8]
            else:
                saint_string=row[8]
    return saint_string


saint_matrix = get_matrix_full("Місяцеслов-БД.csv")


matrix = []

month_no = None
day_no = None
header_found = False
header = None
reading_found = False
reading = None
reading_first = False
day = None
i=0
for p in doc.paragraphs:
#for p in doc.paragraphs[:60]:
    #print(i,p.text[:40])
    i+=1
    re_result = re.search(month_list_string,p.text.lower())
    if re_result:
        month_no = month_list[re_result.group(1).capitalize()]
        matrix.append({"month":month_no,"days":[]})
        month = matrix[-1]["days"]


    #знаходимо день
    re_result=re.search('^'+day_list_string+'$',p.text)
    if re_result:
        reading_found = False
        week_day = re_result.group(1)
        delete_paragraph(p)
        continue
    
        
    #знаходимо початок блоку дня. ігноруємо день тижня
    re_result = re.search(r"^(\d{1,2}) (.*)$",p.text)
    if re_result:
        month.append({"day":int(re_result.group(1)),"header":re_result.group(2)})
        day_no = int(re_result.group(1))
        header = month[-1]
        header_found = True
        reading_found = False
        reading = None

        p.insert_paragraph_before(re_result.group(1)+" "+week_day)
        delete_paragraph(p)
        
        continue

    if reading_first:
        reading_first = False
    
    #знаходимо початок блоку читань
    re_result=re.search(reading_indicator_string,p.text)
    if re_result:
        if "reading" in header:
            header["reading"]+='\n'+p.text
        else:
            header["reading"]=p.text
        header_found = False
        reading_found = True
        reading_first = True
        continue
    
    
        
    if header_found and p.text:
        header["header"] +='\n'+p.text
        delete_paragraph(p)
        continue
        
        
    elif reading_first and p.text:
        print("!inserting headers!")
        header["reading"]+='\n'+p.text
        saints_string=get_saints(month_no, day_no)
        saints_list = saints_string.split('\n')
        for item in saints_list:
            p.insert_paragraph_before(item)

doc.save("2024NJUL.docx")

csv_data=[]
year_no=2024
for month_no in range(1,13):
    for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
        csv_data.append([month_no,d])
        #print(month_no,d)
        csv_data[-1].append(get_day_dic(month_no,d)["header"].replace('\n',' '))
        csv_data[-1].append(get_saints(month_no,d).replace('\n',' '))
        
csvfile = open("test.csv",'w',newline='\n',encoding='utf8')
spamwriter=csv.writer(csvfile,delimiter='|',quotechar="\"", quoting=csv.QUOTE_MINIMAL)
spamwriter.writerows(csv_data)
csvfile.close()        

    
