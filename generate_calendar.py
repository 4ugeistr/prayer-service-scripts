import paschalia, calendar, docx, easygui, csv
from datetime import datetime

import ps_date_utils as du
from ps_docx_utils import format_line

#mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1



def get_matrix_full(csv_filename):
    matrix=[]
    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            matrix.append(row)
    return matrix

def insert_header_liturgy(doc,date):
    day_name = dt.day_dic_reversed[date.weekday()+1]
    #month_name = month_dic_reversed[month_no]
    week_no=paschalia.get_week(date,"","")[:2]

    #Субота, Неділя, Тиждень etc...
    special_dates=[{"date":datetime(year_no,12,25),
                     "holiday":"Різдво",
                     "holiday_locative":"Різдві",
                     "holiday_instrumental":"Різдвом"},
                    {"date":datetime(year_no+1,1,6),
                     "holiday":"Богоявленні",
                     "holiday_locative":"Богоявленні",
                     "holiday_instrumental":"Богоявленням"}]
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

def get_triodion_strings(date):
    lines = []
    triodion_params = paschalia.get_day_details(date)
    

    matrix = filter(lambda l:int(l[0])==date.month and int(l[1])==date.day,day_headers_menaion[1:])
    
    
    for l in matrix:
        lines.append({"text":l[9], "format":''.join(l[2:5])})
        #p_new=doc.add_paragraph(l[9])
        #format_line(p_new, ''.join(l[2:5]))
    return lines


def get_menaion_strings(date):
    lines = []
    matrix = filter(lambda l:int(l[0])==date.month and int(l[1])==date.day,day_headers_menaion[1:])
    for l in matrix:
        lines.append({"text":l[9], "format":''.join(l[2:5])})
        #p_new=doc.add_paragraph(l[9])
        #format_line(p_new, ''.join(l[2:5]))
    return lines

if __name__ == "__main__":

    day_headers_menaion = get_matrix_full("Місяцеслов-БД.csv")
    day_headers_triodion = get_matrix_full("Дні-Тріодь.csv")

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
            first_line = f"{d} "
            for line in get_menaion_strings(datetime(year_no,month_no,d)):
                text = first_line+ line["text"]
                p=doc.add_paragraph(text)
                if first_line:
                    first_line=""
                format_line(p,line["format"])
            
            #GET MENAION DAY HEADING
            #GET SPECIAL DAY HEADING
            #GET TRIODION DAY HEADING
            
            doc.add_paragraph()
            doc.add_paragraph()

    doc_filename = f"{year_no}_Календар_{mode}.docx"
    doc.save(doc_filename)
    print("Done!")

    