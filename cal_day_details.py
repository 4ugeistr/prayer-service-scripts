import paschalia
import cal_generate, cal_convert
from datetime import datetime


def get_period(cur_date, mode):
    day_datails = paschalia.get_day_details(cur_date, mode)
    if day_datails[0]=="pascha":
        return "Від Пасхи до Зіслання"
    elif day_datails[0]=="lent":
        return "Великий Піст"
    elif day_datails[0]=="pentecost" and day_details[2]:
        return "Підготовчі тижні до Великого Посту"
    else:
        return "Звичайний період"

def get_weekday_as_word(cur_date):
    if cur_date.weekday()+1 in (0,7):
        return "Неділя"
    else:
        return paschalia.day_dic_reversed[cur_date.weekday()+1]


if __name__ == "__main__":

    year = 2026
    month_no = 6

    cur_date = datetime(year,month_no,29)
    mode = 'u'

    print(paschalia.get_day_details(cur_date, mode))


    print("День тижня:", get_weekday_as_word(cur_date))

    
    print("Глас:", paschalia.get_echos(cur_date, mode))
    if cur_date.weekday()+1 in (0,7): 
        print("Воскресне Євангеліє:", paschalia.get_resurrection_gospel(cur_date, mode))

    header_html = cal_generate.prepare_header_for_html(cur_date,mode)
    header_list = cal_generate.prepare_header_for_docx(cur_date,mode)

    for p in header_list:
        print(p.text)

    print(header_html)
    print("Ранг дня (proper):",cal_convert.find_highest_symbol(header_html))
    print("Ранг дня (simple):",cal_convert.find_highest_symbol_simple(header_html))
            
    if cal_convert.find_highest_symbol(header_list)!=" ":
        print("Ранг дня:",cal_convert.find_highest_symbol(header_html))
    day_lst = []
    for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
        
        cur_date = datetime(year,month_no,29)

        day_lst.append({"date":f"{year}-{month_no%02}-{d%02}",
                        "details":{
                            "число":d,
                            "день":d.weekday()+1,
                            "глас":paschalia.get_echos(cur_date, mode)
                            }
                        })
        
    



