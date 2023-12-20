import re
from datetime import datetime, timedelta

#for 2023
#previous_pascha =  datetime(2022,4,24)
#pascha = datetime(2023,4,16)
#for 2024
previous_pascha =  datetime(2023,4,16)
pascha = datetime(2024,5,5)

paschalia_table = [
    {"mode":"u",
     "year":2022,
     "pascha":datetime(2023,4,24)},
    {"mode":"u",
     "year":2023,
     "pascha":datetime(2023,4,16)},
    {"mode":"u",
     "year":2024,
     "pascha":datetime(2024,5,5)},
    {"mode":"u",
     "year":2025,
     "pascha":datetime(2025,4,20)},
    
    {"mode":"g",
     "year":2022,
     "pascha":datetime(2023,4,17)},
    {"mode":"g",
     "year":2023,
     "pascha":datetime(2023,4,9)},
    {"mode":"g",
     "year":2024,
     "pascha":datetime(2024,3,31)},
    {"mode":"g",
     "year":2025,
     "pascha":datetime(2025,4,20)},
    ]



special_day_list = [
            "Субота перед Різдвом",
            "Навечір.я Різдва",
            "Субота по Різдві",
            "Субота перед Богоявленням",
            "Субота по Богоявленні",
            "Субота заупокійна",
            "Субота сиропусна",
            "Субота Лазарева",
            "Великий понеділок",
            "Великий вівторок",
            "Велика середа",
            "Великий четвер",
            "Велика п.ятниця",
            "Велика субота",
            "Світлий понеділок",
            "Світлий вівторок",
            "Світла середа",
            "Світлий четвер",
            "Світла п.ятниця",
            "Світла субота",
            "Понеділок Святого Духа",
            "Субота перед Воздвиженням",
            "Субота по Воздвиженні"
            ]
ending_fem_dic = {'1':"ша",
                  '2':"га",
                  '3':"тя",
                  '4':"та",
                  '5':"та",
                  '6':"та",
                  '7':"ма",
                  '8':"ма",
                  '9':"та",
                  '0':"та",
                  '40':"ва"}

ending_masc_dic = {'1':"ий",
                  '2':"ий",
                  '3':"ій",
                  '4':"ий",
                  '5':"ий",
                  '6':"ий",
                  '7':"ий",
                  '8':"ий",
                  '9':"ий",
                  '0':"ий"}
'''
for p in paschalia:
    p["meatfare_sunday"]=p["pascha"]-timedelta(days=7*8)
    p["cheesefare_sunday"]=p["pascha"]-timedelta(days=7*7)
    p["palm_sunday"]=p["pascha"]-timedelta(days=7)
    #lent_start = pascha - timedelta(days=7*7-1)
    p["pentecost"] = p["pascha"] + timedelta(days=7*7)
'''
for p in paschalia_table:
    p["meatfare_sunday"]=p["pascha"]-timedelta(days=7*8)
    p["cheesefare_sunday"]=p["pascha"]-timedelta(days=7*7)
    p["lent_start"] = p["pascha"] - timedelta(days=7*7-1)
    p["palm_sunday"]=p["pascha"]-timedelta(days=7)
    p["pentecost"] = p["pascha"] + timedelta(days=7*7)


def get_prev_next_pascha(cur_date, mode='u'):
    lst  =  list(filter(lambda p: (p['year'] == cur_date.year-1 or p['year'] == cur_date.year) and p['mode']==mode , paschalia_table))
    if len(lst)!=2:
        print(len(lst))
        raise Exception
    return lst

mode='u'
paschalia = get_prev_next_pascha(datetime(2024,1,1),mode)

'''
# 0 - previous, 1 - current
paschalia = [{},{}]
paschalia[0]["pascha"]=previous_pascha
paschalia[1]["pascha"]=pascha
'''


'''
0123456|789
0111111|122
'''

def get_week_from_50(cur_date):
    if cur_date > paschalia[1]["pentecost"]:
        days = (cur_date - paschalia[1]["pentecost"]).days
    else:
        days = (cur_date - paschalia[0]["pentecost"]).days
    modifier = 0 if days%7==0 else 1
    weeks = days // 7 + modifier    
    return weeks

    
'''
0123456|789
0000000|222
'''     
def get_week_from_pascha(cur_date):
    if cur_date > paschalia[1]["pascha"]:
        days = (cur_date - paschalia[1]["pascha"]).days
    else:
        days = (cur_date - paschalia[0]["pascha"]).days
    modifier = 0 if days // 7==0 else 1
    weeks = days // 7 + modifier
    return weeks



def get_echos(cur_date):
    echos_list = [1,2,3,4,5,6,7,8]
    #додати повернення None для дат які не мають Гласу?
    return echos_list[get_week_from_pascha(cur_date)%8-2]
    
def get_resurrection_gospel(cur_date):
    gospel_list = [1,2,3,4,5,6,7,8,9,10,11]
    if cur_date > paschalia[1]["pascha"]:
        days = (cur_date - paschalia[1]["pascha"]).days
    else:
        days = (cur_date - paschalia[0]["pascha"]).days

    if days ==7*1:
        return 1
    elif days == 7*2:
        return 3
    else:
        return gospel_list[get_week_from_pascha(cur_date)%11-2]
        
        
    

def get_echos_depr(date,mode='u'):
    if mode=='u':
        return (int(date.strftime("%U"))-25) % 8 + 1
    elif mode=='g':
        return (int(date.strftime("%U"))-24) % 8 + 1

def get_resurrection_gospel_depr(date,mode='u'):
    if mode=='u':
        return (int(date.strftime("%U"))-25) % 8 + 1
    elif mode=='g':
        return (int(date.strftime("%U"))-24) % 8 + 1
    

            
day_dic = {"Понеділок":1,
            "Вівторок":2,
            "Середа":3,
            "Четвер":4,
            "П'ятниця":5,
            "Субота":6,
            "Неділя":7}
day_dic_reversed = {v:k for k,v in day_dic.items()}

def get_day_label(cur_date):
    day_name = day_dic_reversed[cur_date.weekday()+1]
    #paschalia = get_prev_next_pascha(cur_date)
    #ending_fem = ending_fem_dic[str(week_no)[-1]]
    #ending_masc = ending_masc_dic[str(week_no)[-1]]
    
    if cur_date < paschalia[1]["meatfare_sunday"]-timedelta(days=7*3):
        weeks = ((cur_date - paschalia[0]["pentecost"]).days -1)// 7 + 1
        if cur_date.weekday()+1 in (1,2,3,4,5):
            ending_masc = ending_masc_dic[str(weeks)[-1]]
            return f"Тиждень {weeks:02}-{ending_masc} по Зісланні Святого Духа"
        else:
            ending_fem = ending_fem_dic[str(weeks)[-1]]
            return f"{day_name} {weeks:02}-{ending_fem} по Зісланні Святого Духа"
            
    if cur_date == paschalia[1]["meatfare_sunday"]-timedelta(days=7*3):
        return f"Неділя про Закхея"
    
    if paschalia[1]["meatfare_sunday"]-timedelta(days=7*3) < cur_date < paschalia[1]["meatfare_sunday"]-timedelta(days=7*2):
        weeks = ((cur_date - paschalia[0]["pentecost"]).days -1)// 7 + 1
        if cur_date.weekday()+1 in (1,2,3,4,5):
            ending_masc = ending_masc_dic[str(weeks)[-1]]
            return f"Тиждень {weeks:02}{ending_masc} по Зісланні Святого Духа"
        else:
            ending_fem = ending_fem_dic[str(weeks)[-1]]
            return f"{day_name} {weeks:02}{ending_fem} по Зісланні Святого Духа"
    
    if cur_date == paschalia[1]["meatfare_sunday"]-timedelta(days=7*2):
        return f"Неділя про Митаря і Фарисея"

    if paschalia[1]["meatfare_sunday"]-timedelta(days=7*2) < cur_date < paschalia[1]["meatfare_sunday"]-timedelta(days=7*1):
        weeks = ((cur_date - paschalia[0]["pentecost"]).days -1)// 7 + 1
        if cur_date.weekday()+1 in (1,2,3,4,5):
            ending_masc = ending_masc_dic[str(weeks)[-1]]
            return f"Тиждень {weeks:02}-{ending_masc} по Зісланні Святого Духа"
        else:
            ending_fem = ending_fem_dic[str(weeks)[-1]]
            return f"{day_name} {weeks:02}-{ending_fem} по Зісланні Святого Духа"
    
    if cur_date == paschalia[1]["meatfare_sunday"]-timedelta(days=7):
        return f"Неділя про Блудного сина"
    
    if paschalia[1]["meatfare_sunday"]-timedelta(days=7) < cur_date < paschalia[1]["meatfare_sunday"]:
        weeks = ((cur_date - paschalia[0]["pentecost"]).days -1)// 7 + 1
        if cur_date.weekday()+1 in (1,2,3,4,5):
            ending_masc = ending_masc_dic[str(weeks)[-1]]
            return f"Тиждень {weeks:02}-{ending_masc} по Зісланні Святого Духа"
        else:
            ending_fem = ending_fem_dic[str(weeks)[-1]]
            return f"{day_name} {weeks:02}-{ending_fem} по Зісланні Святого Духа"    
    
    
    if paschalia[1]["meatfare_sunday"]-timedelta(days=7) < cur_date < paschalia[1]["meatfare_sunday"]-timedelta(days=1):
        return f"Тиждень м'ясопусний"
    
    if cur_date == paschalia[1]["meatfare_sunday"]-timedelta(days=1):
        return f"Субота заупокійна"
    
    if cur_date == paschalia[1]["meatfare_sunday"]:
        return f"Неділя м'ясопусна, про Страшний суд."
    
    if paschalia[1]["cheesefare_sunday"]-timedelta(days=7) < cur_date < paschalia[1]["cheesefare_sunday"]-timedelta(days=1):
        return f"Тиждень сиропусний"
    
    if cur_date == paschalia[1]["cheesefare_sunday"]-timedelta(days=1):
        return f"Субота сиропусна. Всіх преподобних отців, що в подвизі просіяли."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]:
        return f"Неділя сиропусна, прощення."
    
    #Початок Великого Посту
    weeks = ((cur_date - paschalia[1]["cheesefare_sunday"]).days -1)// 7 + 1
    ending_masc = ending_masc_dic[str(weeks)[-1]]
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=1):
        return f"Початок св. Великого Посту. Строгий піст."

    if paschalia[1]["cheesefare_sunday"]+timedelta(days=1) <  cur_date < paschalia[1]["cheesefare_sunday"]+timedelta(days=6):
        return f"Тиждень {weeks}-{ending_masc} Великого посту."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=6):
        return f"Субота 1-ша Ведликого Посту, пам'ять великомученика Теодора Тирона."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7):
        return f"Неділя 1-ша Великого посту, Православ’я."

    if paschalia[1]["cheesefare_sunday"]+timedelta(days=7) < cur_date < paschalia[1]["cheesefare_sunday"]+timedelta(days=7+6):
        return f"Тиждень {weeks}-{ending_masc} Великого посту."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7+6):
        return f"Субота заупокійна. Прп. йоана Ліствичника"
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*2):
        return f"Неділя 2-га Великого посту."

    #Закінчення перших трьох місяців
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*2+6):
        return f"Субота заупокійна."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*3):
        return f"Неділя 3-тя Великого посту, Хрестопоклонна."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*3+6):
        return f"Субота заупокійна."

    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*4):
        return f"Неділя 4-та Великого посту."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*4+6):
        return f"Субота акафістова."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*5):
        return f"Неділя 5-та Великого посту."
    
    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*5+6):
        return f"Субота Лазарева."

    if cur_date == paschalia[1]["cheesefare_sunday"]+timedelta(days=7*5):
        return f"Неділя 6-та Великого посту."
    

    #Після Пасхи
    if cur_date == paschalia[1]["pascha"]+timedelta(days=7*4):
        return f"Неділя 4-та після Пасхи, розслабленого."

    if cur_date == paschalia[1]["pascha"]+timedelta(days=7*5):
        return f"Неділя 5-та після Пасхи, самарянки."
    


    
    '''
    elif cur_date > paschalia[1]["meatfare_sunday"] and cur_date < paschalia[1]["cheesefare_sunday"]:
        return f"00s"
                               
    elif cur_date > paschalia[1]["cheesefare_sunday"] and cur_date < paschalia[1]["palm_sunday"]:
        weeks = (cur_date - paschalia[1]["cheesefare_sunday"]).days // 7 + 1
        return f"{weeks:02}p"
    elif cur_date > paschalia[1]["pascha"] + timedelta(days=7) and cur_date < paschalia[1]["pentecost"]:
        weeks = (cur_date - paschalia[1]["pascha"]).days // 7 + 1
        return f"{weeks:02}e"

    elif cur_date > paschalia[1]["pentecost"]:
        weeks = (cur_date - paschalia[1]["pentecost"]).days // 7 + 1 
        return f"{weeks:02}d"
    else:
        print("Warning:")
        print(f"Week number not found for {cur_date}, {day_title}")
        return "err"
    '''
    print("Warning:")
    print(f"Week number not found for {cur_date}, {day_name}")
    return "err"

    
def get_week(cur_date, day_title,day_type):
    # пропускаємо номер тижня для 🕀 свят
    #if day_type=="#":
    #    return "***"
    
    # пропускаємо номер тижня для неділь (покривається наступним блоком)
    #if cur_date.weekday()==6:
    #    return "***"
    # пропускаємо номер тижня для спеціальних днів
    for word in special_day_list:
        if re.search(word,day_title):
            return "***"

    # дні від 1.01 до 31.12, в залежності від дат Пасхи на цей і на минулий рік
    if cur_date < paschalia[1]["meatfare_sunday"]-timedelta(days=7):
        weeks = ((cur_date - paschalia[0]["pentecost"]).days -1)// 7 + 1
        return f"{weeks:02}d"
    elif  cur_date > paschalia[1]["meatfare_sunday"]-timedelta(days=7) and cur_date < paschalia[1]["meatfare_sunday"]:
        return f"00m"

    elif cur_date > paschalia[1]["meatfare_sunday"] and cur_date < paschalia[1]["cheesefare_sunday"]:
        return f"00s"
                               
    elif cur_date > paschalia[1]["cheesefare_sunday"] and cur_date < paschalia[1]["palm_sunday"]:
        weeks = (cur_date - paschalia[1]["cheesefare_sunday"]).days // 7 + 1
        return f"{weeks:02}p"
    elif cur_date > paschalia[1]["pascha"] + timedelta(days=7) and cur_date < paschalia[1]["pentecost"]:
        weeks = (cur_date - paschalia[1]["pascha"]).days // 7 + 1
        return f"{weeks:02}e"

    elif cur_date > paschalia[1]["pentecost"]:
        weeks = (cur_date - paschalia[1]["pentecost"]).days // 7 + 1 
        return f"{weeks:02}d"
    else:
        print("Warning:")
        print(f"Week number not found for {cur_date}, {day_title}")
        return "err"

def get_week_related_label(cur_date):
    # пропускаємо номер тижня для 🕀 свят
    #if day_type=="#":
    #    return "***"
    
    # пропускаємо номер тижня для неділь (покривається наступним блоком)
    if cur_date.weekday()==6:
        return "***"
    # пропускаємо номер тижня для спеціальних днів
    for word in special_day_list:
        if re.search(word,day_title):
            return "***"

    # дні від 1.01 до 31.12, в залежності від дат Пасхи на цей і на минулий рік
    if cur_date < paschalia[1]["meatfare_sunday"]-timedelta(days=7):
        weeks = (cur_date - paschalia[0]["pentecost"]).days // 7 + 1
        return f"{weeks:02}d"
    elif  cur_date > paschalia[1]["meatfare_sunday"]-timedelta(days=7) and cur_date < paschalia[1]["meatfare_sunday"]:
        return f"00m"

    elif cur_date > paschalia[1]["meatfare_sunday"] and cur_date < paschalia[1]["cheesefare_sunday"]:
        return f"00s"
                               
    elif cur_date > paschalia[1]["cheesefare_sunday"] and cur_date < paschalia[1]["palm_sunday"]:
        weeks = (cur_date - paschalia[1]["cheesefare_sunday"]).days // 7 + 1
        return f"{weeks:02}p"
    elif cur_date > paschalia[1]["pascha"] + timedelta(days=7) and cur_date < paschalia[1]["pentecost"]:
        weeks = (cur_date - paschalia[1]["pascha"]).days // 7 + 1
        return f"{weeks:02}e"

    elif cur_date > paschalia[1]["pentecost"]:
        weeks = (cur_date - paschalia[1]["pentecost"]).days // 7 + 1 
        return f"{weeks:02}d"
    else:
        print("Warning:")
        print(f"Week number not found for {cur_date}, {day_title}")
        return "err"

