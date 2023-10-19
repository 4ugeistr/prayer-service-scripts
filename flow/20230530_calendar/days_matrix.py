from datetime import datetime,timedelta
import csv,re

start_date = datetime(2023, 1, 1)
end_date = datetime(2023, 8, 31)

# 0 - previous, 1 - current
paschalia = [{},{}]
paschalia[0]["pascha"]=datetime(2022,4,24)
paschalia[1]["pascha"]=datetime(2023,4,16)

for p in paschalia:
    
    p["meatfare_sunday"]=p["pascha"]-timedelta(days=7*8)
    p["cheesefare_sunday"]=p["pascha"]-timedelta(days=7*7)
    p["palm_sunday"]=p["pascha"]-timedelta(days=7)
    #lent_start = pascha - timedelta(days=7*7-1)
    p["pentecost"] = p["pascha"] + timedelta(days=7*7)    

cur_date = datetime.now()


# ’ - неправильний апостроф. треба міняти на '
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

def get_week(cur_date, day_title,day_type):
    # пропускаємо номер тижня для 🕀 свят
    if day_type=="#":
        return "***"
    # пропускаємо номер тижня для неділь (покривається наступним блоком)
    if cur_date.weekday()==6:
        return "***"
    # пропускаємо номер тижня для спеціальних днів
    for word in word_list:
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


day_matrix=[]
for i in range((end_date-start_date).days):
    cur_date = start_date + timedelta(days=i)
    day_matrix.append([f"{cur_date.year}{cur_date.month:02}{cur_date.day:02}",cur_date.weekday()+1,get_week(cur_date,"","")])    

csvfile = open('weektest.txt','w',newline='',encoding='utf8')
spamwriter=csv.writer(csvfile,delimiter='|',quotechar='"', quoting=csv.QUOTE_MINIMAL)
spamwriter.writerows(day_matrix)
csvfile.close()
