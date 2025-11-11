import re, easygui
from datetime import datetime, timedelta
from operator import itemgetter

day_dic = {"Понеділок": 1,
           "Вівторок": 2,
           "Середа": 3,
           "Четвер": 4,
           "П'ятниця": 5,
           "Субота": 6,
           "Неділя": 7}
day_dic_reversed = {v: k for k, v in day_dic.items()}

paschalia_dates_table = [
    {"mode": "u",
     "year": 2021,
     "pascha": datetime(2021, 5, 2)},
    {"mode": "u",
     "year": 2022,
     "pascha": datetime(2022, 4, 24)},
    {"mode": "u",
     "year": 2023,
     "pascha": datetime(2023, 4, 16)},
    {"mode": "u",
     "year": 2024,
     "pascha": datetime(2024, 5, 5)},
    {"mode": "u",
     "year": 2025,
     "pascha": datetime(2025, 4, 20)},
    {"mode": "u",
     "year": 2026,
     "pascha": datetime(2026, 4, 12)},
    {"mode": "u",
     "year": 2027,
     "pascha": datetime(2027, 5, 2)},
    
    {"mode": "g",
     "year": 2021,
     "pascha": datetime(2021, 4, 4)},
    {"mode": "g",
     "year": 2022,
     "pascha": datetime(2022, 4, 17)},
    {"mode": "g",
     "year": 2023,
     "pascha": datetime(2023, 4, 9)},
    {"mode": "g",
     "year": 2024,
     "pascha": datetime(2024, 3, 31)},
    {"mode": "g",
     "year": 2025,
     "pascha": datetime(2025, 4, 20)},
    {"mode": "g",
     "year": 2026,
     "pascha": datetime(2026, 4, 5)},
    {"mode": "g",
     "year": 2027,
     "pascha": datetime(2027, 3, 28)},
    
]

special_day_list = [
    "Субота перед Різдвом",
    "Навечір.я Різдва",
    "Субота по Різдві",
    "Субота перед Богоявленням",
    "Субота по Богоявленні",
    "Субота заупокійна",
    "Неділя про блудного сина",
    "Неділя м.ясопусна",
    #"Субота сиропусна",
    "Неділя сиропусна",
    "Субота Лазарева",
    "Квітна",
    "Великий понеділок",
    "Великий вівторок",
    "Велика середа",
    "Великий четвер",
    "Велика п.ятниця",
    "Велика субота",
    "НЕДІЛЯ ПАСХИ",
    "Світлий понеділок",
    "Світлий вівторок",
    "Світла середа",
    "Світлий четвер",
    "Світла п.ятниця",
    "Світла субота",
    #"Антипасха",
    "ЗІСЛАННЯ СВЯТОГО ДУХА",
    "Понеділок Святого Духа",
    "Субота перед Воздвиженням",
    "Субота по Воздвиженні",
]
ending_fem_dic = {
    '11': "та",
    '12': "та",
    '13': "та",
    '17': "та",
    '18': "та",
    '40': "ва",
    '1': "ша",
    '2': "га",
    '3': "тя",
    '4': "та",
    '5': "та",
    '6': "та",
    '7': "ма",
    '8': "ма",
    '9': "та",
    '0': "та",
}

ending_masc_dic = {'1': "ий",
                   '2': "ий",
                   '3': "ій",
                   '4': "ий",
                   '5': "ий",
                   '6': "ий",
                   '7': "ий",
                   '8': "ий",
                   '9': "ий",
                   '0': "ий"}

for p in paschalia_dates_table:
    p["meatfare_sunday"] = p["pascha"] - timedelta(days=7 * 8)
    p["cheesefare_sunday"] = p["pascha"] - timedelta(days=7 * 7)
    p["lent_start"] = p["pascha"] - timedelta(days=7 * 7 - 1)
    p["palm_sunday"] = p["pascha"] - timedelta(days=7)
    p["pentecost"] = p["pascha"] + timedelta(days=7 * 7)




# 0 - previous, 1 - current
def get_prev_next_pascha(cur_date, mode='u'):
    #prev = max(filter(lambda p: (p['pascha']-timedelta(days=7*9) <= cur_date) and p['mode']==mode , paschalia_dates_table), key = lambda x: x['pascha'])
    #next = min(filter(lambda p: (p['pascha']-timedelta(days=7*9) > cur_date) and p['mode']==mode , paschalia_dates_table), key = lambda x: x['pascha'])
    try:
        prev = max(filter(lambda p: (p['pascha'] <= cur_date) and p['mode'] == mode, paschalia_dates_table),
                   key=lambda x: x['pascha'])
        next = min(filter(lambda p: (p['pascha'] > cur_date) and p['mode'] == mode, paschalia_dates_table),
                   key=lambda x: x['pascha'])
    except ValueError as e:
        print("ValueError on data:",cur_date, mode)
        raise e

    lst = [prev, next]
    #legacy
    #lst  =  list(filter(lambda p: (p['year'] == cur_date.year-1 or p['year'] == cur_date.year) and p['mode']==mode , paschalia_dates_table))

    if len(lst) != 2:
        print(cur_date, mode)
        print(len(lst))
        if lst:
            for i in lst:
                print("Pascha date:", i["pascha"])
        raise Exception

    return lst


#TODO: треба буде позбутись статики.
#mode='u'
#paschalia_dates = get_prev_next_pascha(datetime(2024,1,1),mode)

'''
0123456|789
0111111|122
'''


def get_week_from_50(cur_date, paschalia_dates):
    days = (cur_date - paschalia_dates[0]["pentecost"]).days
    if days == 0:
        weeks = 1
    else:
        modifier = 0 if days % 7 == 0 else 1
        weeks = days // 7 + modifier
    return weeks


'''
0123456|789
0000000|222
'''


def get_week_from_pascha(cur_date, paschalia_dates):
    days = (cur_date - paschalia_dates[0]["pascha"]).days
    #modifier = 0 if days // 7==0 else 1
    modifier = 1
    weeks = days // 7 + modifier
    return weeks


def get_week_from_lent_start(cur_date, paschalia_dates):
    days = (cur_date - paschalia_dates[1]["lent_start"]).days
    #modifier = 0 if days // 7==0 else 1
    modifier = 1
    weeks = days // 7 + modifier
    return weeks


def get_echos(cur_date, mode):
    #print("getting echos for: ", cur_date, mode)
    paschalia_dates = get_prev_next_pascha(cur_date, mode)
    echos_list = [1, 2, 3, 4, 5, 6, 7, 8]

    if cur_date >= paschalia_dates[1]['palm_sunday']:
        return None

    days_from_pascha = (cur_date - paschalia_dates[0]['pascha']).days
    match days_from_pascha:
        case 0:
            return 1
        case 1:
            return 2
        case 2:
            return 3
        case 3:
            return 4
        case 4:
            return 5
        case 5:
            return 6
        case 6:
            return 8

    if days_from_pascha >= 7 and days_from_pascha < 7 * 2:
        return None

    return echos_list[get_week_from_pascha(cur_date, paschalia_dates) % 8 - 2]


def get_resurrection_gospel(cur_date, mode):
    paschalia_dates = get_prev_next_pascha(cur_date,mode)
    gospel_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]

    if cur_date >= paschalia_dates[1]['palm_sunday']:
        return None

    #щось незрозуміле, залишкове:
    #if cur_date == paschalia_dates[1]["cheesefare_sunday"] + timedelta(days=7 * 5):
    #    return f"Неділя 6-та Великого посту."

    match get_week_from_pascha(cur_date, paschalia_dates):
        case 0:  #Пасха
            return None
        case 2:  #Антипасха
            return 1
        case 3:  #мироносиць
            return 3
        case 4:
            return 4
        case 5:
            return 7
        case 6:
            return 8
        case 7:
            return 10
        case 8:  #50-ця
            return None
        case _:  #по 50-ці
            return gospel_list[get_week_from_50(cur_date, paschalia_dates) % 11 - 1]


'''
period = {lent|pascha|pentecost}
week from last large period start
week till next pascha (pentecost period only?)
weekday

'''


def get_day_details(cur_date, mode):
    if type(mode)!=str or len(mode)!=1:
        raise Exception(f"ERROR. Невалідне значення mode: {mode}")
    paschalia_dates = get_prev_next_pascha(cur_date,mode)
    weeks_till_lent = None
    if cur_date >= paschalia_dates[0]["pentecost"] and cur_date < paschalia_dates[1]["lent_start"]:
        period = 'pentecost'
        week = get_week_from_50(cur_date, paschalia_dates)

        weeks_till_lent = ((paschalia_dates[1]["lent_start"] - cur_date).days - 1) // 7 + 1
        weeks_till_lent = None if weeks_till_lent > 5 else weeks_till_lent
    elif cur_date > paschalia_dates[1]["cheesefare_sunday"] and cur_date <= paschalia_dates[1]["pascha"]:
        period = 'lent'
        week = ((cur_date - paschalia_dates[1]["cheesefare_sunday"]).days - 1) // 7 + 1
    else:
        period = 'pascha'
        week = get_week_from_pascha(cur_date, paschalia_dates)

    #порядок днів в Пасхальний період: 
    #0 - неділя
    #123456 - з понеділка по суботу

    #порядок днів в інший період: 
    #123456 - з понеділка по суботу
    #7 - неділя

    if (period == 'pascha') and cur_date.weekday() + 1 == 7 or cur_date == paschalia_dates[0]["pentecost"]:
        day = 0
    else:
        day = cur_date.weekday() + 1

    return period, week, weeks_till_lent, day


def get_week_code(cur_date, day_title, mode):
    paschalia_dates = get_prev_next_pascha(cur_date, mode=mode)
    #print(cur_date)
    # пропускаємо номер тижня для 🕀 свят
    #if day_type=="#":
    #    return "***"

    # пропускаємо номер тижня для неділь (покривається наступним блоком)
    #if cur_date.weekday()==6:
    #    return "***"
    # пропускаємо номер тижня для спеціальних днів
    for word in special_day_list:
        if re.search(word.lower(), day_title.lower()):
            return "***"

    # дні від 1.01 до 31.12, в залежності від дат Пасхи на цей і на минулий рік
    if cur_date > paschalia_dates[1]["meatfare_sunday"] - timedelta(days=7) and cur_date < paschalia_dates[1][
        "meatfare_sunday"]:
        return f"00m"

    elif cur_date > paschalia_dates[1]["meatfare_sunday"] and cur_date < paschalia_dates[1]["cheesefare_sunday"]:
        return f"00s"

    elif cur_date > paschalia_dates[1]["cheesefare_sunday"] and cur_date < paschalia_dates[1]["palm_sunday"]:
        return f"{get_week_from_lent_start(cur_date, paschalia_dates):02}p"

    elif cur_date >= paschalia_dates[0]["pascha"] + timedelta(days=7) and cur_date < paschalia_dates[0]["pentecost"]:
        return f"{get_week_from_pascha(cur_date, paschalia_dates):02}e"

    elif cur_date > paschalia_dates[0]["pentecost"]:
        return f"{get_week_from_50(cur_date, paschalia_dates):02}d"

    else:
        print("Warning:")
        print(f"(get_week) Week number not found for {cur_date}, {day_title}")
        return "err"


if __name__ == "__main__":
    mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u', 'g'])
    cur_date = datetime(2026, 5, 2)
    paschalia_dates = get_prev_next_pascha(cur_date, mode)
    print(get_day_details(cur_date, 'u'))
    print(get_echos(cur_date, 'u'))
