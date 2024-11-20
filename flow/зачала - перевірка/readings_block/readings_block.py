import re,json

input_string='<i>Великі (Царські) часи.</i><br>На 1-му часі<br>Ап. – Ді. 33 зач.; 13, 25-33.<br>Єв. – Мт. 5 зач.; 3, 1-11.<br>На 3-му часі<br>Ап. – Ді. 42 зач.; 19, 1-8.<br>Єв. – Мр. 1 зач.; 1, 1-8.<br>На 6-му часі<br>Ап. – Рим. 91 зач.; 6, 3-11.<br>Єв. – Мр. 2 зач.; 1, 9-15.<br>На 9-му часі<br>Ап. – Тит. 302 зач.; 2, 11-14; 3, 4-7.<br>Єв. – Мт. 6 зач.; 3, 13-17.<br><i>Літургія св. Василія Великого з вечірнею.</i><br>Ап. – 1 Кор. 143 зач.; 9, 19-27.<br>Єв. – Лк. 9 зач.; 3, 1-18.<br><i>На осв. води:</i><br>Ап. – 1 Кор. 143 зач. (від половини); 10, 1-4.<br>Єв. – Мр. 2 зач.; 1, 9-11.'
ev_dict = {
    'Йо.':'iv',
    'Лк.':'lk',
    'Мт.':'mt',
    'Мр.':'mr'
    }
WORD_SPLIT=re.compile("(?<=\W)(\w+)(?=\W)")

#print(input_string)

with open('apostol.json', 'r',encoding='utf-8') as j:
     reading_list=json.loads(j.read())
for item in re.findall('(Ап\. – .*?)(?=<br>|$)',input_string):
    for r in reading_list:
        if WORD_SPLIT.findall(r['reading'])==WORD_SPLIT.findall(item[6:]):
            zach_filename = f"ap{r['zach_num']:0>3}{r['zach_abc']}"
            a_string=f"<a href='#' data-remote='../zachala/{zach_filename}.html' data-read='../benedict/{zach_filename}.html' data-toggle='modal' data-target='#exampleModal' >{item}</a>"
            input_string=input_string.replace(item,a_string)

with open('evanhelie.json', 'r',encoding='utf-8') as j:
     reading_list=json.loads(j.read())
for item in re.findall('(Єв\. – .*?)(?=<br>|$)',input_string):
    for r in reading_list:
        if WORD_SPLIT.findall(r['reading'])==WORD_SPLIT.findall(item[6:]):
            zach_filename = f"{ev_dict[r['book']]}{r['zach_num']:0>3}{r['zach_abc']}"
            a_string=f"<a href='#' data-remote='../zachala/{zach_filename}.html' data-read='../benedict/{zach_filename}.html' data-toggle='modal' data-target='#exampleModal' >{item}</a>"
            input_string=input_string.replace(item,a_string)

#print(input_string)
