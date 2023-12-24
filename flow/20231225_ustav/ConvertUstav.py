import sys, os, re, mammoth, logging
logging.basicConfig(filename='dates.log', filemode='w', format='%(message)s', level=logging.DEBUG)

'''
Запускаємо з параметрами <u|g> <Місяць> <ім'я doc файлу>
python ConvertUstav.py u Грудень Устав-Грудень.docx
'''

MONTH = 'Січень'
YEAR = 2023
MODE= 'g'
filehtm = f"temp_{YEAR}_{MODE}.html"
filedoc = f'ustav-{YEAR}-{MODE}.docx'
#filedoc = f'2023 - КАЛЕНДАР УГКЦ З УСТАВОМ_20221031.docx'

if len(sys.argv)>1:
    assert sys.argv[1]=='u' or sys.argv[1]=='g' 
    MODE=sys.argv[1]
    if sys.argv[2]:
        MONTH=sys.argv[2]
    if sys.argv[3]:
        filedoc=sys.argv[3]
    
# [0] - pattern, [1] - replacement pattern
re_pattern={
    "u":[
        '(<p>\w+ )(\d{1,2})(.*?)(\n<p>.*?)(\d{1,2}) (\(\d{1,2}\)) ',
        '\g<1><b>\g<5></b> \g<6>,\g<3>\n<hr>\g<4>'
        ],
    "g":[
        "(<p>)\s*?(<i>)?\s*?"\
        "(Понеділок|Вівторок|Середа|Четвер|П’ятниця|П'ятниця|Субота|Неділя)"\
        "\s*?(</i>)*?\s*?(</p>)?"\
        "(\n<p>)(<b>)?(<i>)?(\d{1,2})\s*?(<i>)?\s*?(</i>)?",
        '\g<1>'+MONTH+' <b>\g<9></b>, \g<3></p>\n<hr>\n<p>\g<7>\g<8>' 
        ]
    }


CLEANR = re.compile('<a id="_.*?</a>')

word_list = ['Вечірня', 'Вечірня з Літургією св. Василія Великого',
             'Утреня', 'Часи',
             'Літургія Передосвячених Дарів',
             'Літургія',
             'Літургія св. Йоана Золотоустого','Літургія св. Василія Великого']

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

month_list_caps = {'Січень':1,
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

month_list_string='('+'|'.join(month_list.keys())+')'

day_list_string="(Понеділок|Вівторок|Середа|Четвер|П’ятниця|П'ятниця|Субота|Неділя)"

# Перевизначаємо дефолтні (em & strong) теги для італіка та болда
style_map = """
i => i
b => b
"""


# Функція зливає в fallout.html рядки з помилками чи нестандартним форматуванням.
def get_fallout(filehtm, word=word_list[0], pattern=re_pattern['g'][0]):
    with open(filehtm, "r",encoding='utf-8') as f:
        l=f.readlines()
    full_dict={i:l[i]+l[i+1] if i<len(l)-1 else l[i] for i in range(len(l))}    
    filtered_dict={k:v for k,v in full_dict.items() if re.search(word,v.split('\n')[0])}
    filtered_dict_after_pattern={k:v for k,v in full_dict.items() if re.search(pattern,v)}
    fallout={k:v for k,v in filtered_dict.items() if not k in filtered_dict_after_pattern}
    with open('fallout.html', "w",encoding='utf-8') as f:
        f.writelines([f'{k} {v}' for k,v in fallout.items()])
    return len(fallout)

def reread():
    with open(filehtm, "r",encoding='utf-8') as html_file:
        content=html_file.read()
    return content
# функція для футноутів
def move_footnotes(filehtm,content):
    with open(filehtm, "r",encoding='utf-8') as html_file:
        content=html_file.read()
    with open(filehtm+'_footnotes', "w",encoding='utf-8') as html_file:
        html_file.write(content)

    
    RE_FOOTNOTE=re.compile('<li id="footnote-(\d+)">(.*?)<a href="#footnote-ref-\d+">.</a></p></li>')
    with open(filehtm, "r",encoding='utf-8') as html_file:
        content=html_file.read()
    #реконструюємо нормальні переноси в футноутах вкінці файлу
    content=re.sub('<ol><li','<ol>\n<li',content)
    content=re.sub('</li><li','</li>\n<li',content)
    content=re.sub('</li></ol>','</li>\n</ol>',content)
    content=re.sub('(<li.*?>)\n','\g<1>',content)
  
    #переносимо футноути в змінну
    footnote_list=re.findall('<li.*</li>',content)
    footnote_dict={}
    for item in footnote_list:
        footnote_dict[RE_FOOTNOTE.search(item).group(1)]= RE_FOOTNOTE.search(item).group(2)
    
    #стираємо футноути вкінці файлу
    #content=re.sub('\n<p></p><ol>.*</ol>','',content,flags=re.DOTALL)
    content=re.sub('<ol>.*</ol>','',content,flags=re.DOTALL)
    content=re.sub('(<sup>)<a href="#footnote-(\d+).*?(</sup>)','\g<1><i>\g<2></i>\g<3>',content)
    #для кожного дня додаємо футноути в кінець та міняємо № на *
    #NB: без 31 грудня. 
    days=re.findall(month_list_string+'( <b>.*?)(?=<p>'+month_list_string+')',content,flags=re.DOTALL)
    print(len(days))
    for day in days:
        note_list=re.findall('<sup><i>(\d+)</i></sup>',day[1])
        if note_list:
            i=1
            asterisks='*'
            #print(f'{day[0]} {day[1][4:6]} {note_list}')
            day_tmp=day[1]+'<hr>'
            content=content.replace(day[1],day_tmp)
            for note in note_list:
                content_bak=content
                day_tmp2=re.sub(f'<sup><i>{note}</i></sup>',f'<sup><i>{asterisks}</i></sup>',day_tmp) \
                          +'<p>'+f'<sup><i>{asterisks}</i></sup> - '+footnote_dict[note][3:]+'\n'
                content=content.replace(day_tmp,day_tmp2)
                day_tmp=day_tmp2
                asterisks+='*'        
    with open(filehtm, "w",encoding='utf-8') as html_file:
        html_file.write(content)
    

# Конвертуємо док в тичасовий великий хтмл
with open(filedoc, "rb") as docx_file:
    result = mammoth.convert_to_html(docx_file, style_map=style_map)
with open(filehtm, "w",encoding='utf-8') as html_file:
    html_file.write(result.value)

# Розбиваємо файл на стрічки, попутно міняємо кутові дужки та видаляємо всі <br /> 
with open(filehtm, 'r',encoding='utf-8') as f:
    file_lines = [''.join([x
                           .replace('<p>', '\n<p>')
                           .replace('<h1>', '\n<h1>')
                           .replace('&lt;','<')
                           .replace('&gt;','>')
                           .replace('<br />','\n<p>')
						   .replace(chr(160),chr(32))
                           ]) for x in f.readlines()]
with open(filehtm, 'w',encoding='utf-8') as f:
    f.writelines(file_lines)


# Очищуємо усі кастомні теги та коди від <p>, записуємо тільки непорожні рядки
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()
with open(filehtm, 'w',encoding='utf-8') as f:
    for line in file_lines:
        if not line.isspace():
            line = re.sub(CLEANR, '', line)
            f.write(line)


with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()


# Формування fallout.html з вийнятками
if MODE=='g':
    fallout_no = get_fallout(filehtm,
              word="(<p>)\s*?(<i>)?\s*?(Понеділок|Вівторок|Середа|Четвер|П’ятниця|П'ятниця|Субота|Неділя)",
              pattern=re_pattern[MODE][0])
    print(f'{fallout_no} значень не потрапило в regexp. Див fallout.html')

# Чистимо найбільш часті криві послідовності тегів форматування
# Наступний блок вичистить 95% фороматування для заголовків рубрик уставу ("Утреня:"...)
# Лишаться випадки, коли за рубрикою одразу йдуть якісь вказівки курсивом (напр. про багряні ризи, поклони).
with open(filehtm, "r",encoding='utf-8') as html_file:
    content=html_file.read()
    content=re.sub('<i> *</i>',' ',content)
    content=re.sub('<b> *</b>',' ',content)
    content=re.sub('<i>: *</i>',': ',content)
    content=re.sub('<b>: *</b>',': ',content)
    content=re.sub(': *<i>',':<i> ',content)
    content=re.sub(': *</i>',':</i> ',content)
    content=re.sub('</i>:',':</i>',content)
    content=re.sub(' +',' ',content)
    # пробіли перед кінцем абзацу
    content=re.sub(' *</p>','</p>',content)
    # подвійні пробіли
    content=re.sub('<p> *<i>','<p><i>',content)
    # пусті строки
    content=re.sub('<p>(</p>)?\n','',content)
    # додаємо <hr> перед "Вечірня:"
    content=re.sub('(<p><i>)(Вечірня|Літургія Передосвячених)','<hr>\n\g<1>\g<2>',content)
    # прибираємо зайві пробіли до і після заголовку служби
    for word in word_list:
        content=re.sub(f'<i> ({word})',f' <i>\g<1>',content)
        content=re.sub(f'<b>(<i>{word}:</i>) *</b>','\g<1>',content)
    # прибираємо подвійні sup
    content=re.sub('(<sup>)+','<sup>',content)
    content=re.sub('(</sup>)+','</sup>',content)
    # виправляємо косі апострофи, на прямі
    content=re.sub('’',"'",content)

    #перетворюємо UPPERCASE, lowercase місяці в Capitalize:
    for m in month_list.keys():
        content= re.sub(m.upper(),m,content)
        content= re.sub(m.lower(),m,content)
        
with open(filehtm, "w",encoding='utf-8') as html_file:
    html_file.write(content)


with open(filehtm,'r',encoding='utf-8') as f:
    content=f.read()
    for item in re.findall(re_pattern[MODE][0],content):
        logging.debug(item)
    with open(filehtm+'_debug','w',encoding='utf-8') as f1:
        f1.write(content)
with open(filehtm,'w',encoding='utf-8') as f:
    # Переміщаємо дату зі строки 2 в строку 1
    if MODE=='u':
        content,sub_qty=re.subn(re_pattern[MODE][0],re_pattern[MODE][1],content)
    if MODE=='g':
        month_parts = re.split(f'(?:<p><b>|<h1>){month_list_string}(?:</b></p>|</h1>)\n',content)
        if len(month_parts)!=24+1:
            found_months_qty = (len(month_parts)-1) /2
            found_months = month_parts[1::2]
            print(f'Бракує місяців. Знайшло: {found_months_qty}')
            print(found_months)
            raise Exception
        content=month_parts[0]
        month_parts=month_parts[1:]
        month_parts_dict={month_parts[2*i]:month_parts[2*i+1] for i in range(int(len(month_parts)/2))}
        sub_qty_total=0

        pattern="(<p>)\s*?(<i>)?\s*?"\
            "(Понеділок|Вівторок|Середа|Четвер|П’ятниця|П'ятниця|Субота|Неділя)"\
            "\s*?(</i>)*?\s*?(</p>)?"\
            "(\n<p>)(<b>)?(<i>)?(\d{1,2})\s*?(<i>)?\s*?(</i>)?"
        
        for month,month_part in month_parts_dict.items():
            content+=f'<h1>{MONTH}</h1>\n'

            pattern_sub='\g<1>'+month+' <b>\g<9></b>, \g<3></p>\n<hr>\n<p>\g<7>\g<8>' 
            
            month_part,sub_qty=re.subn(pattern,pattern_sub,month_part)
            sub_qty_total+=sub_qty
            content+=month_part
        sub_qty=sub_qty_total
    content = re.sub('<p> ','<p>',content)
    content=re.subn("((\n.*?)(</p>)?(\n<hr>))","\g<2></p>\g<4>",content)[0]

    

    f.write(content)
    print(f'Знайдено {sub_qty} випадків для переміщення дати. Деталі в dates.log')

#розкидуємо футноути по днях
move_footnotes(filehtm,content)
    
# виділяємо червоним неділі
with open(filehtm,'r',encoding='utf-8') as f:
    content=f.read()
content=re.subn("((\n<p>)(.*?, Неділя.*?)(</p>\n<hr>))","\g<2><span>\g<3></span>\g<4>",content)[0]
with open(filehtm,'w',encoding='utf-8') as f:
    f.write(content)


# Розбиваємо тимчасовий файл по днях - 1 місяць
'''
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()

file = None
file_index = 1
for line in file_lines:
    if line.startswith('<p>'+MONTH):
        file_index = int(line.rsplit(" ")[1])
        if file:
            file.close() 
        file = open('u{:02d}.html'.format(file_index), 'w',encoding='utf-8') 
    if file:    
        file.writelines(line)
if file:
    file.close()
'''
# Розбиваємо тимчасовий файл по днях - всі місяці
for i in range(12):
    try:
        os.mkdir(f'{i+1:02}')
    except FileExistsError:
        #print(f'Directory {i+1:02} already exists.')
        pass
file = None
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()
for line in file_lines:
    if re.search('<p>.*?'+month_list_string,line):
        day = int(re.search('<b>(\d{1,2})</b>',line)[1])
        month = month_list[re.search(month_list_string,line)[0]]
        if file:
            file.close()
        file = open(f'{month:02}\\u{day:02}.html', 'w',encoding='utf-8')
    if file:
        file.writelines(line)
if file:
    file.close()        
