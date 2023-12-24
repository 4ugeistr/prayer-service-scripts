import sys, re, mammoth, logging
logging.basicConfig(filename='dates.log', filemode='w', format='%(message)s', level=logging.DEBUG)

CLEANR = re.compile('<a id="_.*?</a>')
MOVE_DATE = re.compile('(<p>\w+ )(\d{1,2})(.*?\n<p>.*?)(\d{1,2} \(\d{1,2}\)) ')

filedoc = 'ustav-u.docx'
filehtm = "temp.html"

# Перевизначаємо дефолтні (em & strong) теги для італіка та болда
style_map = """
i => i
b => b
"""
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

# Переміщаємо дату зі строки 2 в строку 1
with open(filehtm,'r',encoding='utf-8') as f:
    file_content=f.read()
    for item in MOVE_DATE.findall(file_content):
        logging.debug(item)
with open(filehtm,'w',encoding='utf-8') as f:
    result=MOVE_DATE.subn('\g<1>\g<4>\g<3>',file_content)
    f.write(result[0])
    print('Знайдено {} випадків для переміщення дати. Деталі в dates.log'.format(result[1]))

# Розбиваємо тимчасовий файл по днях
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()

file = None
file_index = 1

for line in file_lines: 
    if line.startswith('<p>Грудень'):
        file_index = int(line.rsplit(" ")[1])
        if file:
            file.close() 
        file = open('u{:02d}.html'.format(file_index), 'w',encoding='utf-8') 

    if file:    
        file.writelines(line)
        
if file:
    file.close()
