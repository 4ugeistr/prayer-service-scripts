import sys, re, mammoth

MONTH = 'Вересень' 
CLEANR = re.compile('<a id="_.*?</a>')
MOVE_SUB = re.compile("(<p>)\s*?(<i>)?\s*?(Понеділок|Вівторок|Середа|Четвер|П’ятниця|П'ятниця|Субота|Неділя)\s*?(</i>)*?\s*?(</p>)?(\n<p>)(<b>)?(<i>)?(\d{1,2})\s*?(<i>)?\s*?(</i>)?")
REPL='\g<1>'+MONTH+' \g<9> \g<3></p>\n<p>\g<7>\g<8>' 

filedoc = 'ustav-g.docx'
filehtm = "temp_full.html"

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
with open(filehtm+'1', 'w',encoding='utf-8') as f:
    f.writelines(file_lines)


# Очищуємо усі кастомні теги та коди від <p>, записуємо тільки непорожні рядки
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()
with open(filehtm, 'w',encoding='utf-8') as f:
    for line in file_lines:
        if not line.isspace():
            f.write(CLEANR.sub('', line))

# Переміщаємо дату зі строки 2 в строку 1
with open(filehtm,'r',encoding='utf-8') as f:
    file_content = f.read()

with open(filehtm,'w',encoding='utf-8') as f:
    f.write(MOVE_SUB.subn(REPL,file_content)[0])
    print(MOVE_SUB.subn(REPL,file_content)[1])

# Розбиваємо тимчасовий файл по днях
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


