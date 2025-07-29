import sys, re, mammoth
from glob import glob
from datetime import datetime
from bs4 import BeautifulSoup
#import calendar
import os
import easygui
from lit_check_integrity import checkLiturgyIntegrity

CLEANA = re.compile('<a.*?</a>')
CLEANR = re.compile('<.*?>')

#mode='g'
mode_dic = {'НЮ':'u',
            'ГР':'g',}
mode_dic_reversed = {v:k for k,v in mode_dic.items()}

#mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
#dirs = easygui.diropenbox()
#paths = glob(f'{dirs}\\*.docx', recursive=True)
#filedoc = paths[0]

filedoc = easygui.fileopenbox(
    title="Select a .docx file",
    filetypes=["*.docx"],
    default="*.docx"
)
#dirs = os.path.dirname(filedoc)
os.chdir(os.path.dirname(filedoc))
filedoc = os.path.basename(filedoc)
print(filedoc)

mode = mode_dic[filedoc.split('.')[0][-2:]]
month_no = int(filedoc[:2])
print(f'Month: {month_no}')
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1

if not os.path.exists('drafts'):
    os.makedirs('drafts')
folder_name=f'drafts\\{year_no}-{month_no:02}-{mode}'
if not os.path.exists(folder_name):
    os.makedirs(folder_name)


filehtm = folder_name + '/temp.html'

# Перевизначаємо дефолтні (em & strong) теги для італіка та болда
style_map = """
i => i
b => b
h2 => p
"""
#style_map = ''

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

lit_template_list='(liz_pascha|liz|lvv)'

month_list_string='('+'|'.join(month_list.keys())+')'



# Валідуємо docx на рахунок цілісності тегів
checkLiturgyIntegrity(filedoc)

# Конвертуємо док в тичасовий великий хтмл
with open(filedoc, "rb") as docx_file:
    result = mammoth.convert_to_html(docx_file, style_map=style_map).value


with open(filehtm, "w",encoding='utf-8') as html_file:
    html_file.write(result)
with open(filehtm+'_init.html', "w",encoding='utf-8') as html_file:
    html_file.write(result)

# Розбиваємо файл на стрічки, попутно міняємо кутові дужки та видаляємо всі <br /> 
with open(filehtm, 'r',encoding='utf-8') as f:
    text = f.read().replace('<p>', '\n<p>')\
                   .replace('<h1>', '\n<h1>')\
                   .replace('&lt;','<')\
                   .replace('&gt;','>')\
                   .replace('<br />','')\
                   .replace('liturgia=iz','liturgia=liz')\
                   .replace('liturgia=vv','liturgia=lvv')\
                   .replace('<h2>','\n<h2>')\
                   .replace(chr(160),chr(32))

with open(filehtm, 'w',encoding='utf-8') as f:
    f.writelines(text)

# Очищуємо усі кастомні теги та коди від <p>, записуємо тільки непорожні рядки
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()
    
with open(filehtm, 'w',encoding='utf-8') as f:
    for line in file_lines:
        if not line.isspace():
            line = re.sub(CLEANA, '', line)
            if line.find('ustav') != -1 or line.find('antifon') != -1 or line.find('vhidne') != -1 or line.find('tropari') != -1 or line.find('prokimen') != -1 or line.find('apostol') != -1 or line.find('aleluia') != -1 or line.find('evanhelie') != -1 or line.find('dostoino') != -1 or line.find('prychasnii') != -1 or line.find('vidpust') != -1  or line.find('trysviate') != -1:
                line =  line.replace('<p>','').replace('</p>','')
            if line.find('#') != -1:
                line = re.sub(CLEANR, '', line)
            f.write(line)

with open(filehtm,'r',encoding='utf-8') as f:
    with open(filehtm+'bak', 'w',encoding='utf-8') as f2:
        f2.write(f.read())
        
def check_for_broken_tags(filehtm):
    with open(filehtm,'r',encoding='utf-8') as f:
        lines = f.readlines()
        err_lines = []
        for line in lines:
            if re.search(r'[a-zA-Z0-9]',line) and not re.search(r'^<.*>$',line) and not "#" in line:
                err_lines.append(line)
        print("err_lines:",len(err_lines),err_lines)
        if err_lines:
            for item in err_lines:
                print("ERR:",item)
            raise Exception("Увага: в строках вище помилки в користувацьких html мітках.")
        
check_for_broken_tags(filehtm)


#робимо весь чорний шрифт жирним.
black_list_set = {'i','b','vidpust'}
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()
    for line in range(len(file_lines)):
        if file_lines[line].startswith('<p>') and not file_lines[line].startswith('<p><i>Священ'):
            #print(len(file_lines[line]),file_lines[line][-1],ord(file_lines[line][-1]))
            soup = BeautifulSoup(file_lines[line], 'html.parser')
            for tag in soup.find_all(string=True):
                if not (len(str(tag))==1 and ord(str(tag))==10):                    
                    parent_tags_set = {x.name for x in tag.parents}
                    if not (black_list_set & parent_tags_set):
                        tag.wrap(soup.new_tag('b'))                        
            file_lines[line] = str(soup)
with open(filehtm,'w',encoding='utf-8') as f:
    f.writelines(file_lines)



with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()
# Розбиваємо тимчасовий файл по днях 
file = None

#ANDU - розбивка на базі значення дати, зчитаної в описі дня
cur_date=None
liturgy_template=None

for line in file_lines:
    #print(line[:40])
    re_result=re.search(r'^(?:<p>|<h\d>)(?:<i>)?(?:<b>)?'+f'{month_list_string}'+r' (<b>)?(\d+)(</b>)?',line)
    if re_result:
        cur_date = int(re_result.group(3))
        if file:
            file.close()

    if line.startswith('<ustav'):
        lit_template=re.search(f'liturgia={lit_template_list}',line).group(1)
        try:
            file = open(f'{folder_name}/b{cur_date:02d}.html', 'w',encoding='utf-8')
            print('/b{:02d}.html'.format(cur_date))
        except TypeError as e:
            print(line)
            print(cur_date)
            raise e
            
    if file and not file.closed:    
        file.writelines(line)

if file:
    file.close()

print("Файли успішно сконвертовані!")
