import sys, re, mammoth
from glob import glob
from datetime import datetime
import os
import easygui

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
                           .replace('<br />','')
                           .replace('liturgia=iz','liturgia=liz')
                           .replace('liturgia=vv','liturgia=lvv')
                           .replace('<h2>','\n<h2>')
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
            line = re.sub(CLEANA, '', line)
            if line.find('ustav') != -1 or line.find('antifon') != -1 or line.find('vhidne') != -1 or line.find('tropari') != -1 or line.find('prokimen') != -1 or line.find('apostol') != -1 or line.find('aleluia') != -1 or line.find('evanhelie') != -1 or line.find('dostoino') != -1 or line.find('prychasnii') != -1 or line.find('vidpust') != -1  or line.find('trysviate') != -1:
                line =  line.replace('<p>','').replace('</p>','')
            if line.find('#') != -1:
                line = re.sub(CLEANR, '', line)
            f.write(line)

# Розбиваємо тимчасовий файл по днях
with open(filehtm,'r',encoding='utf-8') as f:
    file_lines = f.readlines()

file = None

# ANMA - розбивка по послідовності індексів. Чутлива до пропущених днів, коли набрі днів не послідовний.
'''
file_index = 0
for line in file_lines: 
    line1 = line.strip()
    if line1.startswith('<ustav'):
        #file_index = int(line.rsplit(" ")[1])
        file_index = file_index + 1
        if file:
            file.close() 
        file = open(dirs+'/b{:02d}.html'.format(file_index), 'w',encoding='utf-8') 

    if file:    
        file.writelines(line)
        
    if line.startswith('</vidpust'):
        file.close()
        file = None
'''

#ANDU - розбивка на базі значення дати, зчитаної в описі дня
#потрібно слідкувати за mode = (u|g), може бути чутливе до зміни формату
cur_date=None
liturgy_template=None
for line in file_lines:
    #print('L: ',line)
    #if mode == 'u':
    re_result=re.search(f'^(?:<p>|<h\d>)(?:<i>)?(?:<b>)?{month_list_string} (<b>)?(\d+)(</b>)?',line)
    if re_result:
        cur_date = int(re_result.group(3))
    #elif mode == 'g':
    #   re_result=re.search(f'^(<p>|<h\d>)(<i>)?(<b>)?(\d+)',line)
    #   if re_result:
    #       cur_date = int(re_result.group(4))
    if line.startswith('<ustav'):
        lit_template=re.search(f'liturgia={lit_template_list}',line).group(1)
        if file:
            file.close()
        try:
            file = open(folder_name+'/b{:02d}.html'.format(cur_date), 'w',encoding='utf-8')
            print('/b{:02d}.html'.format(cur_date))
        except TypeError as e:
            print(line)
            print(cur_date)
            raise e
    if file:    
        file.writelines(line)
        
    if line.startswith('</vidpust'):
        #print(line)
        if lit_template=='liz_pascha':
            file.writelines([
            '<p><i>Тоді співаємо кінцеве:</i> Христос воскрес: <i>тричі, цілий тропар.</i> <i>А потім закінчуємо:</i></p>',
            '<p>I нам дарував життя вічне, поклоняємось Його тридневному воскресінню.</p>'])
        file.close()
        file = None


if file:
    file.close()
