import os, sys, re, shutil, mammoth
import easygui    
path = easygui.diropenbox()

# path = os.getcwd()
style_map = """
i => i
b => b
u => em
"""


print('Створюємо тимчасову та вихідну папки')
temppath = os.path.join(path, 'temp')
outpath = os.path.join(path, 'output')
if not os.path.exists(temppath):
    try:
        os.mkdir(temppath)
    except OSError as error:
        print(error)
if not os.path.exists(outpath):        
    try:
        os.mkdir(outpath)
    except OSError as error:
        print(error)    

print('Отримуємо перелік doc')
included_extensions = ['doc','docx']
doc_files = [fn for fn in os.listdir(path)
              if any(fn.endswith(ext) for ext in included_extensions)]

print('Конвертуємо doc в однойменні html')
for filedoc in doc_files:
    split_tup = os.path.splitext(filedoc)
    file_name = split_tup[0]
    file_ext = split_tup[1]
    if file_ext.startswith('.doc'):
        with open(path+'/'+filedoc, "rb") as docx_file:
            print(filedoc)
            result = mammoth.convert_to_html(docx_file, style_map=style_map)
        with open(temppath+'/'+file_name+'.html', "w", encoding='utf-8') as html_file:
            html_file.write(result.value)    

print('Отримуємо перелік html в тимчасовій папці')
included_extensions = ['html']
html_files = [fn for fn in os.listdir(temppath)
              if any(fn.endswith(ext) for ext in included_extensions)]
print('Форматуємо html файли')
for filehtml in html_files:
    with open(temppath+'/'+filehtml, 'r', encoding='utf-8') as f:
        file_lines = [''.join([x
                               .replace('<p>', '\n<p>')
                               .replace('<h1>', '\n<h1>')
                               .replace('<h2>', '\n<h2>')
                               .replace('<h3>', '\n<h3>')
                               ]) for x in f.readlines()]
    with open(temppath+'/'+filehtml, 'w', encoding='utf-8') as f:
        f.writelines(file_lines)

for filehtml in html_files:
    with open(temppath+'/'+filehtml, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()
    with open(temppath+'/'+filehtml, 'w',encoding='utf-8') as f:
        for line in file_lines:
            if not line.isspace():
                if line.upper().find("ПОВЕЧІР") != -1:
                    line = line.replace('p>', 'h2>')
                line = re.sub(re.compile('<a.*?</a>'), '', line) # очистка сміття типу <a id="_Hlk12345678"></a>
                f.write(line)

print('Формуємо вихідні файли')
for filehtml in html_files:
    
    with open(temppath+'/'+filehtml, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()
    
    file = None
    file_index = int(filehtml.split('.')[0])
    file = open(outpath+'/'+'l{:02d}.html'.format(file_index), 'w',encoding='utf-8') 
    file.writelines(file_lines)            
		
    file.close()

print('Готово!')

# shutil.rmtree(temppath, ignore_errors=True, onerror=None)

openpath = os.path.realpath(outpath)
os.startfile(openpath)
