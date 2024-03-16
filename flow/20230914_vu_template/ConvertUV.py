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
print(f'   {len(doc_files)} файлів')

print('Конвертуємо doc в однойменні html')
for filedoc in doc_files:
    split_tup = os.path.splitext(filedoc)
    file_name = split_tup[0]
    file_ext = split_tup[1]
    if file_ext.startswith('.doc'):
        with open(path+'/'+filedoc, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, style_map=style_map)
        with open(temppath+'/'+file_name+'.html', "w", encoding='utf-8') as html_file:
            html_file.write(result.value)    

print('Отримуємо перелік html в тимчасовій папці')
included_extensions = ['html']
html_files = [fn for fn in os.listdir(temppath)
              if any(fn.endswith(ext) for ext in included_extensions)]
print(f'   {len(html_files)} файлів')

print('Форматуємо html файли')
for filehtml in html_files:
    with open(temppath+'/'+filehtml, 'r', encoding='utf-8') as f:
        file_lines = [''.join([x
                               .replace('<p>', '\n<p>')
                               .replace('<h1>', '\n<h1>')
                               .replace('<h2>', '\n<h2>')
                               .replace('<h3>', '\n<h3>')
                               .replace(chr(160),chr(32))
                               ]) for x in f.readlines()]
    with open(temppath+'/'+filehtml, 'w', encoding='utf-8') as f:
        f.writelines(file_lines)

for filehtml in html_files:
    file_content=''
    with open(temppath+'/'+filehtml, 'r', encoding='utf-8') as f:
        file_content = f.read()
        
    with open(temppath+'/'+filehtml+'_bak', 'w', encoding='utf-8') as f:
        f.write(file_content)    
    if file_content.find('<h') == -1:
        print(f"    WARNING: у файлі {filehtml} відсутні заголовки в документі." )
        file_content = re.sub('<p>(<b>)?(<i>)?Вечірні молитви(</i>)?(</b>)?</p>','<h3><b><i>Вечірні молитви</i></b></h3>',file_content)
        file_content = re.sub('<p>(<b>)?(<i>)?Велика .ктенія(</i>)?(</b>)?</p>','<h3><b><i>Велика єктенія</i></b></h3>',file_content)
        file_content = re.sub('<p>(<b>)?(<i>)?Мирна .ктенія(</i>)?(</b>)?</p>','<h3><b><i>Велика єктенія</i></b></h3>',file_content)
        file_content = re.sub('<p>(<b>)?(<i>)?Утренні молитви(</i>)?(</b>)?</p>','<h3><b><i>Утренні молитви</i></b></h3>',file_content)
        file_content = re.sub('<p>(<b>)?(<i>)?Канон(</i>)?(</b>)?</p>','<h2><b><i>Канон</i></b></h2>',file_content)
    with open(temppath+'/'+filehtml, 'w', encoding='utf-8') as f:
        f.write(file_content)    
    with open(temppath+'/'+filehtml, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()
    with open(temppath+'/'+filehtml, 'w',encoding='utf-8') as f:
        islist = False
        for line in file_lines:
            if not line.isspace():
                line = re.sub(re.compile('<a id="_.*?</a>'), '', line) # очистка сміття типу <a id="_Hlk12345678"></a>

                if islist:
                    line = line.replace('p>','li>')
                    linestart = line.find('.') + 1
                    line = '<li>' + line[linestart:]
                if line.find('Утренні молитви') != -1 or line.find('Вечірні молитви') != -1:
                    continue
                if line.find('Під час рецитування псалмів') != -1:
                    line = line + '<details><summary>Утренні молитви</summary><ol>'
                    islist = True
                if line.find('Під час 103-го псалма') != -1:
                    line = line + '<details><summary>Вечірні молитви</summary><ol>'
                    islist = True
                if islist and line.find('<h') != -1:
                    line = '</ol></details>\n' + line.replace('<li>', '')    
                    islist = False
                if line.find('Пісня Богородиці') != -1:        
                    line = '<h3 id="magnificat"><b><i>Пісня Богородиці</i></b></h3>'
                if line.find('Канон') != -1 and line.find('<h2>') != -1:
                    line += '<i>(Якщо скорочується Канон, перейди до <a href="#magnificat">Пісня Богородиці</a>).</i>\n'
                
                
                if line.startswith('<li></li>'):
                    print(f'УВАГА: щось пішло не так. Перевірити {filehtml}')
                    f.close()
                    raise Error
                #print(line)
                f.write(line)

print('Формуємо вихідні файли')
for filehtml in html_files:
    with open(temppath+'/'+filehtml, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()

    fileV = None
    fileU = None
    print(filehtml)
    file_index = int(filehtml.split('-')[0].split('.')[0])

    for line in file_lines:

        if line.find('ПЕРЕДШЕОСВЯЧЕНИХ') != -1:
            line = line.replace('p>', 'h1>')
            fileV = open(outpath+'/'+'l{:02d}.html'.format(file_index), 'w',encoding='utf-8') 
            
        if line.find('ВЕЧІРНЯ') != -1:
            line = line.replace('p>', 'h1>')
            fileV = open(outpath+'/'+'t{:02d}v.html'.format(file_index), 'w',encoding='utf-8') 
            
        if line.find('УТРЕНЯ') != -1:
            line = line.replace('p>', 'h1>')
            if fileV:
                fileV.close()
                fileV = None
            fileU = open(outpath+'/'+'t{:02d}u.html'.format(file_index), 'w',encoding='utf-8')

        if fileV:    
            fileV.write(line)            
        if fileU:    
            fileU.write(line)

    if fileV:
        fileV.close()
        fileV = None

    if fileU:
        fileU.close()
        fileU = None 

print('Готово!')

# shutil.rmtree(temppath, ignore_errors=True, onerror=None)

#openpath = os.path.realpath(outpath)
#os.startfile(openpath)
