import os,re,mammoth,easygui
from glob import glob

custom_style_map = """
i => i
b => b
u => em
"""


header=[
    '<h2><b>ЧАС ПЕРШИЙ</b></h2>',
    '<h2><b>ЧАС ТРЕТІЙ</b></h2>',
    '<h2><b>ЧАС ШОСТИЙ</b></h2>',
    '<h2><b>ЧАС ДЕВ\'ЯТИЙ</b></h2>',
    '<h2><b>ЧАС ДЕВ’ЯТИЙ</b></h2>',
    '<h2 align="center" ><a id=\'t1\'><b>ЧАС ПЕРШИЙ</b></a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a href=\'#t9\'>[9]</a></h2>',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a id=\'t3\'><b>ЧАС ТРЕТІЙ</b></a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a href=\'#t9\'>[9]</a></h2>',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a id=\'t6\'><b>ЧАС ШОСТИЙ</b></a> &middot; <a href=\'#t9\'>[9]</a></h2>',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a id=\'t9\'><b>ЧАС ДЕВ\'ЯТИЙ</b></a></h2>',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a id=\'t9\'><b>ЧАС ДЕВ\'ЯТИЙ</b></a></h2>'
    ]

dirs = easygui.diropenbox()
print(f'Converting: {dirs}')
os.chdir(dirs)

filenames = glob(f'*.docx')
if not os.path.exists('output'):
    os.mkdir('output')
    
for filename in filenames:
    with open(filename,'rb') as docx_file:
        result = mammoth.convert_to_html(docx_file,style_map=custom_style_map).value
    n=int(len(header)/2)
    for j in range(n):
        result = re.sub(header[j],header[j+n],result)

    html_filename=f'output/t'+filename.split('.')[0]+'c.html'
    with open(html_filename,'w',encoding='utf8') as html_file:
        html_file.write(result)

print("Done!")
