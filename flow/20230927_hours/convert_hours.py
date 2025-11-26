import os,re,mammoth,easygui
from glob import glob

custom_style_map = """
i => i
b => b
u => em
p[style-name='Heading 1'] => h1:separator('\n')
p[style-name='Heading 2'] => h2:separator('\n')
p => p:separator('\n')
"""


header=[
    '<h\d>(<b>)?ЧАС ПЕРШИЙ(</b>)?</h\d>',
    '<h\d>(<b>)?ЧАС ТРЕТІЙ(</b>)?</h\d>',
    '<h\d>(<b>)?ЧАС ШОСТИЙ(</b>)?</h\d>',
    '<h\d>(<b>)?ЧАС ДЕВ\'ЯТИЙ(</b>)?</h\d>',
    '<h\d>(<b>)?ЧАС ДЕВ’ЯТИЙ(</b>)?</h\d>',
    '<h2 align="center" ><a id=\'t1\'><b>ЧАС ПЕРШИЙ</b></a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a href=\'#t9\'>[9]</a></h2>\n',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a id=\'t3\'><b>ЧАС ТРЕТІЙ</b></a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a href=\'#t9\'>[9]</a></h2>\n',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a id=\'t6\'><b>ЧАС ШОСТИЙ</b></a> &middot; <a href=\'#t9\'>[9]</a></h2>\n',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a id=\'t9\'><b>ЧАС ДЕВ\'ЯТИЙ</b></a></h2>\n',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a id=\'t9\'><b>ЧАС ДЕВ\'ЯТИЙ</b></a></h2>\n'
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
    result = result.replace('</p>','</p>\n')

    # очистка сміття типу <a id="_Hlk12345678"></a>
    result = re.sub('<a id="_.*?</a>', '', result)

    html_filename=f'output/t'+filename.split('.')[0]+'c.html'
    with open(html_filename,'w',encoding='utf8',newline='\n') as html_file:
        html_file.write(result)

print("Done!")
