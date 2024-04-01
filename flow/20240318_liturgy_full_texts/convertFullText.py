import os, mammoth, easygui,re
from glob import glob

custom_style_map = """
i => i
b => b
u => em
"""

heading_strings = [
    'Мирна єктенія',
'Молитва першого антифону',
'Мала єктенія',
'Молитва другого антифону',
'«Єдинородний Сину»',
'Мала єктенія',
'Молитва третього антифону',
'Молитва входу',
'Молитва Трисвятої пісні',
'Прокімен',
'Апостол',
'Алилуя',
'Молитва перед Євангелієм',
'Євангеліє',
'Єктенія усильного благання',
'Молитва за усопших',
'Єктенія за вірних',
'Перша молитва за вірних',
'Друга молитва за вірних',
'Херувимська пісня',
'Молитва Херувимської пісні',
'Великий вхід із Чесними Дарами',
'Єктенія за принесені Чесні Дари',
'Молитва приношення',
'Символ віри',
'Анафора',
'Єктенія за освячені Чесні Дари',
'Молитва',
'«Отче наш»',
'Причастя',
'Молитва приготування до святого причастя',
'Благодарна єктенія',
'Заамвонна молитва',
'Молитва на споживання Святих Дарів',
'Відпуст',
]

heading_strings2 = []
for i in heading_strings:
    heading_strings2.append({"text":i,"found":False})



def insert_line_breaks(html):
    # Add line breaks after block-level elements
    block_level_elements = ['</h1>', '</h2>', '</h3>', '</h4>', '</h5>', '</h6>', '</p>', '</ol>', '</ul>']
    for element in block_level_elements:
        html = html.replace(element, element + '\n')
    return html 


def convert_a_to_h(filename):
    with open(filename,'r',encoding='utf8') as html_file:
        lines = html_file.readlines()
    with open(filename,'w',encoding='utf8') as html_file:
        lvl = 1
        for line in lines:
            
            line = re.sub('^<p><a.*></a><i>(.*)</i></p>$',f'<h{lvl}>\g<1></h{lvl}>',line)    
            
            if lvl == 1:
                line = re.sub('^<p><a.*></a>(.*)</p>$',f'<h{lvl}>\g<1></h{lvl}>',line)    
                lvl = 4
            html_file.write(line)



def main():
    '''
    dirs = easygui.diropenbox()
    print(f'Converting docs in dir: {dirs}')
    os.chdir(dirs)
    filenames = glob(f'*.docx')
    '''
    docx_filename = easygui.fileopenbox(
        title="Select a .docx file",
        filetypes=["*.docx"],
        default="*.docx"
    )
    '''
    if not os.path.exists('output'):
        os.mkdir('output')
    '''
    
    with open(docx_filename,'rb') as docx_file:
        result = mammoth.convert_to_html(docx_file,style_map=custom_style_map,include_default_style_map=False ).value
        #result = mammoth.convert_to_html(docx_file).value

    result = insert_line_breaks(result)
    #change &nbsp to regular space
    result=result.replace(u'\xa0', ' ')

    filename = docx_filename.split('\\')[-1].split('.')[0]
    html_filename=filename + '.html'
    with open(html_filename,'w',encoding='utf8') as html_file:
        html_file.write(result)
    

    convert_a_to_h(html_filename)


    print(f"Finished converting {filename}.docx")

if __name__ == "__main__":
    main()
