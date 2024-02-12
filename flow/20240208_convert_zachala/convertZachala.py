import os, mammoth, easygui,re
from glob import glob

custom_style_map = """
i => i
b => b
u => em
"""



def insert_line_breaks(html):
    # Add line breaks after block-level elements
    block_level_elements = ['</h1>', '</h2>', '</h3>', '</h4>', '</h5>', '</h6>', '</p>', '</ol>', '</ul>']
    for element in block_level_elements:
        html = html.replace(element, element + '\n')
    return html


def main():
    dirs = easygui.diropenbox()
    print(f'Converting docs in dir: {dirs}')
    os.chdir(dirs)
    filenames = glob(f'*.docx')
    if not os.path.exists('output'):
        os.mkdir('output')

    for filename in filenames:
        with open(filename,'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file,style_map=custom_style_map,include_default_style_map=False ).value
            #result = mammoth.convert_to_html(docx_file).value

        result = insert_line_breaks(result)
        #change &nbsp to regular space
        result=result.replace(u'\xa0', ' ')

        #mammoth doesn't process superscript. workaround
        result = re.sub('(\d{1,2})\. ','<i>\g<1>.</i> ',result)
        #for some reason dots are out side italics run. getting them in        
        result = re.sub('<i>(\d{1,2})</i>. ','<i>\g<1>.</i> ',result)
        
        html_filename=f'output/'+filename.split('.')[0]+'.html'
        with open(html_filename,'w',encoding='utf8') as html_file:
            html_file.write(result)

if __name__ == "__main__":
    main()
