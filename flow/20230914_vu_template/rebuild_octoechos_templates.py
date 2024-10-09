import docx,os
import ps_docx_utils as pdu
import ps_date_utils as pdt


everyday_part_list=['вечірня',
                'господи_взиваю',
                '4_перші_стихи',
                'вхід_седмиця',
                'сподоби_господи',
                'нині_відпускаєш',
                'середній_відпуст',
                'утреня',
                'тропарі_седмиця',
                'мала_єктенія_перед_другим_сідальним',
                'псалом_50',
                'молитва_9',
                'канон',
                'достойно',
                'мала_єктенія_після_канона',
                'хвалитні',
                'решта_стихів_хвалитних',
                'прохальна_єктенія_утрені',
                'після_стиховні',
                'потрійна_єктенія_утрені',
                'середній_відпуст',
]
sunday_part_list=['вечірня',
                'господи_взиваю',
                '4_перші_стихи',
                'вхід_седмиця',
                'сподоби_господи',
                'нині_відпускаєш',
                'середній_відпуст',
                'утреня',
                'тропарі_седмиця',
                'мала_єктенія_перед_другим_сідальним',
                'псалом_50',
                'молитва_9',
                'канон',
                'достойно',
                'мала_єктенія_після_канона',
                'хвалитні',
                'решта_стихів_хвалитних',
                'прохальна_єктенія_утрені',
                'після_стиховні',
                'потрійна_єктенія_утрені',
                'середній_відпуст',
]

octoechos_template_source_folder = '01-Октоїх'


def find_octoechos_template(day,echos, folder=octoechos_template_source_folder):

    #tmp
    hyphen = '' if day==7 else '-'

    path = folder+f'\\Глас_{echos}\\{echos}-{pdt.day_short_dic_reversed[day]}{hyphen}.docx'
    print(path)
    if os.path.exists(path):
        return path
    else:
        return -1


list_of_lines_that_end_parts = ['Світло тихе',
                                'Вхід',
                                'Сподоби, Господи',
                                'Господеві помолімся.',
]


def get_vu_octoechos_variable_parts_from_template(day,echos):
    path = find_octoechos_template(day, echos)
    doc = docx.Document(path)
    matrix = []
    service = None
    key=None
    for p in doc.paragraphs:
        
        if p.text.lower() == 'вечірня':
            service = 'вечірня'
        if p.text.lower() == 'утреня':
            service = 'утреня'

        if not key and p.text =='Стихири':
            key = 'стихири_ГВ'
            matrix.append({'key':key,'text':[p]})
            continue

        if not key and p.text.startswith('Прокімен'):
            key = f'прокімен_{service}'
            matrix.append({'key':key,'text':[p]})
            continue


        if key and (p.text == 'Світло тихе' or p.text == 'Вхід'):
            key = None

        if key:
            matrix[-1]['text'].append(p)

    return matrix        

def get_vu_misc_variable_parts(path):
    doc = docx.Document(path)
    matrix = []
    key=key2=None


    for p in doc.paragraphs:
        if p.style.name == 'Heading 1':
            key = p.text
            continue
        if p.style.name == 'Heading 2':
            key2 = p.text
            matrix.append({'key':key,'key2':key2,'text':[]})
            continue
    
        matrix[-1]['text'].append(p)
    return matrix


def get_vu_template_parts(path):
    doc = docx.Document(path)
    matrix = []
    key=None
    for p in doc.paragraphs:
        if '#' in p.text:
            if p.text[1:]=='кінець':
                break
            
            key = p.text[1:]
            matrix.append({'key':key, 'text':[] })
            continue
        
        else:
            matrix[-1]['text'].append(p)
    return matrix

def get_template_part_text(matrix, key):
    for item in matrix:
        if item['key']==key:
            return item['text']
    print(f"Warning: {key} not found")
    return -1

def build_template(path,day,echos):
    doc = docx.Document()
    partlist = sunday_part_list if day==7 else everyday_part_list
    for item in partlist:
        pdu.copy_paragraph_list(doc, get_template_part_text(vu_template_parts, item))
    doc.save(path)

vu_template_parts = get_vu_template_parts('vu_template_parts.docx')
vu_misc_variable_parts = get_vu_misc_variable_parts('vu_template_misc_variable_parts.docx')

if __name__ == "__main__":
    build_template('test.docx', 1,1)