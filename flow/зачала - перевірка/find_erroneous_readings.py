import re,glob
import readings_widget

filenames= glob.glob('calendar*.txt')#+glob.glob('*/*/*.txt')

fa = open('all_apostol.txt','w',encoding='utf-8')
fe = open('all_evanhelie.txt','w',encoding='utf-8')

for filename in filenames:
    print(f'\nФАЙЛ: {filename}')
    pass

    file = open(filename,'r',encoding = 'utf8')
    lines=file.readlines()
    file.close()


    lines_dict = {}
    i=0
    lines_dict = {i+1:lines[i] for i in range(len(lines))}

    apostol_source_list='(Гал\.|Ді\.|Еф\.|Євр\.|1 Йо\.|2 Йо\.|3 Йо\.|Кол\.|1 Кор\.|2 Кор\.|1 Пт\.|2 Пт\.|Рим\.|1 Сол\.|2 Сол\.|1 Тим\.|2 Тим\.|Тит\.|Флм\.|Флп\.|Юд\.|Як\.)'

    for k,v in lines_dict.items():
        #print(v)
        #print(re.search('^.*?|.*?|.|.|.|.|.|.*?|(.*?)|.*?$',v))
        #if re.search('^.*?|.*?|.|.|.|.|.|.*?|(.*?)|.$',v):
        readings_widget.find_reading_filename(re.sub('^\d{8}\|.{3}\|\d\|\d\|\d\|\d\|.\|.*\|(.*)\|.*$','\g<1>',v))
        
    for k,v in lines_dict.items():
        #if len(re.findall('(Ап\. – .*?)(<br>|<i>|\n|\|)',v))>2:
        #    print(len(re.findall('(Ап\. – .*?)(<br>|<i>|\n|\|)',v)))
        for item in re.findall('(Ап\. – .*?)(<br>|<i>|\n|\|)',v):
            
            fa.write(item[0][6:]+'\n')

            pattern ='(\d+, \d+|\d+, \d+-\d+|\d+-\d+|\d+, \d+ – \d+, \d+|\d+ – \d+, \d+|\d+)'
            middle_part ='('+pattern+'; )*?'
            last_part = pattern+'\.'
            if not re.search(f'Ап\. – {apostol_source_list} \d+ зач\.( \(від половини\))?; '+middle_part+last_part ,item[0]):
                print(f'{item[0][6:]}')

    for k,v in lines_dict.items():
        for item in re.findall('(Єв\. – .*?)(<br>|<i>|\n|\|)',v):
            fe.write(item[0][6:]+'\n')
            if not re.search('Єв\. – (Мт\.|Лк\.|Мр\.|Йо\.) \d+ зач\.( \(від половини\))?;',item[0]):
                print(f'{item[0][6:]}')
        
fa.close()
fe.close()
