import docx
from datetime import datetime

def get_variable_parts_from_template(path):
    #path = find_octoechos_template(echos,day)
    
    doc = docx.Document(path)
    matrix = []
    service = None
    key=None
    for p in doc.paragraphs:
        
        if p.text.lower() == 'вечірня':
            service = 'вечірня'
        if p.text.lower() == 'утреня':
            service = 'утреня'

        if key:
            for item in list_of_lines_that_end_parts:
                if p.text.startswith(item):
                    key= None
                    continue

        #SPAGHETTI CODE. Redo later
        if not key and p.text =='Стихири':
            
            key = 'стихири_ГВ' if service == 'вечірня' else 'стихири_хвалитні'
            matrix.append({'key':key,'text':[p]})
            continue

        if not key and p.text.startswith('Прокімен'):
            key = f'прокімен_{service}'
            matrix.append({'key':key,'text':[p]})
            continue

        if not key and p.text.startswith('Стихири на стиховні'):
            key = f'стиховня_{service}'
            matrix.append({'key':key,'text':[p]})
            continue
        if not key and p.text.startswith('Тропарі'):
            key = f'тропарі_{service}'
            matrix.append({'key':key,'text':[p]})
            continue

        if not key and p.text.startswith('Після Вечірні'):
            key = f'тропар_вкінці_{service}'
            matrix.append({'key':key,'text':[p]})
            continue

        if not key and p.text.startswith('Перший сідальний'):
            key = f'сідальний1'
            matrix.append({'key':key,'text':[p]})
            continue
        
        if not key and p.text.startswith('Другий сідальний'):
            key = f'сідальний2'
            matrix.append({'key':key,'text':[p]})
            continue
        
        if not key and p.text.startswith('Іпакой'):
            key = f'іпакой'
            matrix.append({'key':key,'text':[p]})
            continue

        if not key and p.text.startswith('Степенна пісня'):
            key = f'степенна'
            matrix.append({'key':key,'text':[p]})
            continue
        
        if not key and p.text.startswith('Канон'):
            key = f'канон'
            matrix.append({'key':key,'text':[p]})
            continue


        if not key and p.text.startswith('Пісня 9'):
            key = f'пісня9'
            matrix.append({'key':key,'text':[p]})
            continue

        if not key and p.text.startswith('Світильний'):
            key = f'світильний'
            matrix.append({'key':key,'text':[p]})
            continue
    
        if not key and p.text==('Тропар'):
            key = f'воскресний_тропар'
            matrix.append({'key':key,'text':[p]})
            continue


        #!!!!!!!! Чи це працюватиме, з послідовними змінними частинами???

        

        if key:
            matrix[-1]['text'].append(p)

    return matrix

list_of_lines_that_end_parts = ['Світло тихе',
                                'Вхід', 
                                'Сподоби, Господи', #кінець прокімена седмичного дня
                                'Господеві помолімся.', #кінець прокімена неділі
                                'Пісня Симеона', #кінець стиховні вечірні
                                "Великий відпуст", #кінець тропарів вечірні
                                "Утреня", #кінець вечірні
                                "УТРЕНЯ", #кінець вечірні
                                "Тропарі", #кінець початку утрені
                                "По тропарях читає читець чергову катизму, після якої диякон перед св. дверми виголошує:", #кінець тропарів 
                                'Відтак читець читає другу чергову катизму, диякон виголошує другу малу єктенію:',
                                'Полієлей', #кінець другого сідального
                                'Степенна', #кінець іпакоя
                                'Євангеліє', #кінець степенної
                                'Господеві помолімся.', #кінець прокімна
                                #'Мала єктенія',
                                'Псалом 50',
                                'Пісня Богородиці',#кінець канону 1-8
                                'Пісня 9',#кінець Пісні Богородиці
                                'Достойно',
                                '«Свят Господь Бог',
                                'Псалом 148', #кінець світильного
                                'Хвалитні', #кінець світильного
                                "Велике славослов'я",  #кінець хвалитних стихир
                                'Після стиховні читаємо:', #кінець стиховні утрені седмичного дня
                                "Тропар", #кінець Великого славослів'я
                                "Єктенія усильного благання", #кінець воскресного тропаря, кінець веірнього прокімена
                                

]



if __name__ == "__main__":
    start_time = datetime.now()
    
    print("Hello World!")
    path = r"drafts\\07_vu\\15-ВТ-Володимира_Великого.docx"
    doc = docx.Document(path)
    matrix = get_variable_parts_from_template(path)
    

    for item in matrix:
        print(item['key'])

    end_time = datetime.now()
    elapsed_time = (end_time - start_time).total_seconds()
