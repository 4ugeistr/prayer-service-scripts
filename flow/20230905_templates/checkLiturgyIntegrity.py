import docx,re

docx_filename='12-Літургія-НЮ.docx'
mode='u'
cur_month='Грудень'


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

month_list_string='('+'|'.join(month_list.keys())+')'
day_list_string="(Понеділок|Вівторок|Середа|Четвер|П’ятниця|П'ятниця|Субота|Неділя)"

list_of_text_templates=['liz','lvv','liz_pascha']

list_of_tags_mandatory=['antifon1',
                        'antifon2',
                        'antifon3',
                        'vhidne',
                        'tropari',
                        'prokimen',
                        'apostol',
                        'aleluia',
                        'evanhelie',
                        'prychasnii',
                        'vidpust']
list_of_tags_optional=['trysviate',
                       'dostoino']



doc = docx.Document(docx_filename)
history={}

def pretty(d, indent=0):
   for key, value in d.items():
      print('\t' * indent + str(key))
      if isinstance(value, dict):
         pretty(value, indent+1)
      else:
         print('\t' * (indent+1) + str(value))
      
#for i in range(100):
#   print(doc.paragraphs[i].text)


#прибираємо leading / trailing пробіли
for p in doc.paragraphs:
   if re.search('^.*?\s+$',p.text):
      p.text = re.sub('^(.*?)\s+$','\g<1>',p.text)
   if re.search('^\s+.*?$',p.text):
      p.text = re.sub('^\s+(.*?)$','\g<1>',p.text)

#for i in range(100):
#   print(doc.paragraphs[i].text)
cur_date=None
for p in doc.paragraphs:
    if mode == 'u':
       re_result=re.search(f'^{month_list_string} (\d+)',p.text)
       if re_result:
           cur_date = int(re_result.group(2))
    elif mode == 'g':
       #re_result=re.search(f'^\n*({day_list_string}\n)?(\d+)',p.text)
       re_result=re.search(f'^(\d+)',p.text)
       #print(re_result, p.text[:20])
       #print('Y')
       if re_result:
           cur_date = int(re_result.group(1))

    if cur_date and not cur_date in history:
       history[cur_date]={}

    
    if cur_date:
        ustav = re.search('^<ustav.*?liturgia=(\w+).*$', p.text)
        if ustav:
            if ustav.group(1) in list_of_text_templates:
                history[cur_date]['ustav']='OK'
            else:
                history[cur_date]['ustav']=ustav.group(1)
                
        for item in list_of_tags_mandatory+list_of_tags_optional:
            #if item=='vidpust':
            #    print(p.text)
            #if re.search('^\s?<'+item+'>\s?$',p.text):
            if re.search('^<'+item+'>$',p.text):
               if not item in history[cur_date]:
                  history[cur_date][item]='<'
               else:
                  history[cur_date][item]+='<'
            #if re.search('^\s?</'+item+'>\s?$',p.text):
            if re.search('^</'+item+'>$',p.text):
               if not item in history[cur_date]:
                  history[cur_date][item]='>'
               else:
                  history[cur_date][item]+='>'

print(f'Файл: {docx_filename}\n')

optional_tag_summary={}
for k, day in history.items():
   #перевірка правильності <ustav>
   if 'ustav' in day:
      if day['ustav']!='OK':
         print(f'День {k}: <ustav>, неправильний шаблон ('+day['ustav']+')')

      #перевірка правильності основного набору міток
      for item in list_of_tags_mandatory:
         if not item in day:
            print(f'День {k}: пропущено секцію <{item}>')
         else:
            if not '<' in day[item]:
               #print(item, day[item])
               print(f'День {k}: пропущено відкриваючий <{item}>')
            if not '>' in day[item]:
               print(item, day[item])
               print(f'День {k}: пропущено закриваючий </{item}>')
                  
      #підсумки по опціональному набору міток
      for item in list_of_tags_optional:
         if item in day:
            #print('Found optional:',k,item)
            if day[item]!='<>':
               if not '<' in day[item]:
                  print(f'День {k}: пропущено відкриваючий <{item}>')
               if not '>' in day[item]:
                  print(f'День {k}: пропущено закриваючий </{item}>')
            else:
               if not item in optional_tag_summary:
                  optional_tag_summary[item]=[k]
               else:
                  optional_tag_summary[item]+=[k]

      
   else:
      print(f'День {k}: немає Літургії або відсутній <ustav>')
      


print("\nДні використання опціональних рубрик:")
pretty(optional_tag_summary)
            

            
            
      
         


                




