import re, csv,easygui, copy
from thefuzz import fuzz
from datetime import datetime
#import paschalia

#mode='u'
mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1
#month = 'Березень'

month_no = datetime.now().month+1 if datetime.now().month!=12 else 1

'''
month_no=5
year_no=2026
print(f"WARNING: MONTH OVERRIDE!")
'''
print(f"Processing month: {year_no}-{month_no}")


month_dic= {'Січень':1,
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
month_dic_reversed = {v:k for k,v in month_dic.items()}
month = month_dic_reversed[month_no]

weekday_list = {
                1:'пн',
                2:'вв',
                3:'ср',
                4:'чт',
                5:'пт',
                6:'сб',
                7:'нд'}

ev_dic = {
    'Йо.':'iv',
    'Ів.':'iv',
    'Лк.':'lk',
    'Мт.':'mt',
    'Мр.':'mr'
    }

def clean_half_words(s):
    s=s.replace(' (від половини)','').replace(' (від полов.)','')
    return s

def clean_tags(s):
    symbol_dict={
    '#':'🕀',
    '*':'🕁',
    '+':'🕂',
    '@':'🕃', #(red)
    '&':'🕃'} #(black)
    s=re.sub('#','🕀 ',s)
    s=re.sub(r'\*','🕁 ',s)
    s=re.sub(r'\+','🕂 ',s)
    s=re.sub('@','🕃 ',s)
    s=re.sub('&','🕃 ',s)
    
    s=re.sub('<span>','',s)
    s=re.sub('</span><br>',' ',s)
    s=re.sub('</span>','',s)
    s=re.sub('<strong>','',s)
    s=re.sub('</strong>',' ',s)
    s=re.sub('<i>','',s)
    s=re.sub('</i><br>',' ',s)
    s=re.sub('<em>',' ',s)
    s=re.sub('</em><br>',' ',s)
    s=re.sub('<b>','',s)
    s=re.sub('</b><br>',' ',s)
    s=re.sub('<br>',' ',s)
    s=re.sub('  ',' ',s)
    s=re.sub('^ ','',s)
    return s

def restore_format(lines):
    new_lines = [re.sub(r'(\d{2}\|\d\|\d\|\d\|\d\|.)(.*)',r'\g<1>|\g<>',line) for line in lines]
    return new_lines

def transform_readings(k,s):
    #print(k,'before=',s)
    
    s = re.sub('Літ:','Літ.:',s)
    s = re.sub(r'<i>Літургія св. Йоана Золотоустого\.</i><br>','',s)
    s = re.sub(r'Літургія св. Василія Великого\.','Літ.',s)
    s = re.sub(r'Літургія св. Василія Великого\.','Літ.',s)
#    s = re.sub('<i>Вечірня з Літургією св. Василія Великого.</i><br>','',s)
    '''
    if k==26:
        print(s)
    '''
    s = re.sub(r'Літ\.(?:[^:])','Літ.:<',s)
    
    
    s = re.sub(r'<i>Утр(\.)?:</i>.*?<i>Літ(\.)?(:)?<br>(?P<reading>.*)',r'<i>\g<reading>',s)
    s = re.sub(r'<i>Утр(\.)?:</i>.*?<i>Літ(\.)?(:)?</i><br>(?P<reading>.*)',r'\g<reading>',s)
    s = re.sub(r'<i>Утр(\.)?:</i><br>Єв. – (?:.*?)<br>(?P<reading>.*)',r'\g<reading>',s)
    s = re.sub(r'<i>Літ(\.)?(:)?</i><br>(?P<reading>.*)',r'\g<reading>',s)
    #замінити <i>Ряд.:</i>
    #на <i>Літ.:<br>Ряд.:</i>
    
    if re.search('<i>Утр.:</i>',s) and not re.search('Літ',s):
        s=""
    
    if re.search('<i>Утр.:</i>',s) or not s:
        print('Щось не так:',k,s)
    #print(k,'before=',s)
    return s
    
print(f"Укладаємо коди читань для {year_no}-{month_no}")
print(f"Обробляємо файл: c1_{mode}.txt")

with open(f'c1_{mode}.txt','r',encoding='utf8') as f:
    lines = f.readlines()
#lines = restore_format(lines)
lines = [re.sub(r'(\d{2}\|\d\|\d\|\d\|\d\|)(.)(.*)',r'\g<1>\g<2>|\g<3>',line) for line in lines]

lines_dict={}
for line in lines:
    res=re.search(r'(\d{2})\|\d\|\d\|\d\|\d\|.\|(.*)\|(.*)\|.*',line)
    if res and res.group():
        lines_dict[int(res.group(1))] = {
            'header':clean_tags(res.group(2)),
            'readings':res.group(3)
            }
        
        #if int(res.group(1))==5:
        #    print('ініт',int(res.group(1)), lines_dict[int(res.group(1))])
        
    else:
        raise Exception



for k,v in lines_dict.items():
    #if k in (17,20):
    #print(k, lines_dict[k]['readings'])
    lines_dict[k]['readings']= transform_readings(k,v['readings'])
    #print(k,lines_dict[k]['readings'])
    #if k==5:
    #    print(lines_dict[k]['readings'])
    
#зробити масив сталого розміру з пустими строками за відсутніх значень    
def normalize(arr, n=2):
    if not arr:
        arr=[]
    if len(arr)<n:
        arr+=['']*(n-len(arr))
    return arr








for k,v in lines_dict.items():
    #print(k,v['readings'])
    v['array']=[]

    list_of_headings_to_ignore = [

        'Вечірня з Літургією св.Василія Великого.',
        'На вмиванні:',
        'По вмиванні:',
        'Царські часи.',
        'На 1 - му часі:',
        'На 3 - му часі:',
        'На 6 - му часі:',
        'Літургія Передосвячених Дарів:',
        'Літургія св. Івана Золотоустого.',
        'Літургія св.Івана Золотоустого.',
        'Час Шостий:',
        'Вечірня:',
        'На 9 - му часі:',
        'Єрусалимська Утреня:',
    ]

    res = re.findall('<i>(.*?)</i>',v['readings'])
    res2 = res.copy()

    for r in res:
        print(r)
        if r in list_of_headings_to_ignore:
            #print(f"Removing {r} from {res2}")
            res2.remove(r)
            #print(f"Removed: {res2}")

    v['array']+=normalize(res2)
        
    res = re.findall(r'Ап\. (?:–|-) (.*?)(?:<br>|<i>|\n|\||<sup>)',v['readings'])
    v['array']+=normalize(res)

    res = re.findall(r'Єв\. (?:–|-) (.*?)(?:<br>|<i>|\n|<sup>|$)',v['readings'])
    v['array']+=normalize(res)
    v['array_transformed']=copy.copy(v['array'])+[""]*4


#convert apostol
apostol_list=[]
with open('apostol.csv', newline='', encoding='utf-8') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',', quotechar='"')
    for row in csvreader:
        apostol_list+=[row]
evanhelie_list=[]
with open('evanhelie.csv', newline='', encoding='utf-8') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',', quotechar='"')
    for row in csvreader:
        evanhelie_list+=[row]

#for k,v in lines_dict.items():  
#    print(k,v['array'])


for k,v in lines_dict.items():
    #print(k,v)
    v['array']=[clean_half_words(item) for item in v['array']]
    #print(k,v)



for k,v in lines_dict.items():
    for i in range(2):
        #print(k,"i=",2+i)
        #print(v['array'][2+i])
        found=False
        ratio_max=0
        for item in apostol_list:
            ratio=fuzz.token_sort_ratio(v['array'][2+i],item[0])
            if ratio==100:
                found=True
                v['array_transformed'][2+i]=f'#ap{item[2]:0>3}'+f'{item[3]}'
                v['array_transformed'][2+4+i]=item[4]
                ratio_max=ratio
                break  
            if ratio>ratio_max:
                ratio_max=ratio
                item_found=item
        
        if ratio_max>=85 and ratio_max<100:
            found=True
            v['array_transformed'][2+i]=f'#ap{item_found[2]:0>3}{item_found[3]}'
            v['array_transformed'][2+4+i]=item[4]
            print(f"WARNING. ratio: {ratio_max}")
            print(k,2+i,v['array'][2+i], item_found[0])
            print('Input:',lines_dict[k]['readings'])  
        
        if v['array_transformed'][2+i] and not found:
            print(f"ERROR. ratio: {ratio_max}")
            print(k,2+i,v['array'][2+i], item_found[0])
            print('ratio:',ratio_max)
            #raise Exception


for k,v in lines_dict.items():
    for i in range(2):
        found=False
        ratio_max=0
        for item in evanhelie_list:
            ratio=fuzz.token_sort_ratio(v['array'][4+i],item[0])
            #print(f"Checking {v['array'][4+i]} vs {item[0]}, ratio = {ratio}")
            if ratio==100:
                found=True
                v['array_transformed'][4+i]=f'#{ev_dic[item[1]]}{item[2]:0>3}{item[3]}'
                v['array_transformed'][4+4+i]=item[4]
                ratio_max=ratio
                break
            if ratio>ratio_max:
                ratio_max=ratio
                item_found=item
        if ratio_max>=85 and ratio_max<100:
            found=True
            v['array_transformed'][4+i]=f'#{ev_dic[item_found[1]]}{item_found[2]:0>3}{item_found[3]}'
            v['array_transformed'][4+4+i]=item[4]
            print(f"WARNING. ratio: {ratio_max}")
            print(k,4+i,v['array'][4+i],item_found[0])
            print('Input:',lines_dict[k]['readings'])  
        
        if v['array'][4+i] and not found:
            print(f"ERROR. ratio: {ratio_max}")
            print(k,4+i,v['array'][4+i],item_found[0])
            print('ratio:',ratio_max)
            #raise Exception

for k,v in lines_dict.items():
    try:
        weekday = datetime(year_no,month_dic[month],k).isoweekday()
    except ValueError:
        print(f'Achtung: cannot resolve date with Y:{year_no} M:{month_dic[month]} D:{k}')
        print(k,v)
        raise ValueError
    l = [month,f'{k:0>2}.{month_dic[month]:0>2}',weekday, weekday_list[weekday]]
    v['array_transformed']=l+[v['header']]+v['array_transformed']

#reverse dictionary search
#month_no = list(month_list.keys())[list(month_list.values()).index(month)][0]
    
#for k,v in lines_dict.items():  
#    print(k,v['array'])

with open(f'csv\\{year_no}-{month_no}-{mode}.csv','w', newline='', encoding='utf-8') as csvfile:
    csvwriter = csv.writer(csvfile, delimiter=',', quotechar='"')
    for k,v in lines_dict.items():  
        csvwriter.writerow(v['array_transformed'])

validation_matrix=[]
for k,item in lines_dict.items():
    #line = item['array_transformed']
    index_list = [0,2,1,3]
    for i in index_list:
        validation_line = item['array_transformed'][0:5]
        validation_line.append(item['array'][i%2])
        validation_line.append(item['array'][2+i])
        validation_line.append(item['array_transformed'][7+i])
        validation_line.append(item['array_transformed'][7+4+i])
        if validation_line[6]:
            validation_matrix.append(validation_line)

with open(f'csv\\{year_no}-{month_no}-{mode}_VAL.csv','w', newline='', encoding='utf-8') as csvfile:
    csvwriter = csv.writer(csvfile, delimiter=',', quotechar='"')
    for row in validation_matrix:  
        csvwriter.writerow(row)

print("Done!")
