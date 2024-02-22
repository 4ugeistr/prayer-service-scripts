import re, csv,easygui
from thefuzz import fuzz
from datetime import datetime

#mode='u'
mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1
#month = 'Березень'
month_no = datetime.now().month+1 if datetime.now().month!=12 else 1

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
    

with open('c1_'+mode+'.txt','r',encoding='utf8') as f:
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
    #if k==5:
    #    print(lines_dict[k]['readings'])
    lines_dict[k]['readings']= transform_readings(k,v['readings'])
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
    res = re.findall('<i>(.*?)</i>',v['readings'])
    v['array']+=normalize(res)
        
    res = re.findall(r'Ап\. – (.*?)(?:<br>|<i>|\n|\||<sup>)',v['readings'])
    v['array']+=normalize(res)

    res = re.findall(r'Єв\. – (.*?)(?:<br>|<i>|\n|<sup>|$)',v['readings'])
    v['array']+=normalize(res)

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
    for i in range(2):
        #print(k,"i=",2+i)
        #print(v['array'][2+i])
        found=False
        ratio_max=0
        for item in apostol_list:
            ratio=fuzz.token_sort_ratio(v['array'][2+i],item[0])
            if ratio==100:
                found=True
                v['array'][2+i]=f'#ap{item[2]:0>3}'+f'{item[3]}'
                break
            if ratio>ratio_max:
                ratio_max=ratio
                item_found=item[0]
        if v['array'][2+i] and not found:
            print(k,2+i,v['array'][2+i], item_found)
            print('ratio:',ratio_max)
            #raise Exception
ev_dic = {
    'Йо.':'iv',
    'Лк.':'lk',
    'Мт.':'mt',
    'Мр.':'mr'
    }
for k,v in lines_dict.items():
    for i in range(2):
        found=False
        ratio_max=0
        for item in evanhelie_list:
            ratio=fuzz.token_sort_ratio(v['array'][4+i],item[0])
            if fuzz.token_sort_ratio(v['array'][4+i],item[0])==100:
                found=True
                v['array'][4+i]=f'#{ev_dic[item[1]]}{item[2]:0>3}{item[3]}'
                break
            if ratio>ratio_max:
                ratio_max=ratio
                item_found=item[0]
        if v['array'][4+i] and not found:
            print(k,4+i,v['array'][4+i],item_found)
            print('ratio:',ratio_max)
            #raise Exception

for k,v in lines_dict.items():
    try:
        weekday = datetime(year_no,month_dic[month],k).isoweekday()
    except ValueError:
        print('Achtung')
        print(k,v)
        raise ValueError
    l = [month,f'{k:0>2}.{month_dic[month]:0>2}',weekday, weekday_list[weekday]]
    v['array']=l+[v['header']]+v['array']

#reverse dictionary search
#month_no = list(month_list.keys())[list(month_list.values()).index(month)][0]
    
#for k,v in lines_dict.items():  
#    print(k,v['array'])

with open(f'{year_no}-{month_no}-{mode}.csv','w', newline='', encoding='utf-8') as csvfile:
    csvwriter = csv.writer(csvfile, delimiter=',', quotechar='"')
    for k,v in lines_dict.items():  
        csvwriter.writerow(v['array'])

