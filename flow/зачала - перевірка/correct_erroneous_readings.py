import csv,glob,re

filenames = glob.glob('*.txt')+glob.glob('*/*/*.txt')

correction_list=[]
with open('reading_mistakes.csv','r',encoding='utf8') as csv_file:
    spamreader = csv.reader(csv_file, delimiter=';', quotechar='"')
    for line in spamreader:
        correction_list.append(line)
    ccrrection_list = correction_list[1:]

for file_path in filenames:
    with open(file_path,'r',encoding='utf8') as file:
        content= file.read()
        for item in correction_list:
            if item[0]=='1 Кор. 143 зач. (від полов.); 10, 1-4.':
                if re.search(item[0].replace('.','\.'),content):
                    print('DING')
            content = re.subn('(Ап|Єв)\. – '+re.escape(item[0])+'(<br>|<i>|\n|\|)','\g<1>. – '+item[1]+'\g<2>',content)[0]
    with open(file_path,'w',encoding='utf8') as file:
        file.write(content)
    print(f'{file_path} - DONE!')
            
        

    
