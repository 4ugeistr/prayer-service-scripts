import csv, re, calendar, unicodedata
from datetime import datetime


month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year

'''
mode = full | short

 
'''
def get_saints(month,day,mode='full',multiline=False):
    index = 9 if mode == 'full' else 8
    separator = '\n' if multiline else ', '
    saint_string=""
    for row in saint_matrix[1:]:
        if int(row[0])==month and int(row[1])==day:
            if saint_string:
                saint_string+=separator+row[index][0].lower()+row[index][1:]
            else:
                saint_string=row[index]
    return saint_string

def get_saints_for_filename(month,day,multiline=False):
    index = 10
    separator = ','
    saint_string=""
    for row in saint_matrix[1:]:
        if int(row[0])==month and int(row[1])==day:
            if row[index]:
                if saint_string:
                    saint_string+=separator+row[index]
                else:
                    saint_string=row[index]
    return saint_string


def get_matrix_full(csv_filename):
    matrix=[]
    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            matrix.append(row)
    return matrix

#готуємо для відпусту
#прибираємо не-чар символи, робим перший символ lowercase()
def transform_line(line):
    # Remove non-character symbols at the beginning
    processed_line = line.strip()
    for char in processed_line:
        if unicodedata.category(char)[0] != 'L':
            processed_line = processed_line.lstrip(char)
        else:
            break
    # Make the first character lowercase
    return processed_line[0].lower() + processed_line[1:]
    

saint_matrix = get_matrix_full("Місяцеслов-БД.csv")

lines = []
for day in range(1,calendar.monthrange(year_no, month_no)[1]+1):
    lines.append(transform_line(get_saints(month_no,day)))
    print(day, get_saints(month_no,day))

lines = '\n'.join(lines)

with open('output.txt','w',encoding='utf8') as file:
    file.write(lines)

