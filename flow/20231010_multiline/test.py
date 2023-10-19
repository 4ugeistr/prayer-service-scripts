import csv

csv_filename='МЦ Відпусти НЮ.csv'
reading_matrix=[]
with open(csv_filename, newline='', encoding='utf-8') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
    for row in spamreader:
        reading_matrix.append(row)


'''
Priority
1) Special dismissals
2) Replacement dismsissals = day template with a couple of replaced parts
2a) polyeleos
2b) 
3) Regular



def generate_dismissal(date):
    if datetime(year_no, month_no, d).weekday()+1==7
