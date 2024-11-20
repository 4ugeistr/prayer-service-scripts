import csv
import json

dataset={1:'apostol',
         2:'evanhelie'}
chosen = int(input("1 - apostol, 2 - evanhelie\n") )


csv_file = open(f'{dataset[chosen]}.csv', 'r', encoding='utf-8')
json_file = open(f'{dataset[chosen]}.json', 'w', encoding='utf-8')

fieldnames = ('reading', 'book', 'zach_num', 'zach_abc')
reader = csv.DictReader(csv_file, fieldnames)

# Skip the header row
next(reader)

data = []
for row in reader:
    # Skip empty rows
    if not any(row.values()):
        continue
    del row[None]
    #print(row)2
    #print({})
    data.append(row)

json.dump(data, json_file, ensure_ascii=False)

csv_file.close()
json_file.close()
