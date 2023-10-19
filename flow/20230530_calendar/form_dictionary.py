import csv,re

f = open("calendar23.txt", 'r',encoding='utf8')
csvfile = open("holiday_dictionary.csv",'w',newline='',encoding='utf8')
spamwriter=csv.writer(csvfile,delimiter=';',quotechar='"', quoting=csv.QUOTE_MINIMAL)

lines = f.readlines()
dic=[]
for line in lines:
    re_result=re.search("^.{8}\|.{3}\|.\|.\|.\|.\|.\|(.*?)\|",line)
    if re_result:
        day_title=re_result.group(1)
    day_title=day_title.replace("<b>","").replace("</b>","")
    re_result=re.findall("(\{)?<(strong|span|i|em)>(.*?)</(strong|span|i|em)>",day_title)
    if re_result:
        for r in re_result:
            if r[1]!=r[3]:
                print("Warning:", r[1], r[3], line)
            
            if r[2][0]=="{":
                dic.append([r[2][1:-1],r[1],""])
            elif r[2][0] in "@+#*":
                dic.append([r[2][1:].strip(),r[1],r[2][0]])
            else:
                dic.append([r[2],r[1],""])
            
dic_unique = {d[0]:d[1:] for d in dic }
dic_sorted = dict(sorted(dic_unique.items()))
for k, v in dic_sorted.items():
    spamwriter.writerow([k]+v)
f.close()
csvfile.close()
                         
        


    
