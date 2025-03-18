import glob,docx,re

files = glob.glob('*.docx')
for f in files:
    doc = docx.Document(f)
    i=0
    for p in doc.paragraphs:
        found = False
        for r in p.runs:
            if "Спаси́, Бо́же, наро́д Твій" in r.text or "Спаси, Боже, народ Твій" in r.text:
                found = True
                #print(f, ": є Єктенія")
            if "(ім’я, що його́ є храм цей)" in r.text:
                #print(f, ": є вставка храмового святого")
                pass
            if "ім’я, що його́ пам’ять світло зве́ршуємо" in r.text:
                print(f, ": шаблон")
            #if "Спаси́, Бо́же, наро́д Твій" in r.text:
            #    print(f, ": є Єктенія")
            re_result = re.search(", і (.*), і всіх Святи́х:",r.text)
            if re_result:
                print(re_result.group(1))

        if found:
            re_result = re.search('^(.*?)(\(ім’я, що його́ є храм цей\))(.*?)(і всіх .вяти́х:)(.*?)$',doc.paragraphs[i].text)
            if re_result:
                print(f, len(re_result.groups()), re_result.group(3))
            else:
                print(f, 'RE failed')
                for r in doc.paragraphs[i].runs:
                    print('=',r.text)


            found = False
        
        
        i+=1

doc = docx.Document(files[1])
i=0
for i in range(len(doc.paragraphs)):
    if "Спаси́, Бо́же, наро́д Твій" in doc.paragraphs[i].text:
        print(i)
    i+=1
