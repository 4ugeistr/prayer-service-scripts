import re

allowed = ('<p>І нині, і повсякчас, і на віки віків. Амінь.</p>',
           '<p>Господи, помилуй <i>(40 р.)</i>.</p>',
           '<p>Слава Отцю, і Сину, і Святому Духові.</p>',
           '<p><i>*****</i></p>')

f = open( "c:\\tmp\\input.html", "r",encoding='utf-8' )
t=f.read()
f.close()

res=re.findall(r'\n.*?(Тропар|Кондак).*?\n(.*?)\n',t)
for i in res:
    if not i[1] in allowed:
        print(i[0],i[1])
