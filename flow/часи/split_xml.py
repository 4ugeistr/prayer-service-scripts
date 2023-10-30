sep = "<!--END-->"
header=[
    '<h2><b>ЧАС ПЕРШИЙ</b></h2>',
    '<h2><b>ЧАС ТРЕТІЙ</b></h2>',
    '<h2><b>ЧАС ШОСТИЙ</b></h2>',
    '<h2><b>ЧАС ДЕВ\'ЯТИЙ</b></h2>',
    '<h2 align="center" ><a id=\'t1\'><b>ЧАС ПЕРШИЙ</b></a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a href=\'#t9\'>[9]</a></h2>',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a id=\'t3\'><b>ЧАС ТРЕТІЙ</b></a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a href=\'#t9\'>[9]</a></h2>',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a id=\'t6\'><b>ЧАС ШОСТИЙ</b></a> &middot; <a href=\'#t9\'>[9]</a></h2>',
    '<h2 align="center" ><a href=\'#t1\'>[1]</a> &middot; <a href=\'#t3\'>[3]</a> &middot; <a href=\'#t6\'>[6]</a> &middot; <a id=\'t9\'><b>ЧАС ДЕВ\'ЯТИЙ</b></a></h2>'
    ]



first=True
outfile = open( "c:\\tmp\\input_w_sep.html", "w",encoding='utf-8')
for  aLine in open( "c:\\tmp\\input.html", "r",encoding='utf-8' ):
    if aLine.strip()==header[0]:
        if first:
            first=False
            print(header[0], file=outfile)
        else:
            print(sep, file=outfile)
            print(header[0], file=outfile)
    elif aLine.strip()=='<p>&nbsp;</p>':
        continue
    else:
        print(aLine, file=outfile, end="")
        
outfile= open( "c:\\tmp\\input_w_sep.html", "a",encoding='utf-8')
print("",file=outfile)
print(sep, file=outfile)   
outfile.close()

i = 1    
outfile = open( "c:\\tmp\\t"+"{:02d}".format(i)+"c.html", "w",encoding='utf-8')
for aLine in open( "c:\\tmp\\input_w_sep.html", "r",encoding='utf-8' ):
    #print(aLine)
    #if i==30:
        #print(aLine[:10])
    
    if  aLine.strip() == sep:
        outfile.close()
        i += 1
        outfile = open( "c:\\tmp\\t"+"{:02d}".format(i)+"c.html", "w",encoding='utf-8'  )
    else:
        #print('written!')
        written=False
        for j in range(4):
            if aLine.strip()==header[j]:
                print( header[j+4], file=outfile, end="" )
                written=True
        if not written:
            print( aLine, file=outfile, end="" )
outfile.close()
