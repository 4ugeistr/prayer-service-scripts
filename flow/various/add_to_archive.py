import glob,os, tempfile, shutil
import zipfile

zipnames=glob.glob('*/files.zip')
txtnames=glob.glob('*/*.html')

dic={}
for z in zipnames:
    dic[z[0:2]]=[]
for t in txtnames:
    dic[t[0:2]]+=[t[3:]]

def remove_from_zip(zipfname, *filenames):
    tempdir = tempfile.mkdtemp()
    try:
        tempname = os.path.join(tempdir, 'new.zip')
        with zipfile.ZipFile(zipfname, 'r') as zipread:
            with zipfile.ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.infolist():
                    if item.filename not in filenames:
                        data = zipread.read(item.filename)
                        zipwrite.writestr(item, data)
        shutil.move(tempname, zipfname)
    finally:
        shutil.rmtree(tempdir)
        print('done')

for k,v in dic.items():
    for filename in v:
        print(k+'\\files.zip', k+'\\'+filename)
        remove_from_zip(k+'\\files.zip', filename)
        remove_from_zip(k+'\\files.zip', filename)
        remove_from_zip(k+'\\files.zip', filename)
        with zipfile.ZipFile(k+'\\files.zip','a') as zf:
            zf.write(k+'\\'+filename, arcname=filename)

'''
for i in range(len(zipnames)):
    print(zipnames[i], txtnames[i])
    remove_from_zip(zipnames[i], 'c1.txt')
    with zipfile.ZipFile(zipnames[i],'a') as zf:
        zf.write(txtnames[i], arcname=os.path.split(txtnames[i])[1])
'''

    
