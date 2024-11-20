import glob,os, tempfile, shutil
import zipfile

zipnames=glob.glob('*/*/files.zip')
txtnames=glob.glob('*/*/c1.txt')

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


for i in range(len(zipnames)):
    print(zipnames[i], txtnames[i])
    remove_from_zip(zipnames[i], 'c1.txt')
    with zipfile.ZipFile(zipnames[i],'a') as zf:
        zf.write(txtnames[i], arcname=os.path.split(txtnames[i])[1])


    
