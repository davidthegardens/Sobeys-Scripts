import os
import shutil
import zipfile
import uuid
import pathlib
from tkinter import filedialog
import pandas as pd

def Unpack(From,To):
    images=[]
    for dirpath,_,filenames in os.walk(From):
        for f in filenames:
            movethis=os.path.abspath(os.path.join(dirpath, f)) 
            file_extension = pathlib.Path(movethis).suffix
            if file_extension.lower() in [".jpg",".png",".jpeg"]:
                if (("600" in f) and (("1080" in f) or ("630" in f))) or (f.count("510")==2 or (("700" in f) and ("1200" in f or "750" in f))):
                    images.append(f)
                    shutil.copy(movethis,To)
                    print(movethis + " >>> " + To)
    return images

tempdir="C:\\Users\\desjardav\\Downloads\\"+str(uuid.uuid4().hex)
os.mkdir(tempdir)
source=str(filedialog.askopenfilename())

print("unzipping")
with zipfile.ZipFile(source, 'r') as zip_ref:
    zip_ref.extractall(tempdir)

images=Unpack(tempdir,filedialog.askdirectory())



df=pd.DataFrame(images,columns=['images'])
df.reset_index(drop=True, inplace=True)
df.to_clipboard()

shutil.rmtree(tempdir)
shutil.rmtree(source)