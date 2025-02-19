import os
import win32com.client as win32
from win32com.client import constants

# Path to folder containing .doc files
folder_path = os.getcwd()

# Create a Word application object
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

# Get a list of all .doc files in the directory
docs = [f for f in os.listdir(folder_path) if f.endswith(".doc")]

for doc in docs:
    # Open the .doc file
    in_file = os.path.join(folder_path, doc)
    print(in_file)
    doc = word.Documents.Open(in_file)

    # Save as .docx
    out_file = os.path.join(folder_path, in_file[:-4] + ".docx")
    print(out_file)
    doc.SaveAs(out_file, FileFormat=constants.wdFormatXMLDocument)

    # Close the .doc file
    doc.Close()

word.Quit()



        
