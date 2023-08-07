import os
import win32com.client as win32

# Directory of files being converted
os.chdir(r'C:\Desktop')

word = win32.gencache.EnsureDispatch('Word.Application')

for file in os.listdir('.'):
    if file.endswith('.doc') or file.endswith('.docx'):
        doc = word.Documents.Open(os.path.abspath(file))
        doc.SaveAs(os.path.abspath(file[:-4] + '.pdf'), FileFormat=win32.constants.wdFormatPDF)
        doc.Close()

word.Quit()
