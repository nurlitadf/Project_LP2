import sys
import os
import comtypes.client

wdFormatPDF = 17

in_file = os.path.abspath("C:\Users\User\Desktop\word.docx")
out_file = os.path.abspath("C:\Users\User\Desktop\word.pdf")

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()