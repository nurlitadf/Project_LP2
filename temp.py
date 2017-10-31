import sys
import os
import comtypes.client
import os, PythonMagick
from PythonMagick import Image
from datetime import datetime

wdFormatPDF = 17

in_file = os.path.abspath("C:\Users\User\Desktop\word.docx")
out_file = os.path.abspath("C:\Users\User\Desktop\word.pdf")

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

pdf_dir = os.path.dirname("C:\Users\User\Desktop\word.pdf")
bg_colour = "#ffffff"

for pdf in [pdf_file for pdf_file in os.listdir(pdf_dir) if pdf_file.endswith(".pdf")]:

    input_pdf = pdf_dir + "\\" + pdf
    img = Image()
    img.density('300')
    img.read(input_pdf)

    size = "%sx%s" % (img.columns(), img.rows())

    output_img = Image(size, bg_colour)
    output_img.type = img.type
    output_img.composite(img, 0, 0, PythonMagick.CompositeOperator.SrcOverCompositeOp)
    output_img.resize(str(img.rows()))
    output_img.magick('JPG')
    output_img.quality(75)


    output_jpg = input_pdf.replace(".pdf", ".jpg")
    output_img.write(output_jpg)
