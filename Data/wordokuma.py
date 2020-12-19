import os
import docx
dosyayol="C:/deneme.docx"
doc = docx.Document('{}'.format(dosyayol))
content = doc.paragraphs

for p in content:
    print(p.text,"\n")


