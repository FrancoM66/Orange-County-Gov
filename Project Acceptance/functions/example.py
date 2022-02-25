import pandas as pd
from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Inches
from docx2pdf import convert

def create_word(self, df, variation, project, path, date, item, item2):
    original = r"C:\Users\134360\Desktop\Pyhton project\2-1-2022\Project Acceptance/test/test.docx"

    filepath = path + "/{0}-{1}-{2}-{3}-{4} (Asset List).docx".format(variation,date, project, item, self.planfile_entry.text())

    # Open Microsoft Excel
    doc = Document(original)

    t = doc.add_table(df.shape[0]+1, df.shape[1])
    t.allow_autofit = False
    t.autofit = False
    t.style = 'table90'
    # add the header rows.
    print(df.shape[1])
    
    for j in range(df.shape[-1]):
        print(j + 1)
        if j + 1 == df.shape[1]:
            t.cell(0,j).width = Inches(2.5)
        t.cell(0,j).text = df.columns[j]
        t.cell(0,j).width = Inches(1.01)

    # add the rest of the data frame
    for i in range(df.shape[0]):
        for j in range(df.shape[-1]):
            if j + 1== df.shape[1]:
                t.cell(0,j).width = Inches(2.5)
            t.cell(i+1,j).width = Inches(1.01)
            if str(df.values[i,j]) == "Accepted":
                tcolor = t.cell(i+1,j)
                t.cell(i+1,j).text = str(df.values[i,j])
                shading_elm = parse_xml(r'<w:shd {} w:fill="#50C878"/>'.format(nsdecls('w')))
                tcolor._tc.get_or_add_tcPr().append(shading_elm)
            elif str(df.values[i,j]) == "Rejected":
                tcolor = t.cell(i+1,j)
                t.cell(i+1,j).text = str(df.values[i,j])
                shading_elm = parse_xml(r'<w:shd {} w:fill="FF7F7F"/>'.format(nsdecls('w')))
                tcolor._tc.get_or_add_tcPr().append(shading_elm)
            else:
                t.cell(i+1,j).text = str(df.values[i,j])


    doc.save(filepath)
    convert(filepath)
