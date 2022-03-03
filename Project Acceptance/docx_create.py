from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Inches
from win32com import client
from mailmerge import MailMerge

def create_word(self, df, variation, project, path, date, item, item2):
    original = r"O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Temp\Asset List Temp.docx"
    original_temp = r"O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Temp\asset_temp.docx"

    filepath = path + "/{0}-{1}-{2}-{3}-{4} (Asset List).docx".format(variation,date, project, item, self.planfile_entry.text())
    pdf_path = path + "//{0}-{1}-{2}-{3}-{4}(Asset List).pdf".format(variation,date, project, item, self.planfile_entry.text())
    # Open Microsoft Excel
    doc = Document(original)
    
    t = doc.add_table(df.shape[0]+1, df.shape[1])
    t.allow_autofit = True
    t.autofit = True
    t.style = 'table90'
    # add the header rows.
    print(df.shape[1])
    if df.shape[1] <= 5:
        final_width = 4.50
    elif df.shape[1] > 5 and df.shape[1] <= 7:
        final_width = 3.50
    elif df.shape[1] >= 8:
        final_width = 2.00

    for j in range(df.shape[-1]):
        print(j + 1)
        if j + 1 == df.shape[1]:
            t.cell(0,j).width = Inches(final_width)
        t.cell(0,j).text = df.columns[j]
        t.cell(0,j).width = Inches(1.01)

    # add the rest of the data frame
    for i in range(df.shape[0]):
        for j in range(df.shape[-1]):
            if j + 1== df.shape[1]:
                t.cell(0,j).width = Inches(final_width)
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


    doc.save(original_temp)
    

    mergedoc = MailMerge(original_temp)
    print(mergedoc.get_merge_fields())
    mergedoc.merge(Sequence = self.planfile_entry.text())
    mergedoc.write(filepath)
    mergedoc.close()

    wdFormatPDF = 17
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(filepath)
    doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

