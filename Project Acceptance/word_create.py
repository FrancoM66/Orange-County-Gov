from win32com import client
from mailmerge import MailMerge
import datetime
from mail import *

def acceptance_no_deficiencies(self,variation, project, path, date, item, item2, text2):
    template = "O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Templates\Word\{0}-{1}-{2}-{3} Letter-Template.docx".format(variation, project, item, item2)
    print(template)
    document = MailMerge(template)
    print(document.get_merge_fields())

    today = datetime.datetime.now()
    print(today.strftime("%B, %d %G"))
    
    body_text = "All {0} deficiencies are now satisfied or none were found.".format(variation)
    document.merge(Date = today.strftime("%B %d, %G"),Inspection_Date = date, Sequence = self.planfile_entry.text(), Body = body_text, Work_Order = text2 )
    print("This is the path before the write: "+ path)
    file_path = path + "/{0}-{1}-{2}-{3}-{4} (Letter).docx".format(variation,date, project, item, item2)
    pdf_path = path + "/{0}-{1}-{2}-{3}-{4}(Letter).pdf".format(variation, date,project, item, item2)
    document.write(file_path)
    document.close()

    wdFormatPDF = 17

    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(file_path)
    doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    mail_to_sign(self,pdf_path,path)


def rejected_word(self,variation, project, path, date, item, item2, text2):
    template = "O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Templates\Word\{0}-{1}-{2}-{3} Letter-Template.docx".format(variation, project, item, item2)
    print(template)
    document = MailMerge(template)
    print(document.get_merge_fields())

    today = datetime.datetime.now()
    print(today.strftime("%B %d, %G"))
    
    body_text = "Defeciencies were found. Please see attached."
    document.merge(Date = today.strftime("%B %d, %G"),Inspection_Date = date, Sequence = self.planfile_entry.text(), Body = body_text, Work_Order = text2 )
    print("This is the path before the write: "+ path)
    file_path = path + "/{0}-{1}-{2}-{3}-{4} (Letter).docx".format(variation,date, project, item, item2)
    pdf_path = path + "/{0}-{1}-{2}-{3}-{4}(Letter).pdf".format(variation, date,project, item, item2)
    document.write(file_path)
    document.close()

    wdFormatPDF = 17

    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(file_path)
    doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    mail_to_sign(self,pdf_path,path)
