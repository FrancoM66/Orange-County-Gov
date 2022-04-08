import email
from win32com import client
from mailmerge import MailMerge
import datetime
from mail import *

def acceptance_no_deficiencies(self,variation, project, path, date, item, item2, text2, workOrder):
    template = "O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Templates\Word\{0}-{1}-{2}-{3} Letter-Template.docx".format(variation, project, item, item2)
    
    document = MailMerge(template)

    today = datetime.datetime.now()
    
    path_2 = r"O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\info"
    filename = self.planfile_entry.text() + ".txt"

    f = open(path_2 + "/" + filename, "r")
    f1 = f.readlines()

    body_text = "All {0} deficiencies are now satisfied or none were found.".format(variation)
    
    config = configparser.RawConfigParser()
    config.read("O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\config\emailCC.properties")

    ocNum  = f1[1].strip()

    if "U" in ocNum:
        print("Utilities Only")
        emailCC = config.get("CCLIST", "ccListUtilities")
        emailCCList = json.loads(emailCC)
    else:
        print("public works")
        emailCC = config.get("CCLIST", "ccListPublicWorks")
        emailCCList = json.loads(emailCC)

    replaced = f1[2].strip()
    replaced = replaced.replace(".", " ")
    replaced = replaced.replace("1", "")
    replaced = replaced.replace("2", "")
    replaced = replaced.replace("3", "")
    replaced = replaced.replace("4", "")
    replaced = replaced.replace("5", "")
    replaced = replaced.replace("6", "")

    ccListForMerge = "";
    for i in range(len(emailCCList)):

        ccListForMerge = ccListForMerge + emailCCList[i] + "\n"
        
        print(ccListForMerge)
        

    document.merge(EmailToCC = ccListForMerge, Inspector = replaced, Project_Name = f1[0].strip(), OCNumber = f1[1].strip(),Date = today.strftime("%B %d, %G"),Inspection_Date = date, Sequence = self.planfile_entry.text(), Body = body_text, Work_Order = text2 )
    file_path = path + "/{0}-{1}-{2}-{3}-{4} (Letter).docx".format(variation,date, project, item, item2)
    pdf_path = path + "/{0}-{1}-{2}-{3}-{4}(Letter).pdf".format(variation, date,project, item, item2)
    filename = "{0}-{1}-{2}-{3}-{4}(Letter and List)".format(variation, date,project, item, item2)
    project_name = f1[0].strip()
    project_variation = variation
    
    
    document.write(file_path)
    document.close()

    wdFormatPDF = 17

    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(file_path)
    doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    mail_to_sign(self,pdf_path,path, filename, item2, workOrder, project_name, project_variation)


def rejected_word(self,variation, project, path, date, item, item2, text2, workOrder):
    template = "O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Templates\Word\{0}-{1}-{2}-{3} Letter-Template.docx".format(variation, project, item, item2)
    document = MailMerge(template)

    today = datetime.datetime.now()

    path_2 = r"O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\info"
    filename = self.planfile_entry.text() + ".txt"

    f = open(path_2 + "/" + filename, "r")
    f1 = f.readlines()

    config = configparser.RawConfigParser()
    config.read("O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\config\emailCC.properties")

    ocNum  = f1[1].strip()
    print(ocNum)

    if "U" in ocNum:
        emailCC = config.get("CCLIST", "ccListUtilities")
        emailCCList = json.loads(emailCC)
    else:
        emailCC = config.get("CCLIST", "ccListPublicWorks")
        emailCCList = json.loads(emailCC)

    ccListForMerge = "";
    for i in range(len(emailCCList)):
        
        ccListForMerge = ccListForMerge + emailCCList[i] + "\n"
        
        print(ccListForMerge)

    replaced = f1[2].strip()
    replaced = replaced.replace(".", " ")
    replaced = replaced.replace("1", "")
    replaced = replaced.replace("2", "")
    replaced = replaced.replace("3", "")
    replaced = replaced.replace("4", "")
    replaced = replaced.replace("5", "")
    replaced = replaced.replace("6", "")
    
    body_text = "Defeciencies were found. Please see attached."
    print(document.get_merge_fields())
    print(emailCCList)
    document.merge(EmailToCC = ccListForMerge,Inspector = replaced, Project_Name = f1[0].strip(), OCNumber = f1[1].strip(),Date = today.strftime("%B %d, %G"),Inspection_Date = date, Sequence = self.planfile_entry.text(), Body = body_text, Work_Order = text2 )
    file_path = path + "/{0}-{1}-{2}-{3}-{4} (Letter).docx".format(variation,date, project, item, item2)
    pdf_path = path + "/{0}-{1}-{2}-{3}-{4}(Letter).pdf".format(variation, date,project, item, item2)
    filename = "{0}-{1}-{2}-{3}-{4}(Letter and List)".format(variation, date,project, item, item2)
    project_variation = variation
    project_name = f1[0].strip()

    document.write(file_path)
    document.close()

    wdFormatPDF = 17

    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(file_path)
    doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    mail_to_sign(self,pdf_path,path, filename, item2, workOrder, project_name, project_variation)
