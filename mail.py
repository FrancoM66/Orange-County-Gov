import configparser
import json
import win32com.client as win32
import glob
import os

def mail_to_sign(self, path, listpath, filename, AorR, workOrder, project_name,variation):

    path_asset = listpath +"\*(Asset List).pdf"
    list_of_files = glob.glob(path_asset) # * means all if need specific format then *.csv
    latest_file_list = max(list_of_files, key=os.path.getctime)
    print(latest_file_list)
    config = configparser.RawConfigParser()
    config.read("O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\config\emailCC.properties")

    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    # construct email item object
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = filename 
    mailItem.BodyFormat = 1

    if(variation == "Pressurized-Pipe"):

        pipeSendTo = config.get("PRESSURE", "send4sign")
        pipelist = json.loads(pipeSendTo)
        mailItem.HTMLBody = "<p>Mr.Rivera <br><br> Please sign and save the inspection letter using the following links: <br></p><a href='{0}'>{1}</a><br><p>Asset List:</p><a href='{2}'>{3}</a><br><ul><li>Project Name : {4}</li><li>Sequence Number: {5}</li></ul>".format(path, path,latest_file_list,latest_file_list,project_name,workOrder)
        mailItem.To = pipelist[0]

    elif(variation == "Gravity"):

        gravitySendTo = config.get("GRAVITY", "send4sign")
        gravitylist = json.loads(gravitySendTo)
        mailItem.HTMLBody = "<p>Mr.Reyes <br><br> Please sign and save the inspection letter using the following links: <br></p><a href='{0}'>{1}</a><br><p>Asset List:</p><a href='{2}'>{3}</a><br><ul><li>Project Name : {4}</li><li>Sequence Number: {5}</li></ul>".format(path, path,latest_file_list,latest_file_list,project_name,workOrder)
        mailItem.To = gravitylist[0]
        
    elif(variation == "Pump-Station"):

        pumpSendTo = config.get("PUMP", "send4sign")
        pumplist = json.loads(pumpSendTo)
        mailItem.HTMLBody = "<p>Mr.Brown <br><br> Please sign and save the inspection letter using the following links: <br></p><a href='{0}'>{1}</a><br><p>Asset List:</p><a href='{2}'>{3}</a><br><ul><li>Project Name : {4}</li><li>Sequence Number: {5}</li></ul>".format(path, path,latest_file_list,latest_file_list,project_name,workOrder)
        mailItem.To = pumplist[0]
        
    mailItem.Sensitivity  = 2
    mailItem.Display()

def mail_signed(self, path, workOrder):
    path_list = path + "\*(Letter).pdf"
    list_of_files = glob.glob(path_list)
    latest_file_letter = max(list_of_files, key=os.path.getctime)
    print(latest_file_letter)

    path_asset = path +"\*(Asset List).pdf"
    list_of_files = glob.glob(path_asset) 
    latest_file_list = max(list_of_files, key=os.path.getctime)
    print(latest_file_list)

    path_2 = r"O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\info"
    filename = self.planfile_entry.text() + ".txt"

    f = open(path_2 + "/" + filename, "r")
    f1 = f.readlines()
    print(f1)
    print(f1[0])

    replaced = f1[2].strip()
    replaced = replaced.replace(".", " ")
    replaced = replaced.replace("1", "")
    replaced = replaced.replace("2", "")
    replaced = replaced.replace("3", "")
    replaced = replaced.replace("4", "")
    replaced = replaced.replace("5", "")
    replaced = replaced.replace("6", "")

    config = configparser.RawConfigParser()
    config.read("O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\config\emailCC.properties")

    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    if "Warranty" in latest_file_letter:
        walkorwar = "Warranty"
    elif "Walkthrough" in latest_file_letter:
        walkorwar = "Walkthrough"

    ocNum  = f1[1].strip()

    if "U" in ocNum:
        isPublicWorks = False
        devtype = "Utilities"
    else:
        isPublicWorks =True
        devtype = "Public Works"

    # construct email item object
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = self.cip_dev + "-" + walkorwar + " Inspection-" + f1[0].strip() + "- " + f1[1].strip()
    mailItem.BodyFormat = 1
    
    if(self.area == "Pressurized-Pipe" and isPublicWorks == True and self.cip_dev == "Dev"):

        pipeSendTo = config.get("PRESSURE", "sendSignedToDevPub")
        pipelist = json.loads(pipeSendTo)

        pressureCC = config.get("PRESSURE", "sendSignedDevPubCC")
        CCList = json.loads(pressureCC)
        mailItem.CC = CCList[0] + f1[2].strip() + "@ocfl.net"

        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)
        mailItem.To = pipelist[0]

    elif(self.area == "Gravity" and isPublicWorks == True and self.cip_dev == "Dev"):

        gravitySendTo = config.get("GRAVITY", "sendSignedToDevPub")
        gravitylist = json.loads(gravitySendTo)
        
        gravityCC = config.get("GRAVITY", "sendSignedDevPubCC")
        CCList = json.loads(gravityCC)
        mailItem.CC = CCList[0]+ f1[2].strip() + "@ocfl.net"
        
        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = gravitylist[0]
        
    elif(self.area == "Pump-Station" and isPublicWorks == True and self.cip_dev == "Dev"):

        pumpSendTo = config.get("PUMP", "sendSignedToDevPub")
        pumplist = json.loads(pumpSendTo)

        pumpCC = config.get("PUMP", "sendSignedDevPubCC")
        CCList = json.loads(pumpCC)
        mailItem.CC = CCList[0]+ f1[2].strip() + "@ocfl.net"
        
        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = pumplist[0]

    if(self.area == "Pressurized-Pipe" and isPublicWorks == False and self.cip_dev == "Dev"):

        pipeSendTo = config.get("PRESSURE", "sendSignedToDevUtil")
        pipelist = json.loads(pipeSendTo)

        pressureCC = config.get("PRESSURE", "sendSignedDevUtilCC")
        CCList = json.loads(pressureCC)
        mailItem.CC = CCList[0]+ f1[2].strip() + "@ocfl.net"

        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = pipelist[0]

    elif(self.area == "Gravity" and isPublicWorks == False and self.cip_dev == "Dev"):

        gravitySendTo = config.get("GRAVITY", "sendSignedToDevUtil")
        gravitylist = json.loads(gravitySendTo)

        gravityCC = config.get("GRAVITY", "sendSignedDevUtilCC")
        CCList = json.loads(gravityCC)
        mailItem.CC = CCList[0]+ f1[2].strip() + "@ocfl.net"

        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = gravitylist[0]
        
    elif(self.area == "Pump-Station" and isPublicWorks == False and self.cip_dev == "Dev"):

        pumpSendTo = config.get("PUMP", "sendSignedToDevUtil")
        pumplist = json.loads(pumpSendTo)
        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = pumplist[0]

    if(self.area == "Pressurized-Pipe" and isPublicWorks == False and self.cip_dev == "CIP"):

        pipeSendTo = config.get("PRESSURE", "sendSignedCip")
        pipelist = json.loads(pipeSendTo)

        pressureCC = config.get("PRESSURE", "sendSignedCipCC")
        CCList = json.loads(pressureCC)
        mailItem.CC = CCList[0]+ f1[2].strip() + "@ocfl.net"

        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = pipelist[0]

    elif(self.area == "Gravity" and isPublicWorks == False and self.cip_dev == "CIP"):

        gravitySendTo = config.get("GRAVITY", "sendSignedCip")
        gravitylist = json.loads(gravitySendTo)

        gravityCC = config.get("GRAVITY", "sendSignedCipCC")
        CCList = json.loads(gravityCC)
        mailItem.CC = CCList[0]+ f1[2].strip() + "@ocfl.net"

        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = gravitylist[0]
        
    elif(self.area == "Pump-Station" and isPublicWorks == False and self.cip_dev == "CIP"):

        pumpSendTo = config.get("PUMP", "sendSignedCip")
        pumplist = json.loads(pumpSendTo)

        pumpCC = config.get("PUMP", "sendSignedCipCC")
        CCList = json.loads(pumpCC)
        mailItem.CC = CCList[0]+ f1[2].strip() + "@ocfl.net"

        mailItem.Attachments.Add(Source=latest_file_letter)
        mailItem.Attachments.Add(Source=latest_file_list)
        mailItem.HTMLBody = "<p>Attached you will find the results of the <b>{}</b> inspections performed by Operations for the following {} project:</p> <br><br> <ul><li>Project name- {}</li><li>Warranty inspections: {}</li><li>Sequence number: {}</li><li>OC project number: {}</li><li>Development Type: {}</li></ul><br><p>Thanks,</p>".format(self.cip_dev,walkorwar,f1[0].strip(), self.area,self.planfile_entry.text(),f1[1].strip(),devtype)        
        mailItem.To = pumplist[0]


    mailItem.Sensitivity  = 2
    mailItem.Display()
