import openpyxl
import time
import datetime
import os
from tkinter import messagebox
import time

class CheckCdrControl:
    def __init__(self, paths):
        super(CheckCdrControl, self).__init__()
        self.paths = paths
        self.listErrGsm={"01":False,"02":False,"03":False,"04":False}
        self.errGsm=False
        self.listLineErrGsm=[]

        for element in self.paths:
            if "GSM" in element or"gsm"in element or "2g"in element:
                self.errGsm=False
                self.gsm(element)
                if self.errGsm==True:
                    self.createLog(element)
                continue
            elif "WCDMA"in element or "wcdma" in element or "3g" in element or "umts" in element or "UMTS" in element:
                self.errGsm=False
                self.wcdma(element)
                if self.errGsm==True:
                    self.createLog(element)
                continue
            elif "LTE" in element or "lte" in element or "4g"in element:
                self.errGsm=False
                self.lte(element)
                if self.errGsm==True:
                    self.createLog(element)
                continue
        if self.errGsm==True:

            error=messagebox.askyesno(title="Warning",message="There are errors in CDR. Do you want open the log folder?")
            if error:
                os.startfile(".\\Log Error In Cdr")
        elif  self.errGsm==False:
            messagebox.showinfo(message="There aren't errors")


    def createLog(self, path):
        pathSplit=path.split("/")
        nameFile=pathSplit[-1]
        if not os.path.exists("./Log Error In Cdr"):
            os.makedirs("./Log Error In Cdr")
        file=open("./Log Error In Cdr/Log_"+nameFile+"_"+str(datetime.datetime.fromtimestamp(time.time()).strftime('%m_%d %H_%M'))+".txt","w")
        file.writelines(self.listLineErrGsm)
        file.close()


        pass

    def gsm(self, path):
        doc= openpyxl.load_workbook(path, data_only=True)
        nTotCol=countCol(doc,"BTS","E",2)
        elementColE=elemInCol(doc,"BTS","E",nTotCol)
        listRecurrence=[]

        for element in elementColE:

            recurrence=0
            for elementsToCheck in elementColE:
                if element==elementsToCheck:
                    recurrence+=1
                pass
            pass

            if recurrence>1:
                self.errGsm=True
                if (element in listRecurrence)==False:
                    listRecurrence.append(element)
                    self.listLineErrGsm.append(element+" is present more than once in column \"TG\"(column \"E\"))\n\n")
            pass

        for element in range(2,nTotCol+2):
            if not(str(doc["BTS"]["F"+str(element)].value) in str(doc["BTS"]["H"+str(element)].value)) and not(str(doc["BTS"]["G"+str(element)].value) in str(doc["BTS"]["H"+str(element)].value)):
                self.errGsm=True
                self.listLineErrGsm.append("The format of "+str(doc["BTS"]["H"+str(element)].value) + " in \"H"+str(element)+"\" in sheet \"BTS\", is wrong. "+
                                           "It must have inside \"" + str(doc["BTS"]["F"+str(element)].value) + "\" and "+str(doc["BTS"]["G"+str(element)].value)+"\" \n\n")


            if (doc["BTS"]["X"+str(element)].value!=None) and (doc["BTS"]["X"+str(element)].value!="0"):

                nTrx=doc["BTS"]["X"+str(element)].value
                startLett="F"
                startLett1="C"
                cellName=doc["BTS"]["C"+str(element)].value

                if (doc["BTS"]["AB"+str(element)].value==None):
                    self.errGsm=True
                    self.listLineErrGsm.append("The value in "+"AB"+str(element)+" in sheet \"BTS\" must not be empty"+"\n\n")

                if cellName[-2]=="G":
                    if doc["BTS"]["D"+str(element)].value[-1]!="G":
                        self.errGsm=True
                        self.listLineErrGsm.append("The value in "+"D"+str(element)+" in sheet \"BTS\" must end with \"G\""+"\n\n")
                elif cellName[-2]=="D":
                    if doc["BTS"]["D"+str(element)].value[-1]!="D":
                        self.errGsm=True
                        self.listLineErrGsm.append("The value in "+"D"+str(element)+" in sheet \"BTS\" must end with \"D\""+"\n\n")

                for e in range(0, nTrx):
                    if doc["BTS"]["A"+startLett1+str(element)].value==None:
                        self.errGsm=True
                        self.listLineErrGsm.append("The cell "+"\""+"A"+startLett1+str(element)+"\" in sheet \"BTS\" "+"can not be empty.\n\n")
                    startLett1=chr(ord(startLett1)+1)

                for e in range(0,nTrx+1):

                    if doc["BTS"]["A"+startLett+str(element)].value==None:
                        self.errGsm=True
                        self.listLineErrGsm.append("The cell "+"\""+"A"+startLett+str(element)+"\" in sheet \"BTS\" "+"can not be empty.\n\n")
                    else:
                        if cellName[-2]=="G":
                            if not(doc["BTS"]["A"+startLett+str(element)].value>=100) or not(doc["BTS"]["A"+startLett+str(element)].value<=124):
                                self.errGsm=True
                                self.listLineErrGsm.append("The value in "+"A"+startLett+str(element)+" must be between 100 and 124\n\n")

                        elif cellName[-2]=="D":
                            if not(doc["BTS"]["A"+startLett+str(element)].value>=687) or not(doc["BTS"]["A"+startLett+str(element)].value<=710):
                                self.errGsm=True
                                self.listLineErrGsm.append("The value in "+"A"+startLett+str(element)+" must be between 687 and 710\n\n")



                    startLett=chr(ord(startLett)+1)

            elif (doc["BTS"]["X"+str(element)].value==None) or (doc["BTS"]["X"+str(element)].value=="0"):
                self.errGsm=True
                self.listLineErrGsm.append("The Number of TRX of TCH is empty\n\n")

            if doc["BTS"]["P"+str(element)].value==None:
                self.errGsm=True
                self.listLineErrGsm.append("The cell "+"\""+"P"+str(element)+"\" in sheet \"BTS\" "+"can not be empty.\n\n")

            if doc["BTS"]["R"+str(element)].value==None:
                self.errGsm=True
                self.listLineErrGsm.append("The cell "+"\""+"R"+str(element)+"\" in sheet \"BTS\" "+"can not be empty.\n\n")

        G2GValue=doc["ADJ G2G"]["A2"].value
        G2UValue=doc["ADJ G2U"]["A2"].value
        cont=2
        isDifferent=False
        while doc["ADJ G2U"]["A"+str(cont)].value!=None:
            if doc["ADJ G2U"]["A"+str(cont)].value!=G2GValue:
                isDifferent=True
                break
            cont+=1
        if isDifferent:
            self.errGsm=True
            self.listLineErrGsm.append("The values in the column \"A\" of the sheet \"ADJ G2G\" mustn't be different from the values in the column \"A\" of the sheet \"ADJ G2U\"\n\n")

        cont=2
        isDifferentColBG2G=False
        while doc["ADJ G2G"]["B"+str(cont)].value!=None:
            if doc["ADJ G2G"]["A"+str(cont)].value!=doc["BTS"]["B2"]:
                isDifferentColBG2G=True
                break
            cont+=1

        if isDifferentColBG2G:
            self.errGsm=True
            self.listLineErrGsm.append("The values in the column \"B\" of the sheet \"ADJ G2G\" mustn't be different from the values in the column \"B\" of the sheet \"BTS\"\n\n")

        cont=2
        isDifferentColBG2U=False
        while doc["ADJ G2U"]["B"+str(cont)].value!=None:
            if doc["ADJ G2U"]["A"+str(cont)].value!=doc["BTS"]["B2"]:
                isDifferentColBG2U=True
                break
            cont+=1

        if isDifferentColBG2U:
            self.errGsm=True
            self.listLineErrGsm.append("The values in the column \"B\" of the sheet \"ADJ G2U\" mustn't be different from the values in the column \"B\" of the sheet \"BTS\"\n\n")

        doc.close()
        pass

    def wcdma(self, path):
        doc= openpyxl.load_workbook(path)
        i=2
        nTotRow=0
        while doc["BTS"]["C"+str(i)].value!=None:
            nTotRow+=1
            i+=1
            pass
        print(nTotRow)
        pass

    def lte(self, path):
        doc= openpyxl.load_workbook(path)
        i=2
        nTotRow=0
        while doc["BTS"]["C"+str(i)].value!=None:
            nTotRow+=1
            i+=1
            pass
        print(nTotRow)
        pass

def countCol(excel,sheet, col,startRow):

    i=startRow
    nTotCol=0
    while excel[sheet][col+str(i)].value!=None:
        nTotCol+=1
        i+=1
        pass

    return nTotCol
    pass

def elemInCol(excel,sheet, col, nTotCol):
    elementCol=[]
    for element in range(2,nTotCol+2):
        elementCol.append(excel[sheet][col+str(element)].value)
        pass
    return elementCol
    pass
