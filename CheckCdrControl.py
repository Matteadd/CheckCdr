import openpyxl
import time
import datetime
import os
from tkinter import messagebox

class CheckCdrControl:
    def __init__(self, paths):
        super(CheckCdrControl, self).__init__()
        self.paths = paths
        self.listErrGsm={"01":False,"02":False,"03":False,"04":False}
        self.errGsm=False
        self.listLineErrGsm=[]

        for element in self.paths:
            if "GSM" in element or"gsm"in element or "2g"in element:
                self.gsm(element)
                continue
            elif "WCDMA"in element or "wcdma" in element or "3g" in element or "umts" in element or "UMTS" in element:
                self.wcdma(element)
                continue
            elif "LTE" in element or "lte" in element or "4g"in element:
                self.lte(element)
                continue

        if self.errGsm==True:
            if not os.path.exists("./Log Error In Cdr"):
                    os.makedirs("./Log Error In Cdr")
            messagebox.showwarning(message="There are errors in the file. Check the file log in the folder \"Log Error In Cdr\"")
            file=open("./Log Error In Cdr/Log_Error_CDR"+str(datetime.datetime.fromtimestamp(time.time()).strftime('%Y_%m_%d %H_%M'))+".txt","w")
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
                    self.listLineErrGsm.append(element+" is present more than once in column \"TG\"(column \"H\"))\n\n")
            pass

        for element in range(2,nTotCol+2):
            if not(str(doc["BTS"]["F"+str(element)].value) in str(doc["BTS"]["H"+str(element)].value)) and not(str(doc["BTS"]["G"+str(element)].value) in str(doc["BTS"]["H"+str(element)].value)):
                self.errGsm=True
                self.listLineErrGsm.append("The format of "+str(doc["BTS"]["H"+str(element)].value) + " in \"H"+str(element)+"\" in sheet \"BTS\", is wrong. "+
                                           "It must have inside \"" + str(doc["BTS"]["F"+str(element)].value) + "\" and "+str(doc["BTS"]["G"+str(element)].value)+"\" \n\n")
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            if (doc["BTS"]["X"+str(element)].value!=None) or (doc["BTS"]["X"+str(element)].value!="0"):
                if (doc["BTS"]["AB"+str(element)].value==None):
                    self.errGsm=True
                    self.listLineErrGsm.append("The value in "+"AB"+element+" must not be empty"+"\n\n")
                if (doc["BTS"]["AC"+str(element)].value==None):
                    self.errGsm=True
                    self.listLineErrGsm.append("The value in "+"AC"+element+" must not be empty"+"\n\n")
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            if (doc["BTS"]["X"+str(element)].value!=None) or (doc["BTS"]["X"+str(element)].value!="0"):
                nTrx=doc["BTS"]["X"+str(element)].value
                startLett="F"

                for e in range(0,nTrx+1):

                    if doc["BTS"]["A"+startLett+str(element)].value==None:
                        self.errGsm=True
                        self.listLineErrGsm.append("The cell "+"\""+"A"+startLett+str(element)+"\""+"can not be empty.\n\n")

                    startLett=chr(ord(startLett)+1)
# -----------------------------------------------  --------------------------------------------------------------------------------------------------------------------------

            if doc["BTS"]["P"+str(element)].value==None:
                self.errGsm=True
                self.listLineErrGsm.append("The cell "+"\""+"P"+str(element)+"\""+"can not be empty.\n\n")
            if doc["BTS"]["R"+str(element)].value==None:
                self.errGsm=True
                self.listLineErrGsm.append("The cell "+"\""+"R"+str(element)+"\""+"can not be empty.\n\n")
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        adjg2gValue=doc["ADJ G2G"]["A1"].value
        adjg2uValue=doc["ADJ G2U"]["A1"].value

        nColAdjg2g=countCol(doc,"ADJ G2G","A",2)
        nColAdjg2u=countCol(doc,"ADJ G2U","A",2)

        elemInAdjg2g=elemInCol(doc,"ADJ G2G","A",nColAdjg2g)
        elemInAdjg2u=elemInCol(doc,"ADJ G2U","A",nColAdjg2u)

        for element in elemInAdjg2g:
            if element != adjg2gValue:
                self.errGsm=True
                errInSelfG2G=True

            if element != adjg2uValue:
                self.errGsm=True
                errOutSelfG2G=True

        if errInSelfG2G:
            self.listLineErrGsm.append("There are inconsistencies in column \"A\" in the sheet \"ADJ G2G\" All values must be the same.\n\n")

        if errOutSelfG2G:
            self.listLineErrGsm.append("The values in the column \"A\" of the sheet \"ADJ G2G\" mustn't be different from the values in the column \"A\" of the sheet \"ADJ G2U\"\n\n")

        for element in elemInAdjg2u:
            if element != adjg2uValue:
                self.errGsm=True
                errInSelfG2U=True

        if errInSelfG2U:
            self.listLineErrGsm.append("There are inconsistencies in column \"A\" in the sheet \"ADJ G2U\". All values must be the same.\n\n")



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
