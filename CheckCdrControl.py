import openpyxl
import time
import datetime

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
            file=open("Log_Error_CDR"+str(datetime.datetime.fromtimestamp(time.time()).strftime('%Y_%m_%d %H_%M'))+".txt","w")
            file.writelines(self.listLineErrGsm)
            file.close()
            pass

    def gsm(self, path):
        doc= openpyxl.load_workbook(path)
        nTotCol=countCol(doc,"BTS","E")
        elementColH=elemInCol(doc,"BTS","E",nTotCol)
        temp=[]

        for element in elementColH:

            recurrence=0
            for elementsToCheck in elementColH:
                if element==elementsToCheck:
                    recurrence+=1
                pass
            pass

            if recurrence>1:
                self.errGsm=True
                if (element in temp)==False:
                    temp.append(element)
                    self.listLineErrGsm.append(element+" is present more than once in column \"TG\"(column \"H\"))\n")



            else:
                pass

            pass
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

def countCol(excel,sheet, col):

    i=2
    nTotCol=0
    while excel[sheet][col+str(i)].value!=None:
        nTotCol+=1
        i+=1
        pass

    return nTotCol
    pass

def elemInCol(excel,sheet, col, nTotCol):
    elementColH=[]
    for element in range(2,nTotCol+2):
        elementColH.append(excel[sheet][col+str(element)].value)
        pass
    return elementColH
    pass
