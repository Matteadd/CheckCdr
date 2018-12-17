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
        self.errGsm=False
        self.listLineErrGsm=[]

        for element in self.paths:
            if element!= None:
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
                pass


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
        self.listLineErrGsm=[]
        file.close()


        pass

    def gsm(self, path):
        doc= openpyxl.load_workbook(path, data_only=True)
        nTotCol=countCol(doc,"BTS","C",2)
        elementColE=elemInCol(doc,"BTS","E",nTotCol)

        listRecurrence=[]

        # in questo for controllo se ci sono valori ripetuti in H
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
                    self.listLineErrGsm.append("The \""+element+"\" is present more than once in column \"TG\" in sheet \"BTS\"\n\n")
            pass

        # in questo for faccio diversi controlli su alcuni campi
        for element in range(2,nTotCol+2):
            cellName=doc["BTS"]["C"+str(element)].value

            if cellName[-2]=="G":
                if doc["BTS"]["K"+str(element)].value<100 or doc["BTS"]["K"+str(element)].value>124:
                    self.errGsm=True
                    self.listLineErrGsm.append("The value in \""+doc["BTS"]["K"+str(1)].value+"\"(K"+str(element)+") must be between 100 and 124\n\n")

            elif cellName[-2]=="D":
                if doc["BTS"]["K"+str(element)].value<687 or doc["BTS"]["K"+str(element)].value>710:
                    self.errGsm=True
                    self.listLineErrGsm.append("The value in \""+doc["BTS"]["K"+str(1)].value+"\"(K"+str(element)+") must be between 687 and 710\n\n")


            # controllo che se il valoe di h ci devono stare f e g
            if not(str(doc["BTS"]["F"+str(element)].value) in str(doc["BTS"]["H"+str(element)].value)) or not(str(doc["BTS"]["G"+str(element)].value) in str(doc["BTS"]["H"+str(element)].value)):
                self.errGsm=True
                self.listLineErrGsm.append("The format of \""+str(doc["BTS"]["H"+str(element)].value) + "\" in column \""+str(doc["BTS"]["H"+str(1)].value)+"\"(H"+str(element)+") in sheet \"BTS\", is wrong. "+
                                           "It must have inside \""+ str(doc["BTS"]["F"+str(element)].value)+ "\""+"(F"+str(element)+") and \""+str(doc["BTS"]["G"+str(element)].value)+"\"(G"+str(element)+")\n\n")

            # Diversi controlli sul TRX in caso non fosse vuoto o diverso da 0
            if (doc["BTS"]["X"+str(element)].value!=None) and (doc["BTS"]["X"+str(element)].value!="0"):

                nTrx=doc["BTS"]["X"+str(element)].value
                startLett="F"
                startLett1="C"
                cellName=doc["BTS"]["C"+str(element)].value
                # se il valore in HSN (CHGR-1) (AB) è vuoto restituisco errore
                if (doc["BTS"]["AB"+str(element)].value==None):
                    self.errGsm=True
                    self.listLineErrGsm.append("The value in \""+doc["BTS"]["AB"+str(1)].value+"\"(AB"+str(element)+") in sheet \"BTS\" must not be empty"+"\n\n")

                # se la penultima lettera di cellname(C) è G allora anche rSite(D) deve finire in d, stessa cosa per Cellname(c)con G
                if cellName[-2]=="G":
                    if doc["BTS"]["D"+str(element)].value[-1]!="G":
                        self.errGsm=True
                        self.listLineErrGsm.append("The value in \""+doc["BTS"]["D"+str(1)].value+"\"(D"+str(element)+") in sheet \"BTS\" must end with \"G\""+"\n\n")
                elif cellName[-2]=="D":
                    if doc["BTS"]["D"+str(element)].value[-1]!="D":
                        self.errGsm=True
                        self.listLineErrGsm.append("The value in \""+doc["BTS"]["D"+str(1)].value+"\"(D"+str(element)+") in sheet \"BTS\" must end with \"D\""+"\n\n")

                # i maio devono essere compilati almeno per un numero pari ai trx e il loro valore oscilla tra 0 e il numero di colonne tchFreq compilate
                for e in range(0, nTrx):
                    if doc["BTS"]["A"+startLett1+str(element)].value==None:
                        self.errGsm=True
                        self.listLineErrGsm.append("The cell in \""+doc["BTS"]["A"+startLett1+str(1)].value+"\"(A"+startLett1+str(element)+") in sheet \"BTS\" "+"can not be empty.\n\n")
                    elif doc["BTS"]["A"+startLett1+str(element)].value!=None:
                        valMax=countRowFill(doc, element, "AF")
                        if doc["BTS"]["A"+startLett1+str(element)].value>valMax or doc["BTS"]["A"+startLett1+str(element)].value<0:
                            self.errGsm=True
                            self.listLineErrGsm.append("The value in \""+doc["BTS"]["A"+startLett1+str(1)].value+"\"(A"+startLett1+str(element)+") must be between 0 and "+str(valMax)+"\n\n")
                    startLett1=chr(ord(startLett1)+1)

                # i tchFreq devono essere compilati per un numero almeno pari a trx + 1 e
                # il loro vvalore oscilla tra 687 e 710 in caso cellname sia D altrimenti tra 100 e 124 in caso cell name sia d
                for e in range(0,nTrx+1):

                    if doc["BTS"]["A"+startLett+str(element)].value==None:
                        self.errGsm=True
                        self.listLineErrGsm.append("The cell in "+doc["BTS"]["A"+startLett+str(1)].value+"(A"+startLett+str(element)+") in sheet \"BTS\" "+"can not be empty.\n\n")
                    else:
                        if cellName[-2]=="G":
                            if doc["BTS"]["A"+startLett+str(element)].value<100 or doc["BTS"]["A"+startLett+str(element)].value>124:
                                self.errGsm=True
                                self.listLineErrGsm.append("The value in \""+doc["BTS"]["A"+startLett+str(1)].value+"\"(A"+startLett+str(element)+") must be between 100 and 124\n\n")

                        elif cellName[-2]=="D":
                            if doc["BTS"]["A"+startLett+str(element)].value<687 or doc["BTS"]["A"+startLett+str(element)].value>710:
                                self.errGsm=True
                                self.listLineErrGsm.append("The value in \""+doc["BTS"]["A"+startLett+str(1)].value+"\"(A"+startLett+str(element)+") must be between 687 and 710\n\n")
                    startLett=chr(ord(startLett)+1)

            # se TRX è vuota controllo che da af a ak della stessa colonna siano vuote
            elif (doc["BTS"]["X"+str(element)].value==None) or (doc["BTS"]["X"+str(element)].value=="0"):

                errTrxEmpty=False

                if doc["BTS"]["AF"+str(element)].value!=None:
                    errTrxEmpty=True
                if doc["BTS"]["AG"+str(element)].value!=None:
                    errTrxEmpty=True
                if doc["BTS"]["AH"+str(element)].value!=None:
                    errTrxEmpty=True
                if doc["BTS"]["AI"+str(element)].value!=None:
                    errTrxEmpty=True
                if doc["BTS"]["AJ"+str(element)].value!=None:
                    errTrxEmpty=True
                if doc["BTS"]["AK"+str(element)].value!=None:
                    errTrxEmpty=True

                if errTrxEmpty==True:
                    self.errGsm=True
                    self.listLineErrGsm.append("The \"Number of TRX of TCH\"(X"+str(element)+")is empty and some field from coloumn \"tchFreq-0 (CHGR-1)\"(AF) to \"tchFreq-5 (CHGR-1)\"(AK) are filled\n\n")

            #controllo che la frequenza k non sia in nessuna frequenza da af a ak
            if (doc["BTS"]["K"+str(element)].value!=None) and (doc["BTS"]["K"+str(element)].value!="0"):
                valInK=doc["BTS"]["K"+str(element)].value
                startLett="F"
                while startLett!="L":
                    for cont in range(2,nTotCol+2):
                        if valInK==doc["BTS"]["A"+startLett+str(cont)].value:
                            self.errGsm=True
                            self.listLineErrGsm.append("The frequence in \"bcchFreq (CHGR-0)\"(K"+str(element)+") is present also in column \""+doc["BTS"]["A"+startLett+str(1)].value+"\"(A"+startLett+")\n\n")
                    startLett=chr(ord(startLett)+1)

            # Se il campo BSPWRB p sono vuoti restituisco errore
            if doc["BTS"]["P"+str(element)].value==None:
                self.errGsm=True
                self.listLineErrGsm.append("The cell in \""+doc["BTS"]["P"+str(1)].value+"\"(P"+str(element)+") in sheet \"BTS\" "+"can not be empty.\n\n")

            # Se il campo BSPWR r sono vuoti restituisco errore
            if doc["BTS"]["R"+str(element)].value==None:
                self.errGsm=True
                self.listLineErrGsm.append("The cell in \""+doc["BTS"]["R"+str(1)].value+"\"(R"+str(element)+") in sheet \"BTS\" "+"can not be empty.\n\n")

            # il campo in ab deve essere compreso tra 0 e 63
            if doc["BTS"]["AB"+str(element)].value<0 or doc["BTS"]["AB"+str(element)].value>63:
                self.errGsm=True
                self.listLineErrGsm.append("The value in \""+doc["BTS"]["AB"+str(1)].value+"\"(AB"+str(element)+") must be between 0 and 63.\n\n")

        colWithDiffInSameGTU=diffInSameCol(doc,"ADJ G2U",["A", "B"],2)
        if colWithDiffInSameGTU:
            for element in colWithDiffInSameGTU:
                nameCol= doc["ADJ G2U"][element+str(1)].value
                self.errGsm=True
                self.listLineErrGsm.append("There are values different in column \""+ str(nameCol) + "\" in sheet \"ADJ G2U\"\n\n")

        colWithDiffInSameG2G=diffInSameCol(doc,"ADJ G2G",["A", "B"],2)
        if colWithDiffInSameG2G:
            for element in colWithDiffInSameG2G:
                nameCol= doc["ADJ G2G"][element+str(1)].value
                self.errGsm=True
                self.listLineErrGsm.append("There are values different in column \""+ str(nameCol) + "\" in sheet \"ADJ G2G\"\n\n")

        colWithDiffInSameBTS=diffInSameCol(doc,"BTS",["B"],2)
        if colWithDiffInSameBTS:
            for element in colWithDiffInSameBTS:
                nameCol= doc["BTS"][element+str(1)].value
                self.errGsm=True
                self.listLineErrGsm.append("There are values different in column \""+ str(nameCol) + "\" in sheet \"BTS\"\n\n")

        colWithDiffInOtherG2U=diffInOtherCol(doc, "ADJ G2U", "B", 2, "BTS", "B", 2)
        if colWithDiffInOtherG2U:
            nameColBTS= doc["BTS"]["B1"].value
            nameColG2U= doc["ADJ G2U"]["B1"].value
            self.errGsm=True
            self.listLineErrGsm.append("There are values different between column \""+ str(nameColBTS) + "\" in sheet \"BTS\" and column \""+ str(nameColG2U) + "\" in sheet \"ADJ G2U\"\n\n")

        colWithDiffInOtherG2G=diffInOtherCol(doc, "ADJ G2G", "B", 2, "BTS", "B", 2)
        if colWithDiffInOtherG2G:
            nameColBTS= doc["BTS"]["B1"].value
            nameColG2G= doc["ADJ G2G"]["B1"].value
            self.errGsm=True
            self.listLineErrGsm.append("There are values different between column \""+ str(nameColBTS) + "\" in sheet \"BTS\" and column \""+ str(nameColG2G) + "\" in sheet \"ADJ G2G\"\n\n")



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

def countRowFill(excel, row, startCol):

    startColList=list(startCol)
    colFill=0

    while "".join(startColList)!="AL":
        if excel["BTS"]["".join(startColList)+str(row)].value!=None:
            colFill+=1
        startColList[-1]=chr(ord(startColList[-1])+1)
    return colFill
    pass

def elemInCol(excel,sheet, col, nTotCol):
    elementCol=[]
    for element in range(2,nTotCol+2):
        elementCol.append(excel[sheet][col+str(element)].value)
        pass
    return elementCol
    pass

def diffInSameCol(doc, sheet, col, rowStart):
    err=[]

    for column in col:
        cont=rowStart
        diff=False
        valToCheck=doc[sheet][str(column)+str(cont)].value
        while doc[sheet][str(column)+str(cont)].value!=None:
            if doc[sheet][str(column)+str(cont)].value!= valToCheck:
                diff=True
            cont+=1
        if diff:
            err.append(str(column))
    return err

    pass

def diffInOtherCol(doc, sheet, col, rowStart, sheetToCheck, colToCheck, rowToCheck):
    err=[]
    valToCheck=doc[sheetToCheck][str(colToCheck)+str(rowToCheck)].value
    cont=rowStart
    diff=False

    while doc[sheet][str(col)+str(cont)].value!=None:
        if doc[sheet][str(col)+str(cont)].value!= valToCheck:
            diff=True
        cont+=1
    if diff:
        return  True
    else:
        return  False
    pass
