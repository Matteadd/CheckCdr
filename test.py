import openpyxl

doc= openpyxl.load_workbook("./CDR_CE_LT068_GSM_MODERNIZATION_ver1.xlsx")

nTrx=doc["BTS"]["X5"].value
startLett="F"

for e in range(0,nTrx+1):

    if doc["BTS"]["A"+startLett+"5"].value==None:
        print("Errore")
    else:
        print(doc["BTS"]["A"+startLett+"5"].value)
    startLett=chr(ord(startLett)+1)
