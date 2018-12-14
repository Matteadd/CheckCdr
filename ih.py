import openpyxl

startCol="AF"

startColList=list(startCol)
colFill=0

excel= openpyxl.load_workbook("CDR_CE_LT068_GSM_MODERNIZATION_ver1.xlsx", data_only=True)

while "".join(startColList)!="AL":
    if excel["BTS"]["".join(startColList)+str(2)].value!=None:
        colFill+=1
    startColList[-1]=chr(ord(startColList[-1])+1)
