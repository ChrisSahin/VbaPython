import os
import sys
import win32com.client


if __name__ == "__main__":

    #xl = win32com.client.Dispatch("Excel.Application") Fonctionne bien si on est sur d'être sur le fichier Excel actif
    ExcelFile = sys.argv[1] 
    print(ExcelFile)
    xl = win32com.client.GetObject(ExcelFile) #Pour être sur d'écrire dans le bon fichier excel, au cas où plusieurs sessions excel seraient ouvertes en même temps 
    ws = xl.Worksheets("Sheet1")
    ListArg = list(list(ws.Range("A2:E2").value)[0])
    print(ListArg)
    #map_object = map(int, ListArg)
    #listInteger = list(map_object)
    sumValue = sum(ListArg)
    print(sumValue)
    ws.Range("F2").value = sumValue
    sys.exit(0) 