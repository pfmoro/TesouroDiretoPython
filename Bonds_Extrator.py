# -*- coding: utf-8 -*-
"""
Created on Thu Dec 29 09:20:13 2022

@author: PC
"""
import requests
import pandas as pd
from bs4 import BeautifulSoup
import time, os
import re
import pyexcel as p
from openpyxl.reader.excel import load_workbook
path = "./bonds_excel"

#valores nominais de NTNB
text=BeautifulSoup(requests.get("https://sisweb.tesouro.gov.br/apex/f?p=2501:9::::9:P9_ID_PUBLICACAO:28715").content).text
text=(text[text.find("window.location =")+len("window.location ="):text.find(".xlsx")+5]).replace(" ","").replace('"','')

df=pd.read_excel(text,engine="openpyxl",header=8,skiprows=0).dropna()

def get_filename_from_cd(cd):
    """
    Get filename from content-disposition
    """
    if not cd:
        return None
    fname = re.findall('filename="(.+)";', cd)
    if len(fname) == 0:
        return None
    return fname[0]


year=str(time.localtime().tm_year)

urls=[]
titulos=["LTN","NTN-B","NTN-B_Principal","NTN-C","NTN-F"]
for titulo in titulos:
    url="https://cdn.tesouro.gov.br/sistemas-internos/apex/producao/sistemas/sistd/"+year+"/"+titulo+"_"+year+".xls"
    urls.append(url)
index=0
for url in urls:
    try:
    	r = requests.get(url, allow_redirects=True)
    	filename = titulos[index]+"_"+year+".xls"
    	open(os.path.join(path, filename), 'wb').write(r.content)
    	time.sleep(0.2)
    	print("file "+filename+" has been downloaded.")
    except Exception as e:
    	print(e)
    index=index+1
    
dfs = []
df3=pd.DataFrame()
from os.path import abspath

for file in os.listdir(abspath(path)):
    if file.endswith(".xls"):
        full_path = os.path.join(abspath(path), file)
        p.save_book_as(file_name=full_path,dest_file_name=os.path.join(abspath(path), file+"x"))

for file in os.listdir(abspath(path)):
    if file.endswith(".xlsx"): # and "NTN-B" in file:
        full_path = os.path.join(abspath(path), file)
        wb=load_workbook(full_path)
        sheetList=[]
        for sheet in wb:
            sheet["B2"].value="Taxa Compra Abertura"
            sheet["C2"].value="Taxa Venda Abertura"
            sheet["D2"].value="PU Compra Abertura"
            sheet["E2"].value="PU Venda Abertura"
            sheet["F2"].value="Papel"
            sheet["G2"].value="Vencimento"
            for row in range(3,len(sheet['A'])):
                if sheet["A" + str(row)].value != "":
                    sheet["G" + str(row)].value = sheet["B1"].value
                    sheet["F" + str(row)].value = sheet.title
            sheetList.append(sheet.title)
        wb.save(full_path)
        xls = pd.ExcelFile(full_path)
        for sheetname in sheetList:
            data = pd.read_excel(xls, sheet_name = sheetname, header=1)
            data.dropna(how="any")
            df3=pd.concat([df3,data])

df3['Dia']=pd.to_datetime(df['Dia'],errors='coerce')
#df3.to_pickle("Histórico de preços e taxas.pickle")
df3.to_excel("Histórico de preços e taxas.xlsx")
# df3.to_parquet("Histórico de preços e taxas.parquet")
#df3.to_msgpack("Histórico de preços e taxas.msgpack")
