#!/usr/bin/python
#-*- coding:utf-8 -*-
import pandas as pd
from glob import glob
from os import listdir, path, makedirs

"""
Tentei ao máximo não alterar a origem dos dados do site do Tesouro Direto (http://www.tesouro.gov.br/balanco-e-estatisticas)
Contudo, há datas mal formatadas em:
    Algumas planilhas de 2008 - Não registrei neste comentário logo que as alterei.
    A data de 23/08/2016 está em um formato diferente das demais da mesma coluna em todas as planilhas de 2016.
"""

def normalizeSheetName(sheetName):
    sheetName = sheetName.replace("NTN-B Princ ","NTN-B Principal ").replace("NTNBP","NTN-B Principal")
    sheetName = sheetName.replace("NTNB","NTN-B")
    sheetName = sheetName.replace("NTNC","NTN-C")
    sheetName = sheetName.replace("NTNF","NTN-F")
    return sheetName


def processFolder(folderPath, history):
    for fileName in glob(folderPath+"/*.xls"):
        print(fileName)
        xlsFile = pd.ExcelFile(fileName)
        for sheetName in xlsFile.sheet_names:
            df1 = pd.read_excel(xlsFile, sheetName, skiprows=1)
            sheetName = normalizeSheetName(sheetName)
            print(sheetName)
            if df1.columns.size == 5:
                df1 = df1.assign(PUBase = "") #Em 2002 eles não publicavam o PU_base_manha
            df1.columns = ["date","taxa_compra_manha","taxa_venda_manha","PU_compra_manha","PU_venda_manha","PU_base_manha"]

            # df1["Dia"] = pd.to_datetime(df1.Dia, dayfirst=True)
            try:
                df1["date"] = pd.to_datetime(df1.date, format="%d/%m/%Y")
            except ValueError:
                df1["date"] = pd.to_datetime(df1.date, format="%m/%d/%Y")

            if sheetName in history:
                history[sheetName] = pd.concat([df1, history[sheetName]])
            else:
                history[sheetName] = df1


def dumpProcessedData(savePath, processedData):
    if not path.exists(savePath):
        makedirs(savePath)

    for sheetName, df1 in processedData.items():
        print(sheetName)
        fileNameSave = sheetName.replace(" ","_") + ".csv"


        df1.index = df1.date
        df1 = df1.drop("date",axis=1)
        df1 = df1.sort_index()
        # df1 = df1.sort_values(by="date")
        print(df1)
        df1.to_csv(savePath+"/"+fileNameSave,encoding="utf8")
    processedData.clear()


if __name__ == "__main__":
    tdHistory = {}
    historyPath = "./td_history/"
    savePath = "./"
    for folder in listdir(historyPath):
        processFolder(historyPath+folder, tdHistory)
        dumpProcessedData(savePath+folder, tdHistory)
