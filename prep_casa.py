from distutils.command.sdist import sdist
import pandas as pd
import numpy as np
import xlwt
import re
import os

yol = os.getcwd()
yollu = re.sub("\\\\", "/", yol)+"/"
# print(yollu)
X = yollu + "PR-220301 MS 1.xlsx"
def prep_casa_csv():
    x = input("Dosya yolunu girin: ")
    y = input("Oluşacak yeni dosyanın adını girin: ")
    df_casa = pd.read_excel(yollu + x, dtype=str)
    df_casa_dict = df_casa.to_dict("index")
    casalar = {}
    big_casas = df_casa["SANDIK NO"].value_counts().to_dict()
    big_casas_list = [i for i, j in big_casas.items() if j > 10]
    # print(big_casas_list)
    for index, sandik in df_casa_dict.items():
        if sandik["SANDIK NO"] not in casalar.keys():
            temp_casa = {}
            temp_casa["Ürün Adı"] = sandik["ÜRÜN ADI İNGİLİZCE"]
            temp_casa["Adet"] = str(sandik["ADET"])
            casalar[sandik["SANDIK NO"]] = temp_casa
        else:
            casalar[sandik["SANDIK NO"]]["Ürün Adı"] += "VbCr" + str(sandik["ÜRÜN ADI İNGİLİZCE"])
            casalar[sandik["SANDIK NO"]]["Adet"] += "VbCr" + str(sandik["ADET"])
    for i in big_casas_list:
        a = casalar.pop(i)
        b = a["Ürün Adı"].split("VbCr")
        c = a["Adet"].split("VbCr")
        sd = len(b)+(10-(len(b)%10))
        d = [i for i in range(0, sd, 10)]
        counter = 1
        for j in d:
            casalar[i+f"-{counter}"] = {"Ürün Adı": "VbCr".join(b[j:j+10]), "Adet": "VbCr".join(c[j:j+10])}
            counter += 1     

    df_casalar = pd.DataFrame.from_dict(casalar, orient="index")
    df_casalar.sort_index()
    # print(df_casalar)
    df_casalar.to_csv(y+".csv", sep=";", index=True)
    print(f"Dosya {y+'.csv'} adıyla oluşturuldu.")

prep_casa_csv()