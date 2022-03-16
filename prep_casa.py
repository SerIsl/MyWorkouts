import pandas as pd
import numpy as np
import xlwt
import re
import os

yol = os.getcwd()
yollu = re.sub("\\\\", "/", yol)+"/"
# print(yollu)

def prep_casa_csv():
    x = input("Dosya yolunu girin: ")
    y = input("Oluşacak yeni dosyanın adını girin: ")
    df_casa = pd.read_excel(yollu+x, dtype=str)
    df_casa_dict = df_casa.to_dict("index")
    casalar = {}
    for index, sandik in df_casa_dict.items():
        if sandik["SANDIK NO"] not in casalar.keys():
            temp_casa = {}
            temp_casa["Ürün Adı"] = sandik["ÜRÜN ADI İNGİLİZCE"]
            temp_casa["Adet"] = str(sandik["ADET"])
            casalar[sandik["SANDIK NO"]] = temp_casa
        else:
            casalar[sandik["SANDIK NO"]]["Ürün Adı"] += "VbCr" + str(sandik["ÜRÜN ADI İNGİLİZCE"])
            casalar[sandik["SANDIK NO"]]["Adet"] += "VbCr" + str(sandik["ADET"])

    df_casalar = pd.DataFrame.from_dict(casalar, orient="index")
    df_casalar.to_csv(y+".csv", sep=";", index=True)
    print(f"Dosya {y+'.csv'} adıyla oluşturuldu.")

prep_casa_csv()