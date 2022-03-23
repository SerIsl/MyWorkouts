import pandas as pd
import numpy as np
import xlwt
import re
import os

yol = os.getcwd()
yollu = re.sub("\\\\", "/", yol)+"/"
print(yollu)
# print(yollu)
# x = yollu + "PR-220301 MS 1.xlsx"
def prep_casa_csv():
    x = input("Dosya yolunu girin: ")
    y = input("Oluşacak yeni dosyanın adını girin: ")
    z = int(input("Lütfen sayfalanacak satır sayısını giriniz: "))
    df_casa = pd.read_excel(yollu + x, dtype=str)
    df_casa_dict = df_casa.to_dict("index")
    casalar = {}
    big_casas = df_casa["SANDIK NO"].value_counts().to_dict()
    big_casas_list = [i for i, j in big_casas.items() if j > z]
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
        if len(b)%z == 0:
            sd = len(b)
        else:
            sd = len(b)+(z-(len(b)%z))
        d = [i for i in range(0, sd, z)]
        counter = 1
        for j in d:
            casalar[i+f"-{counter}"] = {"Ürün Adı": "VbCr".join(b[j:j+z]), "Adet": "VbCr".join(c[j:j+z])}
            counter += 1     

    df_casalar = pd.DataFrame.from_dict(casalar, orient="index")
    df_casalar['C'] = df_casalar.index
    df_casalar['d'] = df_casalar['C'].str.extract(r'(\w)')
    df_casalar['e'] = df_casalar['C'].str.extract(r'(\d+)').astype(int)
    df_casalar['f'] = df_casalar['C'].str.extract(r'(?<=-)(\d+)').replace(np.nan, 0).astype(int)
    df_casalar.sort_values(["d", "e", "f"], inplace=True)
    df_casalar.drop(["C", "d", "e", "f"], axis=1, inplace=True)
    
    # print(df_casalar)
    df_casalar.to_csv(y+".csv", sep=";", index=True)
    print(f"Dosya {y+'.csv'} adıyla oluşturuldu.")

prep_casa_csv()