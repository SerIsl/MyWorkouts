
from numpy import nan
import pandas as pd
import re
import os
from prep_casav2 import prep_casa_csv

dosya = input("Dosya adını giriniz: ")
satır_sayisi = int(input("Lütfen sayfalanacak satır sayısını giriniz: "))
dosya_adi = input("Oluşacak yeni dosyanın adını girin: ")

df = pd.read_excel(dosya, dtype=str)


sozluk = df.to_dict("records")

# dict_3 = sozluk[133]
# print(sozluk)

def cells(satir):
    malzeme_kodu = ""
    ürün_adı = ""
    ürün_adı_ing = ""
    mydict1 = []
    tmp = []
    
    for key, value in satir.items():
    
        tempdict = {}
        if type(value) == float:
            continue 

        if "CODE" in key:
            malzeme_kodu = value
        elif "TR" in key:
            ürün_adı = value
        elif "ENG" in key:
            ürün_adı_ing = value
        elif "SANDIK NO" in key:
            if value.isnumeric():
                tmp.append('M'+str(value))
            else:
                tmp.append(value)
        elif "ADET" in key:

            tempdict["SANDIK NO"] = tmp.pop()
            tempdict["MALZEME KODU"] = malzeme_kodu
            tempdict["ÜRÜN ADI"] = ürün_adı
            tempdict["ÜRÜN ADI İNGİLİZCE"] = ürün_adı_ing
            tempdict["ADET"] = value
            mydict1.append(tempdict)

    return mydict1    
            

a = list(map(cells, sozluk))

c = [x for b in a for x in b]
e =sorted(c, key=lambda x: int(x["SANDIK NO"][1:]))
d = sorted(e, key=lambda x: x["SANDIK NO"][0:1])

df_d = pd.DataFrame(d)
df_d.to_excel("nevV2"+dosya)
print("Dosya başarılı bir şekilde oluşturuldu.")

prep_casa_csv(df_d, satır_sayisi, dosya_adi)