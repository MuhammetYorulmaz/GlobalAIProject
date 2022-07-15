import pandas as pd
import time
import datetime

gradeConversionTable = {
    "AA":[x for x in range(90,101)],
    "BA":[x for x in range(85,90)],
    "BB":[x for x in range(80,85)],
    "CB":[x for x in range(75,80)],
    "CC":[x for x in range(65,75)],
    "DC":[x for x in range(58,65)],
    "DD":[x for x in range(50,58)],
    "FD":[x for x in range(40,50)],
    "FF":[x for x in range(0,40)]
    }
# Öğrenci Not Bilgilerinin tutulduğu Excel Dosyası
try:
    data = pd.read_excel('Öğrenci Not Listesi.xlsx')
except FileNotFoundError:
    # print("Henüz Dosya oluşmadığından hata alacaktır.")
    pass

class OgrenciNotGirme():
    def __init__(self):
        self.calisma = True

    def program(self):
        secim = self.menuSecim()
        if secim ==1: 
            self.notGoruntuleme()
        if secim==2:
            self.yeniKayit()
        if secim == 3:
            self.programiKapat()
            
    def menuSecim(self):
        try: 
            secim = int(input("""****  Öğrenci Not Kayıt Sistemine Hoşgeldiniz ****\n\n1-Daha Önce Kaydedilen Notları Görüntüle\n2-Yeni Öğrenci-Not Bilgisi Ekle\n3-Programı Kapat\n\nSeçiminizi Giriniz:"""))
            while secim < 1 or secim > 3:
                secim = int(input("Lütfen 1 ile 3 arasında bir değer giriniz:"))
            return secim
        except:
            print("Lütfen 1 ile 3 arasında bir değer giriniz:")
        
    def notGoruntuleme(self):
        try:
            print("\nExcel Dosyasındaki Bilgiler")
            print(data.to_markdown(),"\n\n")
        except:
            print("Daha önce herhangi bir kayıt oluşturulmamış görünmekte. Yeni kayıt oluşturup Progamı kapatıp-açtıktan sonra ilgili veriler görüntülenebilecektir.")
        
    def yeniKayit(self):
        dersList=[]
        ogrAdList=[]
        ogrSoyadList=[]
        ogrNumList=[]
        ogrNotList=[]
        ogrNotHarfList=[]
        ogrDurumList=[]
        
        try:
            #Yeni gelen verileri Excel Dosyasının altına ekleyebilmek için Exceldeki verileri listelere ekleyelim. Excel Dosyasını Veri Tabanı gibi kullanalım..
            dersList.extend(data["Ders"].values)
            ogrAdList.extend(data["Öğrencinin Adı"].values)
            ogrSoyadList.extend(data["Öğrencinin Soyadı"].values)
            ogrNumList.extend(data["Öğrencinin Numarası"].values)
            ogrNotList.extend(data["Öğrencinin Notu"].values)
            ogrNotHarfList.extend(data["Not Harf Bilgisi"].values)
            ogrDurumList.extend(data["Öğrencinin Durumu"].values)
        except NameError:
            pass
            
        def dersSecim():
            try:
                ders = input("Ders Bilgisini Giriniz:(Programlama,Veri Yapıları gibi.) :")
                return ders
            except:
                print("Tekrar deneyiniz. Örneğin: Fizik,Matematik :")
                dersSecim()
                
        def kayitSystem():

            def ogrenciNot():
                global ogrNot
                try:
                    ogrNot = float(input("Öğrencinin Aldığı Notu Giriniz [0-100 arasında olmalıdır]:"))
                    if round(ogrNot) not in range(0,101):
                        print("Girilen Notlar 0 ile 100 arasında olmalıdır.")
                        ogrenciNot()
                except ValueError:
                    print("Not alanında olduğunuzdan sayısal bir değer girmeniz gerekmektedir.\nÖrneğin:'80' veya '80.5' gibi.")
                    ogrenciNot()
            def ogrenciAd():
                global ogrAd
                try: 
                    ogrAd = input("Öğrencinin Adını Giriniz:")
                    if ogrAd.strip() == "":
                        print("Öğrencinin adı boş bırakılmamalıdır.")
                        ogrenciAd()
                except ValueError:
                    print("Ad Giriniz:\nÖrnek:'Muhammet'")
                    ogrenciAd()
            def ogrenciSoyad():
                global ogrSoyad
                try: 
                    ogrSoyad = input("Öğrencinin Soyadını Giriniz:")
                    if ogrSoyad.strip() == "":
                        print("Soyad alanı boş bırakılmamalıdır.")
                        ogrenciSoyad()
                except ValueError:
                    print("Soyad Giriniz:\nÖrnek:'Yorulmaz'")
                    ogrenciSoyad()
            def ogrenciNumara():
                global ogrNum
                try: 
                    ogrNum = int(input("Öğrencinin Numarasını Giriniz:"))
                except:
                    print("Öğrenci numara alanında olduğunuzdan sayısal tam bir değer girmeniz gerekmektedir.\nÖrneğin:'217605021'")
                    ogrenciNumara()
                
            #Call Func.       
            ogrenciAd()
            ogrenciSoyad()
            ogrenciNumara()
            ogrenciNot()
            #Add List
            ogrAdList.append(ogrAd.strip().title())
            ogrSoyadList.append(ogrSoyad.strip().upper())
            ogrNotList.append(ogrNot)
            ogrNumList.append(ogrNum)
            
            # Harf ve Durum kontrol
            for key,value in gradeConversionTable.items():
                if round(ogrNot) in value:
                    ogrNotHarfList.append(key)
                    if key in ["AA","BA","BB","CB","CC","DC","DD"]:
                        ogrDurumList.append("Geçti")
                    else:
                        ogrDurumList.append("Kaldı")
        while True:
            ders = dersSecim()
            kayitSystem()
            dersList.append(ders.strip().title())
            while True:
                devamDurum = input(f"{ders.title()} Dersi İçin Kayıt Ettirme işlemi Devam Ettirilsin mi? E/H :")
            # devamDurum = input("Kayıt Ettirme işlemi Devam Ettirilsin mi? E/H :")
                if devamDurum.upper() =="E":
                    kayitSystem()
                    dersList.append(ders.strip().title())
                else:
                    devamDurum2 = input("Başka Bir Ders İçin Kayıt Ettirme işlemi Devam Ettirilsin mi? E/H :")
                    if devamDurum2.upper() =='E':
                        ders = dersSecim()
                        kayitSystem()
                        dersList.append(ders.strip().title())
                    else:
                        break
            break
                
        dfList = [dersList,ogrNumList,ogrAdList,ogrSoyadList,ogrNotList,ogrNotHarfList,ogrDurumList]
        dfOgrenciBilgi = pd.DataFrame(dfList).transpose()
        dfOgrenciBilgi.columns = ['Ders','Öğrencinin Numarası','Öğrencinin Adı','Öğrencinin Soyadı','Öğrencinin Notu','Not Harf Bilgisi','Öğrencinin Durumu']
        
        # print(dfOgrenciBilgi.to_markdown())
        def dfToExcel(dfOgrenciBilgi):
            try: 
                with pd.ExcelWriter("Öğrenci Not Listesi.xlsx") as writer:  
                    dfOgrenciBilgi.to_excel(writer,index=False,sheet_name='Öğrenci Not Listesi')
                print(str(datetime.datetime.now().strftime("%x %X")),' Öğrenci Not Listesi için hazırlanan Excel Dosyası başarılı bir şekilde işlenmiştir.')
            except Exception as E:
                print("Excel Dosyası açık kalmış olmalı. Dsoyayı kapattıktan bir süre bekler misiniz? ",E)
                time.sleep(60)
                dfToExcel(dfOgrenciBilgi)
        dfToExcel(dfOgrenciBilgi)
        
    def programiKapat(self):
        self.calisma = False
        print("---Oturum Sonlandırılmıştır---")

sistemKontrol = OgrenciNotGirme()     
while sistemKontrol.calisma:
    sistemKontrol.program()
    


