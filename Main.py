import math
import openpyxl
import random
from random import randint
from datetime import datetime, timedelta
import os

baslangic_tarihi = datetime(2023, randint(1, 3), randint(1, 28), randint(1, 22))


def olustur_ve_aktif_et(dosya_adi):
    # Dosya var mı kontrol et
    if not os.path.exists(dosya_adi):
        # Dosya yoksa, yeni bir Excel dosyası oluştur
        workbook = openpyxl.Workbook()
        workbook.save(dosya_adi)
        return dosya_adi
    else:
        # Dosya varsa, dosya adını incele
        dosya_sayisi = 1
        while f"Data{dosya_sayisi}.xlsx" in os.listdir():
            dosya_sayisi += 1

        # Yeni dosya adını oluştur
        yeni_dosya_adi = f"Data{dosya_sayisi}.xlsx"

        # Yeni dosyayı oluştur
        workbook = openpyxl.Workbook()
        workbook.save(yeni_dosya_adi)

        return yeni_dosya_adi


def addData(pageName, firstStep, secondStep, startSpeed, endSpeed, sensitivity):
    for row in range(firstStep, secondStep):
        randomValue = random.uniform(startSpeed, endSpeed) + row / sensitivity
        speedChange = randomValue * 2 * math.pi
        energy = 0.5 * (0.5 * 180 * 0.2 * 0.2) * speedChange * speedChange
        pageName.cell(row=row, column=1, value=str(row - 1) + ". Veri")
        pageName.cell(row=row, column=2, value=randomValue)
        pageName.cell(row=row, column=3, value=speedChange)
        pageName.cell(row=row, column=4, value=energy)
        pageName.cell(row=row, column=5, value=getTime(randomValue))


def getTime(time):
    global baslangic_tarihi

    # 250'de 1 ihtimal gerçekleşirse 1,2 veya 3 saat molaver
    number = randint(1, 250)
    if (number == 150):
        baslangic_tarihi += timedelta(hours=randint(1, 4))
        if baslangic_tarihi.hour > 17:
            baslangic_tarihi += timedelta(hours=15)

    # Cumartesi günü ise günü 2 artır
    if baslangic_tarihi.weekday() == 5:
        baslangic_tarihi += timedelta(days=2)
    # Pazar günü ise günü 1 artır
    elif baslangic_tarihi.weekday() == 6:
        baslangic_tarihi += timedelta(days=1)

    # Saat 17:00 ise saati 15 saat ekleyerek 8:00'e getir
    if baslangic_tarihi.hour == 17:
        baslangic_tarihi += timedelta(hours=15)

    milis = time * 1000
    baslangic_tarihi += timedelta(milliseconds=milis)

    return baslangic_tarihi


def generateAndSaveData(entire=True):
    fileName = "Data1.xlsx"

    fileName = olustur_ve_aktif_et(fileName)

    workbook = openpyxl.load_workbook(fileName)

    yeni_sayfa = workbook.create_sheet("Data")
    workbook.active = yeni_sayfa

    # Başlık ekle
    yeni_sayfa['A1'] = "VERİ"
    yeni_sayfa['B1'] = 'DEVİR DEĞİŞİMİ'
    yeni_sayfa['C1'] = 'HIZ DEĞİŞİMİ'
    yeni_sayfa['D1'] = 'ENERJİ'
    yeni_sayfa['E1'] = 'ZAMAN DAMGASI'

    firstStep = randint(70000, 85000)
    addData(yeni_sayfa, 2, firstStep, 1.1, 1.4, 800000)
    if entire==False:
        workbook.save(fileName)
        return


    secondStep = firstStep + randint(10000, 18000)
    addData(yeni_sayfa, firstStep, secondStep, 1.36, 1.43, 900000)

    thirdStep = secondStep + randint(6000, 10000)
    addData(yeni_sayfa, secondStep, thirdStep, 1.41, 1.47, 1000000)

    forthStep = thirdStep + randint(3000, 15000)
    addData(yeni_sayfa, thirdStep, forthStep, 1.46, 1.54, 500000)

    # Excel dosyasını kaydet
    workbook.save(fileName)






for i in range(1):
    generateAndSaveData(True)
    baslangic_tarihi = datetime(2023, randint(1, 3), randint(1, 28), randint(1, 22))