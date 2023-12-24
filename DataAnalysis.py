import openpyxl
import os
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures, MinMaxScaler
from sklearn.pipeline import make_pipeline
import matplotlib.pyplot as plt

import Main


def show():
    sampleList = [file for file in os.listdir() if file.startswith("Data") and file.endswith(".xlsx")][0]
    second_column_values = get_second_column(sampleList)
    #sayi_listesi = [eleman[0] for eleman in second_column_values]

    Main.updateLastFileWithEstimatedData(second_column_values)
    plt.plot(second_column_values, marker='o', linestyle='-', color='b', label='Sayı Listesi')
    plt.title('Örnek Çizgi Grafiği')
    plt.xlabel('Index')
    plt.ylabel('Değerler')
    plt.legend()  # Eğer bir etiket varsa legend'ı görüntüle
    plt.grid(True)  # Izgara ekle

    # Grafiği göster
    plt.show()

def get_second_column(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Data']
    second_column_values = [cell.value for cell in sheet['B'][1:]]
    workbook.close()
    return second_column_values


def doAnalysis():
    data_files = [file for file in os.listdir() if file.startswith("Data") and file.endswith(".xlsx")]


    # Tüm veri setleri için bir regresyon modeli oluştur
    all_models = []

    for data_file in data_files:
        second_column_values = get_second_column(data_file)
        X = np.array(second_column_values[:-1]).reshape(-1, 1)
        Y = np.array(second_column_values[1:]).reshape(-1, 1)

        # Polynomial regresyon kullanarak modeli oluştur
        model = make_pipeline(PolynomialFeatures(degree=2), LinearRegression())
        model.fit(X, Y)
        all_models.append(model)

    # Son veri setinin eğilimini taklit ederek devam et
    last_model = all_models[-1]
    tahmin_edilecek_set = np.array(get_second_column(data_files[-1])).reshape(-1, 1)

    # Min-max normalizasyon uygula
    scaler = MinMaxScaler()
    tahmin_edilecek_set_normalized = scaler.fit_transform(tahmin_edilecek_set)

    # Yeni veri noktaları ekleyerek seti genişlet
    for i in range(len(tahmin_edilecek_set), len(tahmin_edilecek_set) + 35000):
        # Eğilim tahmini yap
        eğilim_tahmini_normalized = last_model.predict(tahmin_edilecek_set_normalized[-1].reshape(-1, 1))

        # Artışları daha detaylı taklit et
        artis_orani = np.random.normal(loc=1, scale=0.05)  # Rastgele bir artış oranı seç
        eğilim_tahmini_normalized *= artis_orani

        # Tahmini normalize edilmiş değerleri gerçek değerlere çevir
        eğilim_tahmini = scaler.inverse_transform(eğilim_tahmini_normalized.reshape(-1, 1))

        tahmin_edilecek_set = np.concatenate([tahmin_edilecek_set, eğilim_tahmini], axis=0)

    print("Tahmin Edilen Veri Seti: İlk 10")
    print(tahmin_edilecek_set[:10])

    print("Tahmin Edilen Veri Seti: Son 10")
    print(tahmin_edilecek_set[-10:])

    print("Estimated data size is: " + str(len(tahmin_edilecek_set)))

    question = input("Do you want to add estimated data into the data set? (y/n)")
    if question == "y" or question == "Y":
        sayi_listesi = [eleman[0] for eleman in tahmin_edilecek_set]
        Main.updateLastFileWithEstimatedData(sayi_listesi)
        plt.plot(sayi_listesi, marker='o', linestyle='-', color='b', label='Sayı Listesi')
        plt.title('Örnek Çizgi Grafiği')
        plt.xlabel('Index')
        plt.ylabel('Değerler')
        plt.legend()  # Eğer bir etiket varsa legend'ı görüntüle
        plt.grid(True)  # Izgara ekle

        # Grafiği göster
        plt.show()

        print("Analysis is completed and new data is saved")
    else:
        print("Analysis is completed but new data is not saved")


