import openpyxl
import os
import numpy as np
from sklearn.linear_model import LinearRegression

def get_second_column(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Data']

    # İkinci sütunun ikinci satırından sonraki değerleri al
    second_column_values = [cell.value for cell in sheet['B'][1:]]

    workbook.close()
    return second_column_values

# Ana dizindeki Data dosyalarını bul
data_files = [file for file in os.listdir() if file.startswith("Data") and file.endswith(".xlsx")]

# Tüm ikinci sütun değerlerini içerecek liste
all_second_columns = []

# Her dosyayı aç, ikinci sütunu al ve listeye ekle
# Her bir dosyayı aç, ikinci sütunu al ve listeye ekle
for data_file in data_files:
    second_column_values = get_second_column(data_file)
    all_second_columns.append(second_column_values)
    set_values = second_column_values
# Tahmin edilen seti son veri setine ekle


# all_second_columns listesi içindeki listeleri birleştir
X = np.concatenate(all_second_columns)

# Her bir satırın ilk elemanını bağımlı değişken Y olarak al
Y_original = X[:-1]  # Son eleman hariç
Y_to_predict = X[-1]  # Son eleman, tahmin edilecek setin gerçek değeri

# Bağımlı değişkeni (Y_original) ve bağımsız değişkeni (X) kullanarak modeli oluştur
model = LinearRegression()
model.fit(X[:-1].reshape(-1, 1), Y_original)

# 21. setin tahminini yap
tahmin_edilecek_set = np.array(set_values[1:])


tahmin = model.predict(tahmin_edilecek_set.reshape(-1, 1))

# Tahmin edilen seti son veri setine ekle
son_veri_seti = np.concatenate([all_second_columns[-1], [tahmin[-1]]], axis=0)

# Son veri setini ekrana yazdır

son_veri_seti = (all_second_columns[:-1], [tahmin_edilecek_set])

print("Son Veri Seti:")
print(son_veri_seti)
for i in son_veri_seti[1]:
    print(i)

