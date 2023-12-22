import openpyxl
import os
import numpy as np
from sklearn.linear_model import LinearRegression


def get_second_column(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Data']
    second_column_values = [cell.value for cell in sheet['B'][1:]]
    workbook.close()
    return second_column_values

data_files = [file for file in os.listdir() if file.startswith("Data") and file.endswith(".xlsx")]

all_second_columns = []



# Her dosyayı aç, ikinci sütunu al ve listeye ekle
for data_file in data_files:
    second_column_values = get_second_column(data_file)
    all_second_columns.append(second_column_values)



# Tüm ikinci sütun değerlerini içerecek liste
X = np.concatenate(all_second_columns)



# Her bir satırın ilk elemanını bağımlı değişken Y olarak al
Y_original = X[:-1]  # Son eleman hariç
Y_to_predict = X[1:]  # İlk eleman hariç, tahmin edilecek setin gerçek değeri



# Bağımlı değişkeni (Y_original) ve bağımsız değişkeni (X) kullanarak modeli oluştur
model = LinearRegression()



# Eğitici verilerle modeli eğit
model.fit(Y_original.reshape(-1, 1), Y_to_predict.reshape(-1, 1))



# Tahmin edilecek seti oluştur
tahmin_edilecek_set = Y_to_predict.reshape(1, -1)



# Tahminleri gerçekleştir
tahminler = []
step = 100  # Her adımda bir güncelleme yap
for i in range(0, len(tahmin_edilecek_set[0]), step):
    tahmin = model.predict(tahmin_edilecek_set[0][:i+1].reshape(-1, 1))
    tahminler.append(tahmin[-1])



# Tahmin edilen seti son veri setine ekle
son_veri_seti = np.concatenate([Y_to_predict[:len(tahminler)].reshape(-1, 1), np.array(tahminler).reshape(-1, 1)], axis=0)

print("Son Veri Seti:")
for i in son_veri_seti:
    print(i)


print(len(son_veri_seti))

print("Model Coef:", model.coef_)
print("Model Intercept:", model.intercept_)



