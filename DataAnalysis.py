import openpyxl
import os
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures
from sklearn.pipeline import make_pipeline

def get_second_column(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Data']
    second_column_values = [cell.value for cell in sheet['B'][1:]]
    workbook.close()
    return second_column_values

data_files = [file for file in os.listdir() if file.startswith("Data") and file.endswith(".xlsx")]

all_models = []  # Her veri seti için bir modeli saklamak için liste

# Her bir veri seti için ayrı bir regresyon modeli oluştur
for data_file in data_files:
    second_column_values = get_second_column(data_file)
    X = np.array(second_column_values[:-1]).reshape(-1, 1)
    Y = np.array(second_column_values[1:]).reshape(-1, 1)

    # Polynomial regresyon kullanarak modeli oluştur
    model = make_pipeline(PolynomialFeatures(degree=3), LinearRegression())
    model.fit(X, Y)

    all_models.append(model)

# Tahmin edilecek seti oluştur (örneğin, en son veri seti)
tahmin_edilecek_set = np.array(get_second_column(data_files[-1])).reshape(-1, 1)

# Tahmin edilen seti diğer veri setlerine uygun boyuta getir
for i in range(len(data_files) - 1):
    current_model = all_models[i]
    current_set_length = len(get_second_column(data_files[i])) - 1
    predicted_set_length = len(tahmin_edilecek_set)

    while predicted_set_length < current_set_length:
        # Tahminlerde bulun
        tahmin = current_model.predict(tahmin_edilecek_set[-1].reshape(-1, 1))
        tahmin_edilecek_set = np.concatenate([tahmin_edilecek_set, tahmin], axis=0)
        predicted_set_length += 1

# Tahmin edilen seti son veri setine ekle
son_veri_seti = tahmin_edilecek_set

print("Son Veri Seti: İlk 10")
for i in son_veri_seti[:10]:
    print(i)

print("Son Veri Seti: Son 10")
for i in son_veri_seti[-10:]:
    print(i)

print(len(son_veri_seti))
