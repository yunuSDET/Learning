import openpyxl
import os
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import MinMaxScaler

def get_second_column(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Data']
    second_column_values = [cell.value for cell in sheet['B'][1:]]
    workbook.close()
    return second_column_values

data_files = [file for file in os.listdir() if file.startswith("Data") and file.endswith(".xlsx")]

# Tüm veri setleri için bir regresyon modeli oluştur
all_models = []

for data_file in data_files:
    second_column_values = get_second_column(data_file)
    X = np.array(second_column_values[:-1]).reshape(-1, 1)
    Y = np.array(second_column_values[1:]).reshape(-1, 1)

    # Polynomial regresyon kullanarak modeli oluştur
    model = make_pipeline(PolynomialFeatures(degree=3), LinearRegression())
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
    artis_orani = np.random.normal(loc=1, scale=0.1)  # Rastgele bir artış oranı seç
    eğilim_tahmini_normalized *= artis_orani

    # Tahmini normalize edilmiş değerleri gerçek değerlere çevir
    eğilim_tahmini = scaler.inverse_transform(eğilim_tahmini_normalized.reshape(-1, 1))

    tahmin_edilecek_set = np.concatenate([tahmin_edilecek_set, eğilim_tahmini], axis=0)

print("Tahmin Edilen Veri Seti: İlk 10")
for i in tahmin_edilecek_set[:10]:
    print(i)

print("Tahmin Edilen Veri Seti: Son 10")
for i in tahmin_edilecek_set[-10:]:
    print(i)

print(len(tahmin_edilecek_set))
