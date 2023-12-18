# coding: utf-8
import pandas as pd
import glob
import os

# Wczytanie listy plików Excel
files = glob.glob(r'*.xlsx')

for file in files:
    # Wczytanie pliku
    df = pd.read_excel(file)

    # Zliczenie wystąpień 'V' w każdej kolumnie z datą, zastępując NaN pustymi stringami
    v_counts = df.iloc[:, 3:].fillna('').apply(lambda x: x.str.count('V').sum())

    # Ustalenie progowego limitu dla kolumn do zachowania
    # Wybieramy kolumny z top 3-5 największą liczbą 'V'
    threshold = v_counts.nlargest(5).min() - 1

    # Usunięcie kolumn, które nie spełniają progu
    for col in df.columns[3:]:
        if v_counts[col] < threshold:
            df.drop(col, axis=1, inplace=True)

    # Obliczenie frekwencji dla każdego ucznia
    # Zakładamy, że 'V' oznacza obecność
    df['Frekwencja'] = df.iloc[:, 3:].applymap(lambda x: 1 if x == 'V' else 0).sum(axis=1)

    # Tworzenie nowej nazwy pliku z dodanym słowem 'frekwencja'
    new_file_name = os.path.splitext(file)[0] + '_frekwencja.xlsx'

    # Zapisanie wyników do nowego pliku Excel
    df.to_excel(new_file_name, index=False)
