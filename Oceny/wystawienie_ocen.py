# coding: utf-8
import pandas as pd

def ocena(x):
    if x < .4: return 'niedostateczny: 1'
    if x < .56: return 'dopuszczający: 2'
    if x < .71: return 'dostateczny: 3'
    if x < .86: return 'dobry: 4'
    if x < 1: return 'bardzo dobry: 5'
    return 'celujący: 6'

files = glob.glob(r'*.xlsx')  # Lista plików Excel

for file in files:
    df = pd.read_excel(file, skiprows=1, usecols='D:H, K, M:N')
    df = df[df['Stan'] == 'Zwrócono']

    # Wyodrębnienie unikalnych nazw zadań
    unique_tasks = df['Zadania'].unique()
    
    # Dynamiczne tworzenie mapowania
    mapping = {}
    test_counter = 1
    for task in unique_tasks:
        if 'PK1' in task:
            mapping[task] = 'PK1'
        elif 'PK2' in task:
            mapping[task] = 'PK2'
        else:
            mapping[task] = f'test_{test_counter:02d}'
            test_counter += 1

    df['Zadanie'] = df['Zadania'].map(mapping)

    pivot_df = df.pivot_table(index='Imię i nazwisko', columns='Zadanie', values='Punkty', aggfunc='sum')
    pivot_df['Suma zdobytych punktów'] = pivot_df.sum(axis=1)
    max_points = df.groupby('Zadanie')['Maksymalna liczba punktów'].max().sum()
    pivot_df['Średnia'] = pivot_df['Suma zdobytych punktów'] / max_points
    pivot_df['Ocena'] = pivot_df['Średnia'].apply(ocena)

    email_df = df[['Imię i nazwisko', 'Adres e-mail']].drop_duplicates('Imię i nazwisko')
    pivot_df = pivot_df.merge(email_df, on='Imię i nazwisko', how='left')

    # Przekształcenie indeksu na kolumnę i sortowanie po nazwisku
    pivot_df.reset_index(inplace=True)
    pivot_df['Nazwisko'] = pivot_df['Imię i nazwisko'].str.split().str[-1]
    pivot_df.sort_values(by='Nazwisko', inplace=True)
    pivot_df.drop(columns='Nazwisko', inplace=True)

    # Zapisanie wyników do nowego pliku Excel
    pivot_df.to_excel('oceny_' + file, index=False)
