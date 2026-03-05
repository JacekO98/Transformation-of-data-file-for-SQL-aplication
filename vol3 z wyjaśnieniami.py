"""Program konwertujący plik z tabelą gdzie do korpusów za pomocą "x" przypisane są podprogramy
na plik gdzie jest to w dwóch kolumnach a do tego są kolumny z ścieżkami do instrukcji i formatki pomiarowej"""
import pandas as pd
# 1. import pandas
# 2. przypisanie ścieżki do pliku do zmiennej
# 3. przypisanie pliku do zmiennej za pomocą pd.read_excel
# 4. Wypełnienie pustych miejść pustymi stringami a następnie usunięcie pustych stringów
# 5. pobranie wartości z pierwszego wiersza. Doczytać co robi funkcja values
# Ścieżka do pliku Excel (Twojego pliku wejściowego)
xlsx_file_path_new = 'C:\\Users\\jacuu\\Desktop\\Matryca z podprogramami DW1419 do przerobienia\\DW1419 Programy do Apki.xlsx'

# Wczytanie arkusza Excela do obiektu DataFrame, nie zakładając nagłówków (header=None)
# sheet_name - nazwa arkusza, z którego pobieramy dane
data = pd.read_excel(xlsx_file_path_new, sheet_name='DW1419 Programy do apki', header=None)

# Krok 1: Usunięcie spacji i zastąpienie pustych komórek pustymi stringami
# fillna('') - zastępuje NaN pustym stringiem
# applymap(lambda x: str(x).strip()) - dla każdej komórki zmienia wartości na stringi "str(x) i  usuwa białe znaki strip(x)
data = data.fillna('').applymap(lambda x: str(x).strip())  #lambda to tzw. funkcja anonimowa – czyli funkcja jednowierszowa, która nie wymaga nazwy.
                                                           #lambda argument: działanie_na_argumencie


# Krok 2: Pobranie nazw podprogramów oraz dodatkowych numerów z odpowiednich wierszy
# Wiersz 1 (0 w Pythonie) zawiera nazwy podprogramów
subprogram_names = data.iloc[0, 1:].values  # Pomijamy pierwszą kolumnę (A), pobieramy od B dalej
print(subprogram_names)


# Wiersz 2 (index 1) zawiera wartości z drugiego wiersza
numbers_row_2 = data.iloc[1, 1:].values  # Pobieramy wartości od kolumny B dalej

# Wiersz 3 (index 2) zawiera wartości z trzeciego wiersza
numbers_row_3 = data.iloc[2, 1:].values  # Pobieramy wartości od kolumny B dalej

# Krok 3: Przypisywanie kodów korpusów do podprogramów i numerów
# Tworzymy pustą listę, która będzie przechowywać dane w postaci krotek
body_codes_with_subprograms = []

# Iterujemy przez wszystkie wiersze począwszy od piątego (index 4 w Pythonie)
for _, row in data.iloc[3:].iterrows():
    body_code = row.iloc[0]  # Pierwsza kolumna (kod korpusu)
    # Iteracja po pozostałych kolumnach (od drugiej kolumny)
    for i, cell in enumerate(row[1:], start=1):
        if cell.lower() == 'x':  # Jeśli w komórce znajduje się 'x' (oznacza przypisanie)
            subprogram_name = subprogram_names[i - 1]  # Pobieramy nazwę podprogramu
            number_2 = numbers_row_2[i - 1]  # Pobieramy numer z drugiego wiersza
            number_3 = numbers_row_3[i - 1]  # Pobieramy numer z trzeciego wiersza
            # Dodajemy dane jako krotkę do listy
            body_codes_with_subprograms.append((body_code, subprogram_name, number_2, number_3))

# Krok 4: Konwersja zebranych danych do nowego DataFrame
# Tworzymy DataFrame z kolumnami: 'Korpus Code', 'Subprogram', 'Number Row 2', 'Number Row 3'
body_codes_df = pd.DataFrame(body_codes_with_subprograms,
                             columns=['Korpus Code', 'Subprogram', 'Number Row 2', 'Number Row 3'])

# Krok 5: Dodanie kolumny numerującej wystąpienia dla każdego kodu korpusu
# groupby('Korpus Code') grupuje dane po kolumnie 'Korpus Code'
# cumcount() zwraca sekwencyjny numer wystąpienia dla każdego kodu (od 0), dlatego dodajemy +1
body_codes_df['Occurrences'] = body_codes_df.groupby('Korpus Code').cumcount() + 1

# Krok 6: Podgląd wyniku
print("Podgląd pierwszych 5 wierszy wynikowego DataFrame:")
print(body_codes_df.head())

# Krok 7: Zapisanie wynikowego DataFrame do nowego pliku Excel
# to_excel - zapisuje dane do pliku Excel, bez indeksów (index=False)
output_table_path = 'C:\\Users\\jacuu\\Desktop\\CC1210_korpusy_vol7.xlsx'
body_codes_df.to_excel(output_table_path, index=False)

print(f"Plik został zapisany pod ścieżką: {output_table_path}")
