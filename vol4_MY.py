import pandas as pd
file_path = 'C:\\Users\\jacuu\\PycharmProjects\\projekt\\Analiza_danych\\Transformacja pliku z podprogramami — GIT\\losowe_dane.xlsx'
file = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
file = file.fillna('')
podprogram_name = file.iloc[0, 1:]
instructionPDF_path = file.iloc[1, 1:]
measuresPDF_path = file.iloc[2, 1:]
codeNumberPDF = []
for _, row in file.iloc[3:].iterrows():
    codeNumber = row.iloc[0]
    for i, value in enumerate(row[1:], start=1):
        value.lower()
        if value == "x":
            podprogram_name1 = podprogram_name[i]
            instructionPDF = instructionPDF_path[i]
            measuresPDF = measuresPDF_path[i]
            codeNumberPDF.append((codeNumber,podprogram_name1, instructionPDF, measuresPDF))


final_df = pd.DataFrame(codeNumberPDF, columns=['Kod detalu', 'Nazwa podprogramu', 'Ścieżka do instrukcji', 'Ścieżka do formatki pomiarowej'])
grouped = final_df.groupby('Kod detalu').cumcount() + 1
final_df['Numeracja'] = grouped

final_df_path = 'C:\\Users\\jacuu\\PycharmProjects\\projekt\\Analiza_danych\\Transformacja pliku z podprogramami — GIT\\final.xlsx'
final_df.to_excel(final_df_path, index=False)