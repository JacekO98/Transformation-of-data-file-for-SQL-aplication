import pandas as pd
from pathlib import Path


def load_data(input_path: Path) -> pd.DataFrame:
    """
    Load raw exel file
    """
    return pd.read_excel(input_path, 'Sheet1', header=None)


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean loaded DataFrame
    """
    df = df.fillna('').map(lambda x: str(x).strip())
    return df


def transform_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transform data, assignment underprogram name, and path to the part number
    """
    part_numbers = df.iloc[4:, 0]
    first_name = df.iloc[0, 1:]
    second_name = df.iloc[1, 1:]
    third_name = df.iloc[2, 1:]
    table = []

    for index, row in df.iloc[4:].iterrows():
        part_number = row.iloc[0]
        for index, number in enumerate(row, start=1):
            number.lower()
            if number == 'x':
                FName = first_name[index-1]
                SName = second_name[index-1]
                TName = third_name[index-1]
                table.append((part_number, FName, SName, TName))
    final_df = pd.DataFrame(table, columns=['Object Code', 'Subprogram Name', 'Path to instruction',
                                                    'Path to measurement format'])

    grouped = final_df.groupby('Object Code').cumcount() + 1
    final_df['Occurrences'] = grouped

    return final_df

def save_file(input_path: Path, output_path: Path) -> None:
    """
    Full transformation pipeline.
    """
    raw_data = load_data(input_path)
    cleaned_data = clean_data(raw_data)
    result = transform_data(cleaned_data)
    result.to_excel(output_path, index=False)

if __name__ == '__main__':
    BASE_DIR = Path(__file__).resolve().parent

    input_path = BASE_DIR / "input" / "data_input.xlsx"
    output_path = BASE_DIR / "output" / "file_for_SQL.xlsx"
    save_file(input_path, output_path)









