import click
import pandas as pd
import random
import time

@click.command()
@click.option('--liczba_osob',    prompt='Podaj liczbe osob do wylosowania', type=int)
@click.option('--plik_wejsciowy', prompt='Podaj sciezke pliku z lista osob', type=click.Path(exists=True))

def shuffle_and_save(liczba_osob, plik_wejsciowy):
    """
    Reads an Excel file, shuffles its rows, and saves the first specified number of rows to a new Excel file.
    
    Parameters:
    - num_rows: The number of rows to save after shuffling.
    - input_file: The path to the input Excel file.
    - output_file: The path to the output Excel file.
    """
    # Read the Excel file without headers
    df = read_excel_filter_ids(plik_wejsciowy)

    # Wylosuj numery osob (z zachowaniem jednostajnego rozkladu prawdopodobienstwa)
    wylosowane_numery = random.sample(range(0, df.index.size), liczba_osob)

    # Wybierz z tabeli osoby o wybranych numerach
    wylosowane_osoby = df.iloc[wylosowane_numery].sort_index()
    
    # Wyswietl wyniki:
    click.echo("========================================")
    click.echo("Lista wylosowanych osob:")
    click.echo(wylosowane_osoby.to_string(index=False))

    random.seed
    # Zapisz posortowaną listę wybranych osób
    plik_wynikowy = "WynikiLosowania_"+plik_wejsciowy
    with pd.ExcelWriter(plik_wynikowy) as writer:
        wylosowane_osoby.to_excel(writer, sheet_name=plik_wynikowy, index=False)
        
    
    # Wyswietl wyniki i czeka
    click.echo(f"Zapisano {liczba_osob} wylosowanych osob do pliku: {plik_wynikowy}")
    #input("Wcisnij enter aby wyjsc...")

def read_excel_filter_ids(file_path):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)
    
    # Assuming the first column is the ID column
    id_column = df.columns[0]
    
    # Convert the ID column to numeric, forcing errors to NaN
    df[id_column] = pd.to_numeric(df[id_column], errors='coerce')
    
    # Exclude rows where the ID is NaN
    df = df.dropna(subset=[id_column])
    
    # Optionally, convert the ID column back to integers (if IDs are supposed to be integers)
    df[id_column] = df[id_column].astype(int)
    
    return df

# pyinstaller --onefile --noconsole main.py
if __name__ == '__main__':
    shuffle_and_save()
