import pandas as pd
import os
from os.path import exists
from datetime import datetime

def create_offer():
    source_filename = input("Wklej ścieżkę do pliku, z którego chcesz wziąć dane\n")
    if not exists(source_filename):
        raise ValueError("Ten plik nie istnieje")
    filename_to_create = input("Podaj nazwę pliku, jaki chcesz utworzyć\n")
    discount = int(input("Podaj rabat (od 0 do 100)\n"))
    if not 0 <= discount <= 100:
        raise Exception("Rabat musi wynosić między 0 a 100")
    keyword = input("Podaj słowo kluczowe\n")
    look_for_match_in_whole_name = input("Podaj czy słowo kluczowe ma być szukane na początku nazwy produktu, czy w całej nazwie\nPoczątek -> Wpisz 'P'\nCała nazwa -> Wpisz 'C'\n")
    if look_for_match_in_whole_name not in ["c", "C", "P", "p"]:
        raise ValueError("Wpisz literkę 'p' lub 'c")
    elif look_for_match_in_whole_name in ['c', "C"]:
        look_for_match_in_whole_name = True
    else:
        look_for_match_in_whole_name = False
    low_amount = int(input("Podaj stan poniżej którego, można uznać że zostało niewiele tego produktu\n"))
    if low_amount <= 0:
        raise Exception("Próg oznaczający małą ilość towaru musi być większy od 0")
    discount = 1 - (discount / 100) 

    filename_to_create = filename_to_create.split('.')[0]
    if exists(filename_to_create + ".html"):
        print("Plik o ten nazwie już istnieje - do pliku dodana zostanie data")
        filename_to_create += "-" + datetime.now().strftime("%d/%m/%Y %H:%M").replace("/", "_").replace(" ", "-")
    result = []
    df = pd.read_excel(source_filename)
    columns = df.columns
    keyword = keyword.upper()
    for name, amount_left, unit, price in zip(df[columns[0]], df[columns[1]], df[columns[2]], df[columns[3]]):
        if amount_left != 0:
            name = name.replace(",", '.')
            amount_left = "TAK" if amount_left > low_amount else "NIE"
            if look_for_match_in_whole_name:
                if keyword in name.upper():
                    result.append(f"{name}, {unit}, {round(price * discount, 2)}, {amount_left}\n")
            else:
                if name.upper().startswith(keyword):
                    result.append(f"{name}, {unit}, {round(price * discount, 2)}, {amount_left}\n")

    with open(filename_to_create, 'a+') as f:
        f.write(f"Nazwa, Sprzedawane w:, Cena, Ilość większa niż {low_amount}\n")
        for row in result:
            f.write(row)

    csv = pd.read_csv(filename_to_create)
    html = filename_to_create + ".html"
    csv.to_html(html)
    with open(html, 'r') as f:
        lines = f.readlines()
        lines[0] = '<table class="dataframe">\n'
        for index in range(len(lines) - 1):
            if "NIE" in lines[index]:
                lines[index] = '<td style="background-color:orange">NIE</td>\n'
    with open(html, 'w') as f:
        f.write("<style> * {font-size: 1.1em;} body {background-color: rgb(155, 155, 101);} table {margin-left: auto;margin-right: auto;border-color: black;border-spacing: 5px;}tr {text-align: center !important;padding: 15px;}th {padding: 15px;border: 3px solid black;margin: 10px 50px;}td {border: 3px solid black;padding: 15px 50px;margin: 10px 50px;}</style>\n")
        f.writelines(lines)        
    os.remove(filename_to_create)
    


create_offer()