from this import s
import pandas as pd
import os
from os.path import exists
from datetime import datetime


def change_offer():
    while True:
        source_filename = input("Wklej ścieżkę do pliku z ofertą, którą chcesz zmodyfikować\n") + ".html"
        if not exists(source_filename):
            print("Ten plik nie istnieje, podaj poprawną ścieżkę")
        else:
            break

    while True:
        editing_type = input(
            "Wpisz '1' jeśli chcesz usunąć pozycje według liczby\nWpisz '2' jeśli chcesz usunąć pozycje według początku nazwy\nWpisz '3' jeśli chcesz usunąć pozycje według części całej nazwy\n")
        if editing_type in ['1', '2', '3']:
            break

    with open(source_filename, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        start = lines[:12]
        middle = [x + "</tr>" for x in "".join(lines[12:-2]).split("</tr>")]
        end = lines[-2:]
        lines_to_write = []
        counter, deleted_items_counter = 1, 0
        loop = True
        numbers, keywords = [], []

        for line in middle:
            if editing_type == '1':
                while loop:
                    number_to_delete = input(
                        "Podaj numer pozycji jaką chcesz usunąć, wpisz 'stop' żeby przestać edytować\n")
                    if number_to_delete == "stop":
                        loop = False
                        break
                    else:
                        numbers.append(number_to_delete)

                to_add = True
                for num in numbers:
                    if ">" + num + "<" in line:
                        to_add = False
                        numbers.remove(num)
                        deleted_items_counter += 1
                if to_add:
                    line = line.replace(f"<th>{counter + deleted_items_counter}</th>", f"<th>{counter}</th>")
                    counter += 1
                    lines_to_write.append("".join(line))

            if editing_type == "2":
                product_name = "".join(line.split("</td>", maxsplit=1)[0]).split("\n")[-1].replace("<td>",
                                                                                                   "").strip().upper()
                while loop:
                    keyword = input("Podaj słowo kluczowe, wpisz 'stop' żeby przestać edytować\n").upper()
                    if keyword == "STOP":
                        loop = False
                        break
                    else:
                        keywords.append(keyword)
                to_add = True
                for keyword in keywords:
                    if product_name.startswith(keyword):
                        to_add = False
                        deleted_items_counter += 1
                if to_add:
                    line = line.replace(f"<th>{counter + deleted_items_counter}</th>", f"<th>{counter}</th>")
                    counter += 1
                    lines_to_write.append("".join(line))
            if editing_type == "3":
                product_name = "".join(line.split("</td>", maxsplit=1)[0]).split("\n")[-1].replace("<td>",
                                                                                                   "").strip().upper()

                while loop:
                    keyword = input("Podaj słowo kluczowe, wpisz 'stop' żeby przestać edytować\n").upper()
                    if keyword == "STOP":
                        loop = False
                        break
                    else:
                        keywords.append(keyword)
                to_add = True
                for keyword in keywords:
                    if keyword in product_name:
                        to_add = False
                        deleted_items_counter += 1
                if to_add:
                    line = line.replace(f"<th>{counter + deleted_items_counter}</th>", f"<th>{counter}</th>")
                    counter += 1
                    lines_to_write.append("".join(line))

        lines = start + lines_to_write + end

    with open(source_filename, 'w', encoding='utf-8') as f:
        f.writelines(lines)


def get_info():
    while True:
        action = input(
            "Wpisz '1' jeśli chcesz stworzyć nową ofertę\nWpisz '2' jeśli chcesz zmodyfikować istniejącą ofertę\n")
        if action == "1":
            break
        elif action == "2":
            change_offer()
            exit()

    while True:
        print("lol")
        source_filename = input("Wklej ścieżkę do pliku, z którego chcesz wziąć dane!?!?!\n") + ".xlsx"
        print(source_filename)
        print(exists(source_filename))
        if exists(source_filename):
            break
        if exists(source_filename + ".xls"):
            source_filename += ".xls"
            break
        if exists(source_filename + ".xlsx"):
            source_filename += ".xlsx"
            break
        else:
            print("Ten plik nie istnieje, podaj poprawną ścieżkę")

    while True:
        filename_to_create = input("Podaj nazwę pliku, jaki chcesz utworzyć\n").split('.')[0]
        illegals_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in illegals_chars:
            if char in filename_to_create:
                print("Podaj poprawną nazwę pliku - nazwa nie może zawierać żadnego z tych znaków: " + ", ".join(
                    illegals_chars))
        else:
            break

    if exists(filename_to_create + ".html"):
        print("Plik o ten nazwie już istnieje - do pliku dodana zostanie data")
        filename_to_create += "-" + datetime.now().strftime("%d/%m/%Y %H:%M").replace("/", "_").replace(" ",
                                                                                                        "-").replace(
            ":", "_")

    keywords, way_of_searching_list, discounts, low_amounts = [], [], [], []

    while True:
        if input("Wpisz 'stop' jeśli chcesz zakończyć dodawanie produktów lub kliknij 'Enter' żeby kontynować\n") == "stop":
            break
        while True:
            way_of_searching = input("Podaj czy słowo kluczowe ma być szukane na początku nazwy produktu, w całej nazwie, czy jako osobne słowo zawarte w nazwie\nPoczątek -> Wpisz 'P'\nCała nazwa -> Wpisz 'C'\nOsobne słowo -> Wpisz 'O'\n").upper()
            if way_of_searching not in ["C", "P", "O"]:
                print("Wpisz literkę 'p' lub 'c'")
            else:
                break
        way_of_searching_list.append(way_of_searching)

        keyword = input("Podaj słowo kluczowe do znalezienia produktów\n")
        keywords.append(keyword)

        while True:
            try:
                discount = float(input("Podaj rabat (od 0 do 100)\n"))
                if not 0 <= discount <= 100:
                    print("Rabat musi wynosić między 0 a 100")
                else:
                    break

            except ValueError:
                print("Rabat musi być liczbą")

        discount = 1 - (discount / 100)
        discounts.append(discount)
        while True:
            try:
                low_amount = int(input("Podaj stan poniżej którego, można uznać że zostało niewiele tego produktu\n"))
                if low_amount <= 0:
                    print("Próg oznaczający małą ilość towaru musi być większy od 0")
                else:
                    break
            except ValueError:
                print("Stan musi być liczbą")

        low_amounts.append(low_amount)

    create_final_offer(source_filename, filename_to_create, keywords, discounts, low_amounts,
                       way_of_searching_list)


def create_final_offer(source_filename, filename_to_create, keywords, discounts, low_amounts,
                       look_for_match_in_whole_name_list):
    result = []
    df = pd.read_excel(source_filename)
    columns = df.columns
    for name, amount_left, unit, price in zip(df[columns[0]], df[columns[1]], df[columns[2]], df[columns[3]]):
        amount_left = int(amount_left)
        for keyword, low_amount, look_for_match_in_whole_name, discount in zip(keywords, low_amounts,
                                                                               look_for_match_in_whole_name_list,
                                                                               discounts):
            keyword = keyword.upper()
            if amount_left != 0:
                name = name.replace(",", '.')
                enough = "TAK" if amount_left > low_amount else "NIE"
                if look_for_match_in_whole_name == "C":
                    if keyword in name.upper():
                        result.append(f"{name}, {unit}, {round(price * discount, 2)}, więcej niż {low_amount}: {enough}\n")
                        break
                elif look_for_match_in_whole_name == "P":
                    if name.upper().startswith(keyword):
                        result.append(f"{name}, {unit}, {round(price * discount, 2)}, więcej niż {low_amount}: {enough}\n")
                        break
                else:
                    if keyword in name.split(" "):
                        result.append(f"{name}, {unit}, {round(price * discount, 2)}, więcej niż {low_amount}: {enough}\n")
                        break

    with open(filename_to_create, 'a+', encoding='utf-8') as f:
        f.write(f"Nazwa, Sprzedawane w:, Cena, Znaczna ilość\n")
        for row in result:
            f.write(row)

    csv = pd.read_csv(filename_to_create)
    csv.index += 1
    html = filename_to_create + ".html"
    csv.to_html(html)
    with open(html, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        lines[0] = '<table class="dataframe">\n'
        for index in range(len(lines) - 1):
            if "NIE" in lines[index]:
                lines[index] = lines[index].replace('<td>', '<td style="background-color:orange">') + "\n"
    with open(html, 'w', encoding='utf-8') as f:
        f.write("<style> * {font-size: 1.1em;} body {background-color: rgb(155, 155, 101);} table {margin-left: auto;margin-right: auto;border-color: black;border-spacing: 5px;}tr {text-align: center !important;padding: 15px;}th {padding: 15px;border: 3px solid black;margin: 10px 50px;}td {border: 3px solid black;padding: 15px 50px;margin: 10px 50px;}</style>\n")
        f.writelines(lines)
    os.remove(filename_to_create)


if __name__ == "__main__":
    get_info()
