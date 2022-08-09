from __future__ import annotations
import os
# enables using built-in type annotations rather than those from Typing module - added in python 3.9

from os.path import exists
import itertools
import pandas as pd

VALID_EXCEL_EXTENSIONS = [".xls", ".xlsx"]
ILLEGAL_CHARS = ['<', '>', ':', '"', '/', '\\', '|', '?', '*', '.']

INPUT_MESSAGES = {
    "START": ("Wpisz '1' jeśli chcesz stworzyć nową ofertę\n"
              "Wpisz '2' jeśli chcesz zmodyfikować istniejącą ofertę\n"),

    "GET_HTML_PATH": ("Wklej ścieżkę do pliku z ofertą, którą chcesz zmodyfikować\n"
                      "Lub samą nazwę pliku jeśli program i plik są w tym samym folderze\n"),

    "DELETING_OPTIONS": ("Wpisz '1' jeśli chcesz usunąć pozycje według liczby\n"
                         "Wpisz '2' jeśli chcesz usunąć pozycje według początku nazwy\n"
                         "Wpisz '3' jeśli chcesz usunąć pozycje według części całej nazwy\n"),

    "CHANGE_OPTIONS": ("Wpisz '1' jeśli chcesz usunąć pozycje\n"
                       "Wpisz '2' jeśli chcesz dodać pozycje\n"
                       "Wpisz '3' jeśli chcesz zmienić rabat\n"),

    "GET_EXCEL_PATH": ("Wklej ścieżkę do pliku excelowego, z którego chcesz wziąć dane\n"
                       "Lub samą nazwę pliku, jeśli plik jest w tym samym folderze co program.\n"),

    "KEYWORD_MATCH_TYPE": ("\nPodaj czy słowo kluczowe ma być szukane na początku nazwy produktu,"
                           " w całej nazwie, czy jako osobne słowo zawarte w nazwie\n"
                           "Lub wpisz stop żeby zakończyć dodawanie produktów\n"
                           "Cała nazwa -> Wpisz '1'\n"
                           "Początek -> Wpisz '2'\n"
                           "Osobne słowo -> Wpisz '3'\n"
                           "Zakończ dodawanie produktów -> 'stop'\n")

}

HTML_FILE_STRUCTURE = {
    "MAIN_STYLE": ("<style> * {font-size: 1.1em;}"
                   "body {background-color: rgb(155, 155, 101);}"
                   "table {margin-left: auto;margin-right: auto;border-color: black;border-spacing: 5px;}"
                   "tr {text-align: center !important;padding: 15px;}"
                   "th {padding: 15px;border: 3px solid black;margin: 10px 50px;}"
                   "td {border: 3px solid black;padding: 15px 50px;margin: 10px 50px;}"
                   "@keyframes alert {from {background-color: rgb(155, 155, 101);}to {background-color: red;}}"
                   "</style>\n"),

    "TABLE_HEAD": ("<table>\n"
                   "<thead>\n"
                   '<tr style="text-align: right;">'
                   "<th></th>\n"
                   "<th>Nazwa</th>\n"
                   "<th>Sprzedawane w:</th>\n"
                   "<th>Cena</th>\n"
                   "</tr>\n"
                   "</thead>\n"
                   "<tbody>\n"),

    "FILE_END": ("</tbody>\n"
                 "</table>")
}

WARNINGS = {
    "FILE_NOT_FOUND": "Ten plik nie istnieje, podaj poprawną ścieżkę lub nazwę pliku",
    "ILLEGAL_FILENAME": f'Podaj poprawną nazwę - nazwa nie może zawierać tych znaków: {", ".join(ILLEGAL_CHARS)}',
    "THIS_FILE_EXISTS": "Plik o tej nazwie już istnieje - do pliku dodana zostanie liczba oznaczająca wersję",
    "DISCOUNT_RANGE": "Rabat musi wynosić między 0 a 100",
    "MUST_BE_NUM": "Rabat musi być liczbą"
}

# ANSI escape sequences
PRINT_STYLES = {
    "WARNING": '\033[93m',
    "BOLD": '\033[1m'
}



EXTRACT_AND_COMPARE = {
    "name": lambda table_row, name: table_row.split("<td", maxsplit=1)[1]
            .split(">", maxsplit=1)[1]
            .split("<", maxsplit=1)[0].lower() == name.lower(),

    "ordinal_num": lambda table_row, num: table_row.split("<th>")[1].split("</th")[0] == num
}


SET = {
    "price": lambda table_row: table_row.rsplit("</td>", maxsplit=1)[0].rsplit(">", maxsplit=1)[1]
    # split from right as price is at the end of table row structure
}

DELETION = {
    "1": {
        "msg": "Podaj numer pozycji jaką chcesz usunąć, wpisz 'stop' żeby przestać edytować\n",
        "func": EXTRACT_AND_COMPARE["ordinal_num"]
        },
        
    "2": {
        "msg": "Podaj słowo kluczowe, wpisz 'stop' żeby przestać edytować\n",
        "func": EXTRACT_AND_COMPARE["name"]
        }
}



SEARCHING_TYPES = {
    "1": lambda keyword, name: keyword in name,
    "2": lambda keyword, name: name.startswith(keyword),
    "3": lambda keyword, name: keyword in name.split(" ")

}


def start() -> None:
    """
    Start of the program
    Asking user whether they want to create a new offer or modify an exisiting one
    """
    action = limited_options_input(INPUT_MESSAGES["START"], options=["1", "2"])
    MAIN_ACTIONS[action]()


def create_new_offer() -> None:
    excel_file = get_excel_source_filename()
    keywords, searching_list, discounts = keywords_searching_discounts()
    generate_offer_in_html(excel_file, keywords, discounts, searching_list)


def read_from_html_file(filename: str) -> tuple[list[str], list[str], list[str]]:
    with open(filename, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        file_start = list(itertools.takewhile(lambda x: "<tbody>" not in x, lines)) + ["<tbody>\n"]
        # returns lines before <tbody> (exclusive)
        file_end = list(itertools.dropwhile(lambda x: "</tbody>" not in x, lines))
        # returns lines after </tbody> (inclusive)
        table_contents = [table_row + "</tr>" for table_row in "".join(lines[len(file_start):-len(file_end)]).split("</tr>")][:-1]
        # returns contents of the html table and splits into rows - [:-1] is used
        # so there is no additional </tr> at the end
    return file_start, table_contents, file_end


def get_html_source_filename() -> str:
    while True:
        filename = input(INPUT_MESSAGES["GET_HTML_PATH"])
        if exists(filename) and filename.endswith(".html"):
            break
        if exists(filename + ".html"):
            filename += ".html"
            break
        warning(WARNINGS["FILE_NOT_FOUND"])

    return filename


def warning(message: str, styles: tuple[str] = (PRINT_STYLES["BOLD"], PRINT_STYLES["WARNING"])) -> None:
    os.system("color")
    end = '\033[0m'
    print(f'{"".join(styles)}{message}{end}')


def limited_options_input(input_message: str, *, options: list[str]) -> str:
    # using star operator so that options argument is a keyword argument
    """
    Function to make sure the user's input can be accounted for
    """
    while True:
        user_input = input(input_message).lower()
        if user_input in [option.lower() for option in options]:
            return user_input


def ask_for_input_in_loop(message: str, stop_word: str = "stop") -> list[str]:
    inputs = []
    while True:
        user_input = input(message).lower()
        if user_input == stop_word:
            return inputs
        inputs.append(user_input)


def delete_items(content: list[str], edit_type: str) -> list[str]:
    lines_to_write = []
    lines_counter, deleted_items_counter = 1, 0
    keywords = ask_for_input_in_loop(DELETION[edit_type]["msg"])

    for line in content:
        for keyword in keywords:
            if DELETION[edit_type]["func"](line, keyword):
                deleted_items_counter += 1
                break
        else:
            line = line.replace(f"<th>{lines_counter + deleted_items_counter}</th>", f"<th>{lines_counter}</th>")
            # this makes so that the correct order of ordinal numbers is preserved
            lines_counter += 1
            lines_to_write.append("".join(line))

    return lines_to_write


def delete_from_offer() -> None:
    file = get_html_source_filename()

    edit_type = limited_options_input(INPUT_MESSAGES["DELETING_OPTIONS"], options=["1", "2", "3"])
    file_start, table_contents, file_end = read_from_html_file(file)
    edited_table = delete_items(table_contents, edit_type)
    lines = file_start + edited_table + file_end

    with open(file, 'w', encoding='utf-8') as f:
        f.writelines(lines)


def change_offer() -> None:
    action = limited_options_input(INPUT_MESSAGES["CHANGE_OPTIONS"], options=["1", "2", "3"])
    CHANGE_ACTIONS[action]()


def change_discount() -> None:
    excel_file = get_excel_source_filename()
    html_file = get_html_source_filename()
    keywords, searching_list, discounts = keywords_searching_discounts()

    file_start, table_contents, file_end = read_from_html_file(html_file)
    new_items = items_list(excel_file, keywords, discounts, searching_list)

    for new_item in new_items:
        new_item_price, new_item_name = new_item[2], new_item[0].split("<br/>")[0]
        for index, item in enumerate(table_contents):
            if EXTRACT_AND_COMPARE["name"](item, new_item_name):
                item_price = SET["price"](item)
                table_contents[index] = item.replace(f"{item_price}", new_item_price)
                break

    file_contents = file_start + table_contents + file_end

    with open(html_file, 'w', encoding='utf-8') as f:
        f.writelines(file_contents)



def check_for_duplicates_and_write_to_file(new_items: list[list[str]]) -> None:
    html_source_file = get_html_source_filename()
    file_start, previous_content, file_end = read_from_html_file(html_source_file)
    new_items_list = []
    for new_item in new_items:
        new_item_name = new_item[0].split("<br/>")[0]
        for old_item in previous_content:
            if EXTRACT_AND_COMPARE["name"](old_item, new_item_name):
                break
        else:
            new_items_list.append(new_item)

    all_items_list = previous_content + generate_table_contents(new_items_list, counter=len(previous_content) + 1)

    file_contents = file_start + all_items_list + file_end

    with open(html_source_file, 'w', encoding='utf-8') as f:
        f.writelines(file_contents)
        

def add_to_offer() -> None:
    excel_source_file = get_excel_source_filename()
    keywords, searching_list, discounts = keywords_searching_discounts()
    new_items = items_list(excel_source_file, keywords, discounts, searching_list)
    check_for_duplicates_and_write_to_file(new_items)


def get_excel_source_filename() -> str:
    while True:
        source_filename = input(INPUT_MESSAGES["GET_EXCEL_PATH"])
        if exists(source_filename):
            break
                
        for excel_ext in VALID_EXCEL_EXTENSIONS:
            if exists(source_filename + excel_ext):
                return source_filename + excel_ext

        warning(WARNINGS["FILE_NOT_FOUND"])


def add_version(filename: str) -> str:
    """
    Adds version number to filename if it already exists
    """
    version = 1
    while True:
        if exists(filename + f"_{version}.html"):
            version += 1
        else:
            return filename + f"_{version}"


def create_html_file() -> str:
    while True:
        filename = input("Podaj nazwę pliku, jaki chcesz utworzyć\n").replace(".html", "")
        if any([char in filename for char in ILLEGAL_CHARS]):
            warning(WARNINGS["ILLEGAL_FILENAME"])
        if exists(filename + ".html"):
            warning(WARNINGS["THIS_FILE_EXISTS"])
            filename = add_version(filename)
            return filename + ".html"
        else:
            return filename + ".html"


def keywords_searching_discounts() -> tuple[list[str], list[str], list[float]]:
    keywords, searching_list, discounts = [], [], []
    while True:
        searching = limited_options_input(INPUT_MESSAGES["KEYWORD_MATCH_TYPE"], options=["1", "2", "3", "stop"])
        if searching == "stop":
            break
        searching_list.append(searching)
        keyword = input("Podaj słowo kluczowe do znalezienia produktów\n").lower()
        keywords.append(keyword)
        discount = get_discount_input()
        discounts.append(discount)

    return keywords, searching_list, discounts


def get_discount_input() -> float:
    while True:
        try:
            discount = float(input("Podaj rabat (od 0 do 100)\n"))
            if 0 <= discount <= 100:
                return 1 - (discount / 100)
            warning(WARNINGS["DISCOUNT_RANGE"])
        except ValueError:
            warning(WARNINGS["MUST_BE_NUM"])


def items_list(excel_file: str, keywords: list[str], discounts: list[float], searching: list[str]) -> list[list[str]]:
    items = []
    df = pd.read_excel(excel_file)
    columns = df.columns
    for name, amount_left, unit, price in zip(df[columns[0]], df[columns[1]], df[columns[2]], df[columns[3]]):
        amount_left = int(amount_left)
        for keyword, way_of_searching, discount in zip(keywords, searching, discounts):
            name = name.replace(",", '.')
            if SEARCHING_TYPES[way_of_searching](keyword, name.lower()):
                if amount_left == 0:
                    name += "<br/> CHWILOWO NIEDOSTĘPNE"
                item = [name, unit, str(round(price * discount, 2))]
                items.append(item)

    return items


def generate_offer_in_html(filename: str, keywords: list[str], discounts: list[float], searching: list[str]) -> None:
    lines = items_list(filename, keywords, discounts, searching)
    table_contents = generate_table_contents(lines)
    filename = create_html_file()

    with open(filename, 'a+', encoding='utf-8') as f:
        f.write(HTML_FILE_STRUCTURE["MAIN_STYLE"])
        f.write(HTML_FILE_STRUCTURE["TABLE_HEAD"])
        f.writelines(table_contents)
        f.write(HTML_FILE_STRUCTURE["FILE_END"])


def generate_table_contents(lines: list[list[str]], counter=1) -> list[str]:
    """
    Generates list of strings that are later written to file
    Option counter argument lets choose what is first ordinal number of new items
    """
    contents = []
    for item in lines:
        name, selling_unit, price = item[0], item[1], item[2]
        style = ('<td style=" animation-duration: 5s; animation-name:alert;'
                 f' animation-iteration-count: infinite;">{name}</td>\n') if "CHWILOWO NIEDOSTĘPNE" in name else f"<td>{name}</td>\n"
        table_row = ("\n<tr>\n"
                    f"<th>{counter}</th>\n"
                    f'{style}'
                    f"<td>{selling_unit}</td>\n"
                    f"<td>{price}</td>\n"
                    "</tr>\n")
        counter += 1
        contents.append(table_row)
    return contents


MAIN_ACTIONS = {
    "1": create_new_offer,
    "2": change_offer
}
# here actions can be added as function names

CHANGE_ACTIONS = {
    '1': delete_from_offer,
    '2': add_to_offer,
    '3': change_discount
}



if __name__ == "__main__":
    start()
