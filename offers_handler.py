from __future__ import annotations
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
                           "Początek -> Wpisz 'P'\n"
                           "Cała nazwa -> Wpisz 'C'\n"
                           "Osobne słowo -> Wpisz 'O'\n"
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


def start() -> None:
    """
    Start of the program
    Asking user whether they want to create a new offer or modify an exisiting one
    """
    action = limited_options_input(INPUT_MESSAGES["START"], options=["1", "2"])
    if action == "1":
        create_new_offer()

    elif action == "2":
        change_offer()


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
        table_contents = [x + "</tr>" for x in "".join(lines[len(file_start):-len(file_end)]).split("</tr>")][:-1]
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


def delete_items_by_ordinal_number(content: list[str]) -> list[str]:
    """
    User is asked for ordinal numbers of items from the table and then they are deleted
    I mean numbers not users :D
    """
    ordinal_numbers_list, lines_to_write = [], []
    lines_counter, deleted_items_counter = 1, 0
    while True:
        ordinal_number = input("Podaj numer pozycji jaką chcesz usunąć, wpisz 'stop' żeby przestać edytować\n").lower()
        if ordinal_number == "stop":
            break
        ordinal_numbers_list.append(ordinal_number)

    for line in content:
        line_number = line.split("<th>")[1].split("</th")[0]
        for num in ordinal_numbers_list:
            if num == line_number:
                ordinal_numbers_list.remove(num)
                deleted_items_counter += 1
                break
        else:        
            line = line.replace(f"<th>{lines_counter + deleted_items_counter}</th>", f"<th>{lines_counter}</th>")
            # this makes so that the correct order of ordinal numbers is preserved
            lines_counter += 1
            lines_to_write.append("".join(line))

    return lines_to_write


def delete_items_by_part_of_name(content: list[str], edit_type: str) -> list[str]:
    keywords, lines_to_write = [], []
    lines_counter, deleted_items_counter = 1, 0
    while True:
        keyword = input("Podaj słowo kluczowe, wpisz 'stop' żeby przestać edytować\n").lower()
        if keyword == "stop":
            break
        keywords.append(keyword)

    for line in content:
        item_name = line.split("<td", maxsplit=1)[1].split(">", maxsplit=1)[1].split("<", maxsplit=1)[0].lower()
        for keyword in keywords:
            edit_type_2_condition = item_name.startswith(keyword) and edit_type == "2"
            edit_type_3_condition = keyword in item_name and edit_type == "3"
            if edit_type_2_condition or edit_type_3_condition:
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

    if edit_type == '1':
        lines_to_write = delete_items_by_ordinal_number(table_contents)

    else:
        lines_to_write = delete_items_by_part_of_name(table_contents, edit_type)

    lines = file_start + lines_to_write + file_end

    with open(file, 'w', encoding='utf-8') as f:
        f.writelines(lines)


def change_offer() -> None:
    action = limited_options_input(INPUT_MESSAGES["CHANGE_OPTIONS"], options=["1", "2", "3"])

    if action == "1":
        delete_from_offer()

    if action == '2':
        add_to_offer()

    if action == "3":
        change_discount()


def change_discount() -> None:
    excel_file = get_excel_source_filename()
    html_file = get_html_source_filename()
    keywords, searching_list, discounts = keywords_searching_discounts()

    file_start, table_contents, file_end = read_from_html_file(html_file)
    new_items = items_list(excel_file, keywords, discounts, searching_list)

    for new_item in new_items:
        new_item_price, new_item_name = new_item[2], new_item[0].split("<br/>")[0]
        for index, item in enumerate(table_contents):
            item_name = item.split("<td", maxsplit=1)[1].split(">", maxsplit=1)[1].split("<", maxsplit=1)[0]
            if item_name == new_item_name:
                item_price = item.rsplit("</td>", maxsplit=1)[0].rsplit(">", maxsplit=1)[1]
                # split from right as price is at the end of table row structure
                table_contents[index] = item.replace(f"{item_price}", new_item_price)
                break

    with open(html_file, 'w', encoding='utf-8') as f:
        for line in file_start:
            f.write(line)
        for line in table_contents:
            f.write(line)
        for line in file_end:
            f.write(line)


def check_for_duplicates_and_write_to_file(new_items: list[list[str]]) -> None:
    html_source_file = get_html_source_filename()
    file_start, previous_content, file_end = read_from_html_file(html_source_file)
    new_items_list = []
    for new_item in new_items:
        new_item_name = new_item[0].split("<br/>")[0]
        is_copy = False
        for old_item in previous_content:
            old_item_name = old_item.split("<td", maxsplit=1)[1].split(">", maxsplit=1)[1].split("<", maxsplit=1)[0]
            if old_item_name == new_item_name:
                is_copy = True
                break
        if not is_copy:
            new_items_list.append(new_item)

    all_items_list = previous_content + generate_table_contents(new_items_list, counter=len(previous_content) + 1)

    with open(html_source_file, 'w', encoding='utf-8') as f:
        for line in file_start:
            f.write(line)
        for line in all_items_list:
            f.write(line)
        for line in file_end:
            f.write(line)


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

        for char in ILLEGAL_CHARS:
            if char in filename:
                warning(WARNINGS["ILLEGAL_FILENAME"])
                break
        else:
            break

    if exists(filename + ".html"):
        warning(WARNINGS["THIS_FILE_EXISTS"])
        filename = add_version(filename)

    return filename + ".html"


def keywords_searching_discounts() -> tuple[list[str], list[str], list[float]]:
    keywords, searching_list, discounts = [], [], []
    while True:
        searching = limited_options_input(INPUT_MESSAGES["KEYWORD_MATCH_TYPE"], options=["c", "p", "o", "stop"]).lower()
        if searching == "stop":
            break
        searching_list.append(searching)
        keyword = input("Podaj słowo kluczowe do znalezienia produktów\n")
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
            keyword = keyword.lower()
            name = name.replace(",", '.')
            if add_item(way_of_searching, keyword, name):
                if amount_left == 0:
                    name += "<br/> CHWILOWO NIEDOSTĘPNE"
                item = [name, unit, str(round(price * discount, 2))]
                items.append(item)

    return items


def generate_offer_in_html(filename: str, keywords: list[str], discounts: list[float], searching: list[str]) -> None:
    lines = items_list(filename, keywords, discounts, searching)
    filename = create_html_file()

    with open(filename, 'a+', encoding='utf-8') as f:
        f.write(HTML_FILE_STRUCTURE["MAIN_STYLE"])
        f.write(HTML_FILE_STRUCTURE["TABLE_HEAD"])
        for line in generate_table_contents(lines):
            f.write(line)
        f.write(HTML_FILE_STRUCTURE["FILE_END"])


def generate_table_contents(lines: list[list[str]], counter=1) -> list[str]:
    """
    Generates list of strings that are later written to file
    Option counter argument lets choose what is first ordinal number of new items
    """
    contents = []
    for item in lines:
        name, selling_unit, price = item[0], item[1], item[2]
        table_row_string = ""
        table_row_string += "\n<tr>\n"
        table_row_string += f"<th>{counter}</th>\n"
        if "CHWILOWO NIEDOSTĘPNE" in name:
            # adding style
            table_row_string += ('<td style=" animation-duration: 5s; animation-name:alert;'
                                 f' animation-iteration-count: infinite;">{name}</td>\n')
        else:
            table_row_string += f"<td>{name}</td>\n"
        table_row_string += f"<td>{selling_unit}</td>\n"
        table_row_string += f"<td>{price}</td>\n"
        table_row_string += "</tr>\n"
        counter += 1
        contents.append(table_row_string)
    return contents


def add_item(way_of_searching: str, keyword: str, name: str) -> bool:
    name = name.lower()
    if way_of_searching == 'c' and keyword in name:
        return True
    if way_of_searching == 'p' and name.startswith(keyword):
        return True
    if way_of_searching == 'o' and keyword in name.split(' '):
        return True

    return False


if __name__ == "__main__":
    start()
