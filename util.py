from dataclasses import dataclass
import re


def make_variable_name(name):
    name = name.strip()
    name = re.sub(r'([^ ])([A-Z][a-z]+)', r'\1_\2', name)
    name = re.sub(r'\s+', r'_', name)
    name = re.sub(r'\W', r'', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1_\2', name).lower()


def convert_excel_letters_to_number(letters):
    letters = letters.upper()
    accum = 0
    for l in letters:
        val = ord(l) - 64
        accum *= 26
        accum += val

    return accum


def convert_number_to_excel_letters(number):
    base = 26
    chars = []
    while number > base:
        num = number % base
        number = number // base
        # print(num, number, chr(num+64))
        chars.append(chr(num + 64))

    chars.append(chr(number + 64))
    chars.reverse()
    return "".join(chars)


def excel_to_coord(excel_coord):
    excel_coord.replace("$", "")
    match = re.match(r'([a-zA-Z]+)(\d+)', excel_coord)
    col = convert_excel_letters_to_number(match.group(1))
    row = int(match.group(2))
    return (col, row)


def excel_range_iter(range_value):
    match = re.match(r'(.+):(.+)', range_value)
    if not match:
        yield range_value.replace("$", "")
    else:
        start = match.group(1)
        stop = match.group(2)
        (start_col, start_row) = excel_to_coord(start)
        (stop_col, stop_row) = excel_to_coord(stop)
        for col in range(start_col, stop_col+1):
            for row in range(start_row, stop_row+1):
                yield convert_number_to_excel_letters(col) + str(row)


@dataclass(eq=True, frozen=True)
class CellInfo:
    coordinate: str
    variable_name: str
    value: any

    def is_formula(self):
        return str(self.value).startswith("=")


def get_alternating_cell_info(cell_range):
    cell_info = {}
    for i in range(0, len(cell_range), 2):
        header_cell = cell_range[i][0]
        input_cell = cell_range[i+1][0]

        cell_info[input_cell.coordinate] = CellInfo(
            coordinate=input_cell.coordinate,
            variable_name=make_variable_name(header_cell.value),
            value=input_cell.value
        )

    return cell_info


def get_horizontal_cell_info(header_range, target_range):
    cell_info = {}
    for idx, cell in enumerate(target_range):
        val_name = cell.coordinate
        if len(header_range) >= idx:
            header = header_range[idx].value
            if type(header) == str and header.strip() != "":
                val_name = header_range[idx].value

        cell_info[cell.coordinate] = CellInfo(
            coordinate=cell.coordinate,
            variable_name=make_variable_name(val_name),
            value=cell.value
        )

    return cell_info
