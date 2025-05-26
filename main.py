import pandas as pd
from itertools import product

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
from openpyxl.formula.translate import Translator

color_art = {
    "Черный": "BLK",
    "Синий": "BLU",
    "Коричневый": "BRN",
    "Зеленый": "GRN",
    "Серый": "GRY",
    "Любой": "NCA",
    "Оранжевый": "ORG",
    "Розовый": "PNK",
    "Пурпурный": "PPL",
    "Красный": "RED",
    "Золотистый": "TAN",
    "Фиолетовый": "VIO",
    "Белый": "WHT",
    "Желтый": "YEL",
    "Прозрачный": "TRN",
    "Серебристый": "SLV",
    "Хаки": "KHK"
}

def load_excel_data(file_path):
    return pd.read_excel(file_path, sheet_name="Большие рулоны", header=1)

def get_unique_combinations(df):
    required_columns = ['Категория', 'Размер ячейки (мм)', 'Цвет', 'Вес (г/м2)', 'Ширина рулона (м)', 'Длина полезная ']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"В таблице отсутствует колонка: {col}")

    unique_combs = df[required_columns].drop_duplicates()
    return unique_combs.to_dict('records')

def generate_names(combinations):
    names = []
    articules = []
    cat_5 = [] # тут мы храним категории сеток, которые идут в максимальной ширине 5м
    for comb in combinations:
        cat = comb['Категория']
        width = comb['Ширина рулона (м)']
        if (width == 5): 
            cat_5.append(cat)
    for comb in combinations:
        cat = comb['Категория']
        mesh_size = comb['Размер ячейки (мм)']
        color = comb['Цвет']
        weight = comb['Вес (г/м2)']
        width = comb['Ширина рулона (м)']
        length = comb['Длина полезная ']

        # записываем вариант нарезки с текущей шириной при максимально возможной 4м
        config_cut = []
        if not cat in cat_5:
            config_cut = write_config_cut(4, width)
        else: 
            config_cut = write_config_cut(5, width)

        for i in range(len(config_cut)): # убираем лишние нули в записи
            if config_cut[i] % 1 == 0: config_cut[i] = int(config_cut[i])
            config_cut[i] = str(config_cut[i])

        for i in range(len(config_cut)): # убираем лишние нули в записи, точку меняем на запятую
            if '.' in config_cut[i]: config_cut[i] = f"{float(config_cut[i]):.1f}".replace(".", ",")

        if (len(config_cut) > 4): s_config_cut = f"{config_cut[0]}x{len(config_cut)}" # чтобы писалось 0,5x8 например
        else: s_config_cut = "+".join(map(str, config_cut)) # собираем все в одну строчку через +

        mesh_size = str(mesh_size).replace(" ", "") # затираем нули в размере ячейки

        # убираем в записи незначащие нули
        if weight % 1 == 0: weight = int(weight)
        # if width % 1 == 0: width = int(width)
        if length % 1 == 0: length = int(length)

        # точку меняем на запятую
        # width = str(width)
        length = str(length)
        # if '.' in width: width = f"{float(width):.1f}".replace(".", ",")
        if '.' in length: length = f"{float(length):.1f}".replace(".", ",")

        if (cat in cat_5): const_width = "5" # ширина в артикуле 5
        else: const_width = "4" # ширина в артикуле 4

        # добавляем имя
        name = (f"Сетка полимерная, {mesh_size}, {color}, {weight} г/м2, {const_width}x{length} ({s_config_cut})")
        print(name)
        names.append(name)
        # добавляем артикул
        if (cat in cat_5): art_width = "5" # ширина в артикуле 5
        else: art_width = "4" # ширина в артикуле 4
        art = generate_arts(color, mesh_size, weight, art_width, length, config_cut)
        articules.append(art)
    
    return [names, articules]


def generate_arts(color, mesh_size, weight, art_width, length, config_cut):
    if len(config_cut) > 4:  s_config_cut = f"{ str(config_cut[0]) }".replace(",", "")
    else: 
        s_config_cut = ""
        for i in config_cut: s_config_cut += str(i).replace(",", "")
    
    return f"{color_art[color]}_{mesh_size}_{weight}_{art_width}_{length}_{s_config_cut}"


def generate_names_2(combinations):
    names = []
    articules = []
    cat_5 = [] # тут мы храним категории сеток, которые идут в максимальной ширине 5м
    for comb in combinations:
        cat = comb['Категория']
        width = comb['Ширина рулона (м)']
        if (width == 5): 
            cat_5.append(cat)
    for comb in combinations:
        cat = comb['Категория']
        mesh_size = comb['Размер ячейки (мм)']
        color = comb['Цвет']
        weight = comb['Вес (г/м2)']
        width = comb['Ширина рулона (м)']
        length = comb['Длина полезная ']

        # записываем вариант нарезки с текущей шириной при максимально возможной 4м
        config_cut = []
        if not cat in cat_5:
            config_cut = write_config_cut(4, width)
        else: 
            config_cut = write_config_cut(5, width)

        for i in range(len(config_cut)): # убираем лишние нули в записи
            if config_cut[i] % 1 == 0: config_cut[i] = int(config_cut[i])
            config_cut[i] = str(config_cut[i])

        for i in range(len(config_cut)): # убираем лишние нули в записи, точку меняем на запятую
            if '.' in config_cut[i]: config_cut[i] = f"{float(config_cut[i]):.1f}".replace(".", ",")

        if (len(config_cut) > 4): s_config_cut = f"{config_cut[0]}x{len(config_cut)}" # чтобы писалось 0,5x8 например
        else: s_config_cut = "+".join(map(str, config_cut)) # собираем все в одну строчку через +

        mesh_size = str(mesh_size).replace(" ", "") # затираем нули в размере ячейки

        # убираем в записи незначащие нули
        if weight % 1 == 0: weight = int(weight)
        if width % 1 == 0: width = int(width)
        if length % 1 == 0: length = int(length)

        # точку меняем на запятую
        width = str(width)
        length = str(length)
        if '.' in width: width = f"{float(width):.1f}".replace(".", ",")
        if '.' in length: length = f"{float(length):.1f}".replace(".", ",")

        # добавляем имя
        name = (f"Сетка полимерная, {mesh_size}, {color}, {weight} г/м2, {width}x{length}м")
        names.append(name)
        # добавляем артикул
        if (cat in cat_5): art_width = "5" # ширина в артикуле 5
        else: art_width = "4" # ширина в артикуле 4
        art = generate_arts_2(color, mesh_size, weight, width, length)
        articules.append(art)
    
    return [names, articules]


def generate_arts_2(color, mesh_size, weight, width, length):
    return f"{color_art[color]}_{mesh_size}_{weight}_{str(width).replace(",", "")}_{length}"


def write_config_cut(max_width, width):
    var_widths = []
    for i in range (int(max_width // width)):
        var_widths.append(width)
    # print(max_width, width, max_width // width)
    extra = max_width % (width * (max_width // width))
    if extra != 0: var_widths.append(extra)
    return list(reversed(var_widths))


def process_excel_file(input_file, output_file):

    df = load_excel_data(input_file)

    unique_combs = get_unique_combinations(df)

    [generated_names, articules] = generate_names(unique_combs)

    # Загружаем существующую книгу
    book = load_workbook(input_file)


    # ЛИСТ СО СГЕНЕРИРОВАННЫМИ АРТИКУЛАМИ
    if 'Результаты' in book.sheetnames:
        del book['Результаты']

    result_sheet = book.copy_worksheet(book['Большие рулоны'])
    result_sheet.title = 'Результаты'

    # Сохраняем изменения
    book.save(output_file)
    # Вставляем названия в столбец "Название"
    insert_names_to_column(result_sheet, generated_names, column_name='Название')
    # Вставляем артикулы в столбец "Артикул"
    insert_names_to_column(result_sheet, articules, column_name='Артикул')
    
    # result_sheet.delete_cols(find_col_index(result_sheet, "Категория"))
    # Можно удалить, но все сломается, поэтому сузим столбец
    # Получаем буквенное обозначение столбца и меняем ширину
    column_letter = get_column_letter(2)
    result_sheet.column_dimensions[column_letter].width = 0
    column_letter = get_column_letter(1)
    result_sheet.column_dimensions[column_letter].width = 42
    column_letter = get_column_letter(3)
    result_sheet.column_dimensions[column_letter].width = 45

    # Сохраняем изменения
    book.save(output_file)


    # ЛИСТ С НАРЕЗАННЫМИ РУЛОНАМИ
    if 'Нарезанные рулоны' in book.sheetnames:
        del book['Нарезанные рулоны']

    result_sheet = book.copy_worksheet(book['Большие рулоны'])
    result_sheet.title = 'Нарезанные рулоны'

    [generated_names, articules] = generate_names_2(unique_combs)

    # Сохраняем изменения
    book.save(output_file)
    # Вставляем названия в столбец "Название"
    insert_names_to_column(result_sheet, generated_names, column_name='Название')
    # Вставляем артикулы в столбец "Артикул"
    insert_names_to_column(result_sheet, articules, column_name='Артикул')
    
    # result_sheet.delete_cols(find_col_index(result_sheet, "Категория"))
    # Можно удалить, но все сломается, поэтому сузим столбец
    # Получаем буквенное обозначение столбца и меняем ширину
    column_letter = get_column_letter(2)
    result_sheet.column_dimensions[column_letter].width = 0
    column_letter = get_column_letter(1)
    result_sheet.column_dimensions[column_letter].width = 42
    column_letter = get_column_letter(3)
    result_sheet.column_dimensions[column_letter].width = 45

    # Сохраняем изменения
    book.save(output_file)

    print(f"Все сохранено в {output_file}")


def insert_names_to_column(sheet, names, column_name='Название'):
    """Вставляет названия в указанный столбец"""
    # Находим индекс столбца
    col_idx = None
    for cell in sheet[2]:  # Ищем в заголовках
        if cell.value and str(cell.value).lower() == column_name.lower():
            col_idx = cell.column
            break
    
    if col_idx is None:
        raise ValueError(f"Столбец '{column_name}' не найден")
    
    # Вставляем данные (начиная со второй строки)
    for i, name in enumerate(names, start=3):
        sheet.cell(row=i, column=col_idx, value=name)


def find_col_index(sheet, col_name):
    # Находим индекс столбца
    col_id = None
    for cell in sheet[2]:  # Ищем в заголовках
        if cell.value and str(cell.value).lower() == col_name.lower():
            col_id = cell.column
            break
    
    if col_id is None:
        raise ValueError(f"Столбец '{col_name}' не найден")
    
    return col_id


if __name__ == "__main__":
    input_excel = "data.xlsx"
    output_excel = "result.xlsx"
    
    process_excel_file(input_excel, input_excel)