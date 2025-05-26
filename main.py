import pandas as pd
from itertools import product

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
from openpyxl.formula.translate import Translator

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

        s_config_cut = "+".join(map(str, config_cut)) # собираем все в одну строчку

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

        name = (f"{cat}, {mesh_size}, {color}, {weight} г/м2, {width}x{length} ({s_config_cut})")
        print(name)
        names.append(name)
    
    return names


def write_config_cut(max_width, width):
    var_widths = []
    for i in range (int(max_width // width)):
        var_widths.append(width)
    # print(max_width, width, max_width // width)
    extra = max_width % (width * (max_width // width))
    if extra != 0: var_widths.append(extra)
    return var_widths


def process_excel_file(input_file, output_file):

    df = load_excel_data(input_file)

    unique_combs = get_unique_combinations(df)

    generated_names = generate_names(unique_combs)

    # Загружаем существующую книгу
    book = load_workbook(input_file)

    # Создаем новый лист для результатов
    if 'Результаты' in book.sheetnames:
        del book['Результаты']

    result_sheet = book.copy_worksheet(book['Большие рулоны'])
    result_sheet.title = 'Результаты'

    # Сохраняем изменения
    book.save(output_file)
    
    # Вставляем названия в столбец "Название"
    insert_names_to_column(result_sheet, generated_names, column_name='Название')
    
    # result_sheet.delete_cols(find_col_index(result_sheet, "Категория"))
    # Вместо удаления просто сузим столбец
    # Получаем буквенное обозначение столбца
    column_letter = get_column_letter(2)
    
    # Устанавливаем ширину
    result_sheet.column_dimensions[column_letter].width = 0

    # Сохраняем изменения
    book.save(output_file)
    print(f"Результаты сохранены в {output_file}")


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